"""outlook2obsidian export command."""

import copy
import hashlib
import json
import logging
import os
import re
import sys
import unicodedata
from collections.abc import Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal
from enum import StrEnum
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any

import requests
import typer
import yaml
from html_to_markdown import ConversionOptions
from html_to_markdown import convert as convert_to_markdown
from requests import sessions
from rich.progress import BarColumn, Progress, TaskID, TaskProgressColumn

from outlook2obsidian.shared import (
    APP_NAME,
    AUTH_ERROR_USECOMMAND,
    GRAPH_API_BATCH_MAX,
    GRAPH_API_BATCH_MIN,
    GRAPH_API_URL,
    GRAPH_MESSAGE_FIELDS_SET,
    MAILBOX_ROOT,
    AuthRequiredError,
    CommandContext,
    GraphApiConfig,
    GraphApiReturnT,
    GraphFolderT,
    GraphMessageT,
    NoteMetaPropT,
    console,
    err_console,
    get_graph_config,
    graph_get_folder_by_path,
    graph_get_resource,
)

if TYPE_CHECKING:
    from collections.abc import Generator

app = typer.Typer(add_completion=False)
logger = logging.getLogger(APP_NAME)  # initialized in __main__


######################################################################
# Constants

TITLE_UNTITLED = "Untitled"
TITLE_MAXLEN = 255  # macOS limit
MIN_PROGRESS_THRESHOLD = 10  # Hide progress if only a few messages

# These are banned outright: \/:
# Links do not work if title includes: #^[]|
# macOS is fine with: *"<>?
# Title cannot start with a period.
# \x00-\x1f\x7f = [:cntrl:]
RE_TITLE_IGNORECHARS = re.compile(r"[\[\]\/\\:|^#\x00-\x1f\x7f]")
RE_TITLE_WHITESPACE = re.compile(r"\s+")

# Simple detection if filter/search critera contain 'and' or 'or' conditions.
RE_AND_OR = re.compile(r"\s(and|or)\s", re.IGNORECASE)

# Message fields we reference for exporting
EXPORT_MESSAGE_FIELDS = [
    "body",
    "createdDateTime",
    "id",
    "isDraft",
    "lastModifiedDateTime",
    "receivedDateTime",
    "sentDateTime",
    "subject",
    "hasAttachments",
]

EXPECTED_MESSAGE_FIELDS_SET = GRAPH_MESSAGE_FIELDS_SET | {"@odata.etag"}
MESSAGE_DATE_FIELDS = ["createdDateTime", "lastModifiedDateTime", "sentDateTime", "receivedDateTime"]


######################################################################
# Configure YAML to understand Decimal type


def _yaml_represent_decimal(dumper: yaml.dumper.SafeDumper, data: Decimal) -> yaml.ScalarNode:
    return dumper.represent_scalar("tag:yaml.org,2002:float", str(data))


yaml.SafeDumper.add_representer(Decimal, _yaml_represent_decimal)


######################################################################
# Types and classes

FrontmatterT = dict[str, str | int | Decimal | bool | Sequence[str | int]]


class FilterJoinCond(StrEnum):
    """Choices for joining filter conditions."""

    cond_and = "and"
    cond_or = "or"


class ExportStat(StrEnum):
    """Export statistic messages."""

    STAT_NOTE_CREATED = "Notes created"
    STAT_SKIPPED_DUPLICATE = "Skipped; Message already exported"
    STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED = "Skipped; Markdown conversion failed"
    STAT_SKIPPED_NO_MESSAGE_BODY = "Skipped; No message body"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS = "Skipped; No message content, just attachments"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS = "Skipped; No message content, no attachments"


@dataclass
class ExportOptions:
    """Export command options."""

    dry_run: bool  # Whether to perform a dry run without making any changes.
    filter_query: str | None  # --filter criteria from command line and config file.
    filter_test: bool  # Whether to perform a filter test without exporting.
    folder_mkdir: bool  # Whether to create the vault directory if it does not exist.
    graph_config: GraphApiConfig  # Graph API configuration.
    note_body_tags: list[str] | None  # --body-tag tags from command line and config file.
    note_metaprop: NoteMetaPropT | None  # --metaprop properties from command line and config file.
    note_tags: list[str] | None  # --tag tags from command line and config file.
    place_tags_body_end: bool  # True to place body tags at the end of the note instead the top.
    place_tags_frontmatter: bool  # True to place non-body tags in frontmatter.
    search_query: str | None  # --search criteria from command line and config file.
    vault_dir: Path  # Vault directory to save exported messages.


ExportStatsT = dict[ExportStat, int]

######################################################################
# Helpers


@contextmanager
def _logger_friendly_progress(*args, **kwargs) -> Generator[Progress]:  # noqa: ANN002, ANN003
    """Wrapper around rich.progress.Progress that redirects logger output through 'progress.console'."""
    # Emiting messages with standard Python logging breaks the Progress bar because logger handlers
    # uses the system stderr and Rich redirects stdout and stderr (by default but configurable) to the "console".
    # You can also use 'console.print', 'progress.print' and 'progress.log'.
    #
    # Logging breaks progress bar
    # https://github.com/Textualize/rich/issues/4235
    #
    # This context manager solution is modeled after this discussion comment:
    # https://github.com/Textualize/rich/discussions/1578#discussioncomment-5565014

    __logger = kwargs.pop("logger", logging.getLogger())
    __streamhandlers = [(handler, None) for handler in __logger.root.handlers if type(handler) is logging.StreamHandler]

    try:
        with Progress(*args, **kwargs) as progress:
            for i, (handler, _) in enumerate(__streamhandlers):
                prior_stderr = handler.setStream(sys.stderr)
                __streamhandlers[i] = handler, prior_stderr

            yield progress

    finally:
        for handler, prior_stderr in __streamhandlers:
            if prior_stderr is not None:
                handler.setStream(prior_stderr)


def _exportstat_incr(export_stats: ExportStatsT, stat: ExportStat, incr: int = 1) -> None:
    """Increment specified export stat.

    Args:
        export_stats (ExportStatsT): Export stats dictionary.
        stat (ExportStat): Stat to increment.
        incr (int, optional): Increment amount. Defaults to 1.
    """
    export_stats[stat] = export_stats.get(stat, 0) + incr


def _opt_metaprop_callback(
    ctx: typer.Context,
    param: typer.CallbackParam,  # noqa: ARG001
    meta: list[str] | None,
) -> list[str] | None:
    """Handle `--metaprop` command option.

    Note:
        The values of `default_map` are used as the default values for command options.
        This function is called even if `--config-file` is not specified.
        Duplicate keys are merged into a list of values

    Args:
        ctx (typer.Context): Typer context object.
        param (typer.CallbackParam): Typer callback parameter object.
        meta (list[str] | None): Command line option value or None if option is not specified.

    Returns:
        list[str] | None: List of parsed key-value pairs or None if no properties were provided.
    """
    if meta:
        ctx.ensure_object(CommandContext)
        ctx_obj: CommandContext = ctx.obj
        ctx_obj.export_metaprops = {}

        for value in meta:
            try:
                k, v = value.split("=", 1)
            except ValueError as err:
                msg = f"Format must be 'name=value', got: '{value}'"
                raise typer.BadParameter(msg) from err
            else:
                v_k = ctx_obj.export_metaprops.get(k)
                if v_k is not None:
                    # Duplicate keys are merged into a list of values
                    v_list = [v_k] if isinstance(v_k, (str, int)) else [*v_k]
                    v = [*v_list, v]

                ctx_obj.export_metaprops[k] = v

    return meta


def _merge_metaprops(a: NoteMetaPropT, b: NoteMetaPropT | None) -> NoteMetaPropT:
    """Combines two dictionaries, merging duplicate keys into a list of values.

    Args:
        a (NoteMetaPropT): Dictionary.
        b (NoteMetaPropT | None): Dictionary to merge into `a`. If None, a copy of `a` is returned.

    Returns:
        NoteMetaPropT
    """
    c = a.copy()

    if b:
        common_keys = a.keys() & b.keys()
        unique_to_b = b.keys() - a.keys()

        # Add b to a
        for key in unique_to_b:
            c[key] = b[key]

        # Merge duplicate keys into a list of values
        for key in common_keys:
            v_c, v_b = c[key], b[key]
            v_a_list = [v_c] if isinstance(v_c, (str, int)) else [*v_c]
            v_b_list = [v_b] if isinstance(v_b, (str, int)) else [*v_b]
            v_a_list.extend(v_b_list)
            c[key] = v_a_list

    return c


def _log_message_notice(message: GraphMessageT, errmsg: str, level: int = logging.WARNING) -> None:
    """Log a warning related to a message.

    Args:
        message (GraphMessageT): Message object.
        errmsg (str): Warning description.
        level (int): Logging level code (e.g., logging.WARNING).
    """
    logger.log(level, 'message("%s"): %s', message["subject"], errmsg)


def _safe_title(subject: str) -> str:
    """Clean up title so it is safe to be used as a file name.

    Args:
        subject (str): Message subject.

    Returns:
        str
    """
    subject = RE_TITLE_IGNORECHARS.sub("", subject)
    subject = RE_TITLE_WHITESPACE.sub(" ", subject).strip().lstrip(".")

    if len(subject) > TITLE_MAXLEN:
        subject = subject[:TITLE_MAXLEN].rsplit(" ", 1)[0]

    return subject or TITLE_UNTITLED


def _unique_out_path(p: Path) -> Path:
    """Create a Path with unique file name.

    Args:
        p (Path): Existing path.

    Raises:
        ValueError: Raised if too many copies of the file already exist.

    Returns:
        Path: Unique path with a non-conflicting file name.
    """
    n = 1
    max_paths = 100
    unique_path = p

    while unique_path.exists() and n < max_paths:
        unique_path = p.with_stem(f"{p.stem} ({n})")
        n += 1

    if n == max_paths:
        errmsg = f"{n} copies of '{p}' already exist?!"
        raise ValueError(errmsg)

    return unique_path


def _coerce_datetime(message: GraphMessageT, field: str, is_draft: bool | None) -> datetime | None:  # noqa: FBT001
    """Convert a message datetime string in ISO format into a `datetime` object.

    Args:
        message (GraphMessageT): Message object.
        field (str): Message datetime field.
        is_draft (bool): True if message is marked as a draft.

    Returns:
        datetime | None: Datetime object if field exists and conversion is successful, otherwise None.
    """
    dt = message.get(field)

    if field not in message:
        _log_message_notice(message, f"field({field}): missing date field")
    elif not dt:
        if is_draft and field in {"sentDateTime", "receivedDateTime"}:
            pass  # Drafts aren't expected to have sent or received dates
        else:
            _log_message_notice(message, f"field({field}): no value")
    else:
        try:
            return datetime.fromisoformat(dt)
        except ValueError:
            _log_message_notice(message, f"field({field}): unexpected format: '{dt}'")

    return None


def _clean_tags(tags: list[str]) -> list[str]:
    """Normalize tags to follow rules Obsidian defines.

    https://obsidian.md/help/tags

    Args:
        tags (list[str]): List of tags.

    Returns:
        list[str]: Updated list of tags.
    """
    for i, _tag in enumerate(tags):
        tag = _tag.strip()

        if " " in tag:
            tag = tag.title().replace(" ", "")

        tag = "".join(
            # Thanks Gemini
            char
            for char in tag
            if unicodedata.category(char)[0] in "LNMS" or char in "_-/\u200d"
        )

        if tag.isdecimal():
            tag = f"_{tag}"

        tags[i] = tag

    return tags


def _expand_folder_tags(
    folder: GraphFolderT,
    tags: list[str] | None,
    tag_vars: dict[str, str],
) -> list[str] | None:
    """Expand folder related tags.

    Args:
        folder (GraphFolderT): Folder object.
        tags (list[str] | None): List of tags to expand.
        tag_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        list[str] | None: Updated list of tags with folder-related placeholders expanded.
    """
    if not tags:
        return None

    # Create a pattern that searches for all replacement substrings simultaneously in a single pass
    pattern = re.compile("|".join(re.escape(key) for key in tag_vars))

    # Replace placeholders, do not modify original
    tags = [pattern.sub(lambda match: tag_vars[match.group(0)], _tag) for _tag in tags]

    add_tags: list[str] = []
    remove_tags: set[int] = set()

    # Expand placeholders
    for i, tag in enumerate(tags):
        if "{folder.folderPath.parts}" in tag:
            add_tags.extend(part for part in Path(folder["folderPath"]).parts)
            remove_tags.add(i)  # remove placeholder from list

    if remove_tags:
        tags = [tag for i, tag in enumerate(tags) if i not in remove_tags]
    if add_tags:
        tags.extend(add_tags)

    return tags


def _expand_message_tags(
    tags: list[str] | None,
    tag_vars: dict[str, str],
) -> list[str]:
    """Expand tags for a message.

    Args:
        tags (list[str] | None): List of tags to expand.
        tag_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        list[str]: Updated list of tags with message-related placeholders expanded.
    """
    if not tags:
        return []

    # Create a pattern that searches for all replacement substrings simultaneously in a single pass
    pattern = re.compile("|".join(re.escape(key) for key in tag_vars))

    # Replace placeholders, do not modify original
    return [pattern.sub(lambda match: tag_vars[match.group(0)], _tag) for _tag in tags]


def _expand_folder_metaprop(
    folder: GraphFolderT,
    metaprop: NoteMetaPropT | None,
    prop_vars: dict[str, str],
) -> NoteMetaPropT | None:
    """Expand folder related metadata properties.

    Args:
        folder (GraphFolderT): Folder object.
        metaprop (MetaPropT | None): Dictionary of metadata to expand.
        prop_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        MetaPropT | None: Updated dictionary of metadata with folder-related placeholders expanded.

    Raises:
        ValueError
    """
    if not metaprop:
        return None

    # Create a pattern that searches for all replacement substrings simultaneously in a single pass
    pattern = re.compile("|".join(re.escape(key) for key in prop_vars))

    # Replace placeholders, do not modify original
    for name, value in metaprop.items():
        if isinstance(value, str):
            metaprop[name] = pattern.sub(lambda match: prop_vars[match.group(0)], value)

            if "{folder.folderPath.parts}" in value:
                # Frontmatter can include list types so we can simply pass it along
                path_parts = Path(folder["folderPath"]).parts
                metaprop[name] = path_parts[0] if len(path_parts) == 1 else path_parts

    return metaprop


def _expand_note_metaprop(metaprop: NoteMetaPropT | None, prop_vars: dict[str, str], frontmatter: FrontmatterT) -> None:
    """Expand metadata properties for a note and add them to frontmatter.

    Args:
        metaprop (MetaPropT | None): Dictionary of metadata properties to expand.
        prop_vars (dict[str, str]): Dictionary of replacement patterns.
        frontmatter (FrontmatterT): Frontmatter dictionary. Frontmatter properties will be added here if applicable.
    """

    def is_float(value: str) -> bool:
        try:
            float(value)
        except ValueError:
            return False
        return True

    if metaprop:
        # Create a pattern that searches for all replacement substrings simultaneously in a single pass
        pattern = re.compile("|".join(re.escape(key) for key in prop_vars))

        for key, value in metaprop.items():
            if isinstance(value, str):
                # Replace placeholders
                new_value = pattern.sub(lambda match: prop_vars[match.group(0)], value)

                # Properties support some scalar types
                if new_value.isdecimal():
                    frontmatter[key] = int(new_value)
                elif "." in value and is_float(new_value):
                    frontmatter[key] = Decimal(new_value)
                elif new_value.lower() in ("true", "false"):
                    frontmatter[key] = new_value.lower() == "true"
                else:
                    frontmatter[key] = new_value
            else:
                frontmatter[key] = value


def _append_note_tags(
    note_content: str,
    export_opts: ExportOptions,
    tag_vars: dict[str, str],
    frontmatter: FrontmatterT,
) -> str:
    """Append tags to note content.

    Args:
        note_content (str): Note content.
        export_opts (ExportOptions): Export options.
        frontmatter (FrontmatterT): Frontmatter dictionary. Frontmatter tags will be added here if applicable.
        tag_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        str
    """
    body_tags = _expand_message_tags(export_opts.note_body_tags, tag_vars)
    frontmatter_tags = _expand_message_tags(export_opts.note_tags, tag_vars)

    if not export_opts.place_tags_frontmatter:
        # Place all tags in body
        body_tags.extend(frontmatter_tags)
        frontmatter_tags = []

    # Format as valid tags
    body_tags = _clean_tags(body_tags)
    frontmatter_tags = _clean_tags(frontmatter_tags)

    # Remove duplicates while keeping order
    body_tags = list(dict.fromkeys(body_tags))
    frontmatter_tags = list(dict.fromkeys(frontmatter_tags))

    if frontmatter_tags:
        frontmatter["tags"] = frontmatter_tags  # modifies caller

    if body_tags:
        tag_block = " ".join(f"#{tag}" for tag in body_tags)

        if export_opts.place_tags_body_end:
            note_content += "\n\n" + tag_block
        else:
            note_content = tag_block + "\n\n" + note_content

    return note_content


def _convert_to_markdown(body_content: str) -> str | None:
    """Convert message content to Markdown.

    Args:
        body_content (str): The HTML body content of the message.

    Returns:
        str|None: Note content or None if message has no body or Markdown conversion fails.
    """
    # https://github.com/matthewwithanm/python-markdownify
    # from markdownify import markdownify as html_to_md
    # note_content = html_to_md(body_content)

    # https://github.com/xberg-io/html-to-markdown
    # https://pypi.org/project/html-to-markdown/
    # https://docs.html-to-markdown.xberg.io/reference/api-python/

    # Disable metadata extraction to prevent Microsoft Word custom <meta> tags from being added to the frontmatter.
    # I couldn't find a way to parse it and suppress the frontmatter altogether :/
    # Everything is escaped presuming emails aren't meant to include Markdown entities, such as hashtags.
    opts = ConversionOptions(
        extract_metadata=False,
        escape_ascii=True,
        escape_asterisks=True,
        escape_misc=True,
        escape_underscores=True,
    )

    convert_result = convert_to_markdown(body_content, opts)

    return convert_result.content


def _export_message(  # noqa: C901, PLR0912, PLR0915
    message: GraphMessageT,
    vault_dir: Path,
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> None:
    """Export a message.

    Args:
        message (MessageT): The message.
        vault_dir (Path): The directory to export the message to.
        export_opts (ExportOptions): Export options.
        export_stats (ExportStatsDict): Export statistics dictionary to update.
    """
    subject: str = message.get("subject", TITLE_UNTITLED)
    title = _safe_title(subject)
    is_draft: bool | None = message.get("isDraft")
    body_content: str | None = message.get("body", {}).get("content")
    content_type: str | None = message.get("body", {}).get("contentType")
    has_attachments: bool | None = message.get("hasAttachments")
    message_out_path = _unique_out_path(vault_dir / Path(title + ".md"))
    frontmatter: FrontmatterT = {}
    dt_now = datetime.now(tz=timezone.utc)  # noqa: UP017
    note_content = body_content

    # Convert date values into datetime objects
    message_dates = {
        field: _coerce_datetime(message, field, is_draft)
        for field in ("createdDateTime", "lastModifiedDateTime", "receivedDateTime", "sentDateTime")
    }

    if is_draft:
        message_date = message_dates.get("lastModifiedDateTime") or message_dates.get("createdDateTime")
    else:
        message_date = message_dates.get("receivedDateTime") or message_dates.get("sentDateTime")

    if not message_date:
        # Seems improbable but just in case!
        message_date = dt_now
        _log_message_notice(message, "all date fields are missing, using current UTC time")

    # I've seen Outlook report funny things for dates.
    message_dates["createdDateTime"] = min(
        message_dates.get("createdDateTime") or message_date,
        message_dates.get("lastModifiedDateTime") or message_date,
    )

    # I've seen Outlook report funny things for dates.
    message_dates["sentDateTime"] = min(
        message_dates.get("sentDateTime") or message_date,
        message_dates.get("receivedDateTime") or message_date,
    )

    if body_content:
        if content_type == "html":
            note_content = _convert_to_markdown(body_content)
        elif content_type != "text":
            _log_message_notice(message, f"field('contentType'): unexpected type: '{content_type}'")

    if not note_content:
        if body_content is None:
            stat, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING
        elif note_content is None:
            stat, level = ExportStat.STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED, logging.WARNING
        elif has_attachments:
            stat, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO
        else:
            stat, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO

        _exportstat_incr(export_stats, stat)
        _log_message_notice(message, stat.value, level)

        return  # don't write empty notes

    tag_vars = {}

    # Dynamic variables that can be referenced in tags.
    # {now.[datetime|date|time|year]}
    # {message.[date|sent|received|created|modified].[datetime|date|time|year]}
    for var_name, dt in (
        ("now", dt_now),
        ("message.date", message_date),
        ("message.created", message_dates.get("createdDateTime") or message_date),
        ("message.modified", message_dates.get("lastModifiedDateTime") or message_date),
        ("message.received", message_dates.get("receivedDateTime") or message_date),
        ("message.sent", message_dates.get("sentDateTime") or message_date),
    ):
        tag_vars[f"{{{var_name}.datetime}}"] = dt.strftime("%Y-%m-%dT%H:%M:%S")
        tag_vars[f"{{{var_name}.date}}"] = dt.strftime("%Y-%m-%d")
        tag_vars[f"{{{var_name}.time}}"] = dt.strftime("%H:%M:%S")
        tag_vars[f"{{{var_name}.year}}"] = f"{dt.year}"

    # Dynamic variables that can be referenced in metadata properties.
    metaprop_vars = tag_vars | {
        "{note.title}": title,
        "{message.subject}": subject,
    }

    _expand_note_metaprop(export_opts.note_metaprop, metaprop_vars, frontmatter)

    note_content = _append_note_tags(note_content, export_opts, tag_vars, frontmatter)

    if frontmatter:
        frontmatter_yaml = yaml.safe_dump(frontmatter, sort_keys=False).strip()
        note_content = f"---\n{frontmatter_yaml}\n---\n{note_content}"

    if not export_opts.dry_run:
        message_out_path.write_text(note_content, encoding="utf-8")
        os.utime(message_out_path, (message_date.timestamp(), message_date.timestamp()))
        logger.debug("Wrote note to '%s'", message_out_path.name)
    else:
        logger.debug("Wrote note to '%s' (dry run)", message_out_path.name)

    _exportstat_incr(export_stats, ExportStat.STAT_NOTE_CREATED)


def _export_folder_batch(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    progress: Progress,
    progress_task: TaskID,
    first_batch_data: GraphApiReturnT | None,
    url: str,
    params: dict[str, Any] | None,
    batch_limit: int,
    exported_ids: set[str],
    vault_dir: Path,
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> tuple[set[str], int, str | None]:
    """Export a batch of messages from a specific folder.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        progress (rich.Progress): The progress bar to update during the export process.
        progress_task (TaskID): The task ID for updating the progress bar.
        first_batch_data (GraphApiReturnT | None): First batch of messages or None to query `url`.
        url (str): The URL for the API request.
        params (dict[str, Any] | None): The query parameters for the API request.
        batch_limit (int): The maximum number of messages to export in this batch.
        exported_ids (set[str]): The set of message IDs that have already been exported.
        vault_dir (Path): The directory to export the messages to.
        export_opts (ExportOptions): Options for exporting messages.
        export_stats (ExportStats): Export statistics dictionary to update.

    Returns:
        tuple[set[str], int, str | None]: [set(Message-Id), Number of messages exported in this batch (including dups),
            Link to next batch of messages or None if there are no more batches]

    Raises:
        requests.RequestException derived error.
    """
    data = first_batch_data or graph_get_resource(session, graph_config, url, params)

    messages = data.get("value", [])
    next_url = data.get("@odata.nextLink")
    batch_ids = set()

    for message in messages:
        if message["id"] not in exported_ids:
            _export_message(message, vault_dir, export_opts, export_stats)
            batch_ids.add(message["id"])
        else:
            _exportstat_incr(export_stats, ExportStat.STAT_SKIPPED_DUPLICATE)

        progress.update(progress_task, advance=1)

        batch_limit -= 1
        if batch_limit <= 0:
            break

    return batch_ids, len(messages), next_url


def _export_folder(  # noqa: C901, PLR0912, PLR0913, PLR0915, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder: GraphFolderT,
    export_limit: int,
    batch_size: int,
    manifest_dir: Path | None,
    export_opts: ExportOptions,
) -> None:
    """Export messages from a specific folder to the output directory.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder (GraphFolderT): The folder object to export messages from.
        export_limit (int): The maximum number of messages to export. 0 to export all messages.
        batch_size (int): The number of messages to export in each batch.
        manifest_dir (Path|None): The directory to store the manifest file. If None, manifest is not used.
        export_opts (ExportOptions): Options for exporting messages.

    Returns:
        None

    Raises:
        requests.RequestException derived error.
    """
    manifest_ids = set()
    manifest_file_name = Path("export_manifest_" + hashlib.sha256(folder["id"].encode("utf-8")).hexdigest() + ".json")
    manifest_file_path = (manifest_dir / manifest_file_name) if manifest_dir else None

    if manifest_file_path and manifest_file_path.is_file():
        logger.info("Loading export manifest: '%s'", manifest_file_path)
        manifest_data = json.loads(manifest_file_path.read_text())
        manifest_ids = set(manifest_data.get("message_ids", []))

    exported_total_count = len(manifest_ids)
    folder_message_count = folder["totalItemCount"]
    vault_dir = export_opts.vault_dir

    if export_opts.folder_mkdir:
        vault_dir = vault_dir / Path(folder["folderPath"])
        vault_dir.mkdir(parents=True, exist_ok=True)

    if not export_opts.filter_test:
        console.print()
        console.print(f"[bold]Exporting folder '{folder['folderPath']}':")
        console.print(f"Exporting to '{vault_dir}'")

    # Dynamic variables that can be referenced in tags and metadata properties.
    folder_tag_vars = {
        "{folder.displayName}": folder["displayName"],
        "{folder.folderPath}": folder["folderPath"],
    }

    # Expand tags into a copy of export_opts (so caller's copy is not modified)
    export_opts = copy.replace(
        export_opts,
        note_tags=_expand_folder_tags(folder, export_opts.note_tags, folder_tag_vars),
        note_body_tags=_expand_folder_tags(folder, export_opts.note_body_tags, folder_tag_vars),
        note_metaprop=_expand_folder_metaprop(folder, export_opts.note_metaprop, folder_tag_vars),
    )

    export_stats: ExportStatsT = {
        # Start with stats that will always be output
        ExportStat.STAT_NOTE_CREATED: 0,
    }

    # https://learn.microsoft.com/en-us/graph/query-parameters
    params = {
        "$count": "true",
        "$top": batch_size,
        "$select": ",".join(sorted(EXPORT_MESSAGE_FIELDS)),
    }

    if export_opts.filter_query:
        params["$filter"] = export_opts.filter_query
        exported_total_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

    elif export_opts.search_query:
        params["$search"] = export_opts.search_query
        exported_total_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

    else:
        # The Graph API imposes rules when using $orderby and $filter together.
        # $search and $orderby cannot be used together.
        # Orderby multiple fields so order is more deterministic.
        params["$orderby"] = "lastModifiedDateTime,internetMessageId,subject"

    if exported_total_count:
        params["$skip"] = exported_total_count  # skip already exported messages
        _exportstat_incr(export_stats, ExportStat.STAT_SKIPPED_DUPLICATE, incr=exported_total_count)

    # https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
    # https://learn.microsoft.com/en-us/graph/api/message-get
    next_url = f"{GRAPH_API_URL}/me/mailFolders/{folder['id']}/messages"

    # Request the first batch now so we can get a count of the results.
    try:
        first_batch_data = graph_get_resource(session, graph_config, next_url, params)
    except requests.HTTPError:
        if export_opts.filter_query:
            field = "filter"
        elif export_opts.search_query:
            field = "search"
        else:
            field = None

        if field:
            # Try to be somewhat helpful
            err_console.print(
                f"Error: Failed to list messages in folder '{folder['folderPath']}', "
                f"possibly due to '{field}' criteria syntax errors.",
            )
        raise

    first_batch_messages = first_batch_data.get("value", [])

    # I'm guessing '@odata.count' is not provided when search results are less than '$top'.
    export_message_count: int = first_batch_data.get("@odata.count") or len(first_batch_messages)

    if export_opts.filter_test:
        if export_opts.filter_query:
            msg = f"'{folder['folderPath']}': {export_message_count} messages"
        else:
            msg = f"'{folder['folderPath']}': {export_message_count} messages"

        console.print(msg)
        return

    if not export_message_count:
        if not folder_message_count:
            msg = "Folder is empty"
        elif export_opts.filter_query:
            msg = "Folder contains no messages matching 'filter' criteria"
        elif export_opts.search_query:
            msg = "Folder contains no messages matching 'search' criteria"
        else:
            msg = "Folder contains {folder_message_count} messages but Graph API result was empty"

        console.print(msg)
        return

    if not export_limit:
        export_limit = export_message_count

    status_message = ["Exporting"]

    if export_limit == folder_message_count:
        status_message.append(f"all {folder_message_count}")
    elif export_opts.filter_query or export_opts.search_query:
        status_message.append(f"{export_message_count} (filtered) of {folder_message_count}")
    else:
        status_message.append(f"{export_limit} of {export_message_count}")

    status_message.append(f"messages from folder '{folder['folderPath']}'...")

    if manifest_file_path:
        status_message.append("(skipping dups)")
    if export_opts.dry_run:
        status_message.append("(dry run)")

    console.print(" ".join(status_message))

    # No need to show progress for a small number of messages.
    disable_progress = export_limit < MIN_PROGRESS_THRESHOLD

    with _logger_friendly_progress(
        *[BarColumn(), TaskProgressColumn()],
        console=console,
        disable=disable_progress,
        logger=logger,
    ) as progress:
        progress_task = progress.add_task("", total=export_limit, completed=exported_total_count)

        while next_url is not None and exported_total_count < export_limit:
            batch_ids, batch_total, next_url = _export_folder_batch(
                session,
                graph_config,
                progress,
                progress_task,
                first_batch_data,
                next_url,
                params,
                min(batch_size, export_limit - exported_total_count),
                manifest_ids,
                vault_dir,
                export_opts,
                export_stats,
            )

            manifest_ids |= batch_ids
            exported_total_count += batch_total
            params = None  # included in 'nextLink' automatically
            first_batch_data = None

            if manifest_file_path:
                # Doing this on every batch is inefficient for large folders but
                # ensures the manifest is mostly up-to-date if an exception occurs.
                manifest_file_path.write_text(
                    json.dumps(
                        {
                            "folderId": folder["id"],
                            "folderPath": folder["folderPath"],
                            "message_ids": list(manifest_ids),
                        },
                    ),
                )

                logger.debug("Wrote export manifest: '%s'", manifest_file_path)

    console.print("Folder results:")
    for stat, stat_count in export_stats.items():
        console.print(f"- {stat}: {stat_count}")


def _build_filter_query(
    message_filters: list[str] | None,
    filter_cond: FilterJoinCond,
    message_search: list[str] | None,
    search_cond: FilterJoinCond,
    *,
    filter_test: bool,
) -> tuple[str | None, str | None]:
    """Build Graph API filter and search queries based on critera.

    Args:
        message_filters (list[str] | None): List of filter criteria for messages.
        filter_cond (FilterJoinCond): Condition to join multiple filter criteria (AND/OR).
        message_search (list[str] | None): List of search criteria for messages.
        search_cond (FilterJoinCond): Condition to join multiple search criteria (AND/OR).
        filter_test (bool): Whether to perform a filter test without exporting.

    Returns:
        tuple[str | None, str | None]: (filter_query, search_query)
    """

    def _quote_search(s: str) -> str:
        return s if s[:1] == '"' else f'"{s}"'

    def _quote_paren(s: str) -> str:
        return f"({s})" if RE_AND_OR.search(s) else s

    filter_query, search_query = None, None

    if message_filters:
        filter_query = (
            message_filters[0]
            if len(message_filters) == 1
            else f" {filter_cond.value.lower()} ".join(_quote_paren(f) for f in message_filters)
        )

        if filter_test:
            console.print(f"Filter criteria: {filter_query}")
        else:
            logger.debug("Filter criteria: %s", filter_query)

    if message_search:
        search_query = _quote_search(
            message_search[0]
            if len(message_search) == 1
            else f" {search_cond.value.upper()} ".join(_quote_paren(s) for s in message_search),
        )

        if filter_test:
            console.print(f"Search criteria: {search_query}")
        else:
            logger.debug("Search criteria: %s", search_query)

    if filter_query and search_query:
        msg = "'filter' and 'search' criteria cannot be used together."
        raise typer.BadParameter(msg)

    if filter_test and not (filter_query or search_query):
        msg = "'filter' or 'search' criteria must be provided."
        raise typer.BadParameter(msg, param_hint="--filter-test")

    return filter_query, search_query


######################################################################
# Commands


@app.command()
def export(  # noqa: PLR0913, PLR0917
    ctx: typer.Context,
    folder_name: Annotated[
        list[str],
        typer.Argument(
            hidden=True,
            metavar="FOLDER...",
        ),
    ],
    vault_dir: Annotated[
        Path,
        typer.Argument(
            hidden=True,
            exists=True,
            dir_okay=True,
            file_okay=False,
            writable=True,
            metavar="VAULT",
        ),
    ],
    tag: Annotated[
        list[str] | None,
        typer.Option(
            "--tag",
            "-t",
            help="Tag to apply to notes (placed either in frontmatter or body); option can be used more than once.",
        ),
    ] = None,
    body_tag: Annotated[
        list[str] | None,
        typer.Option(
            "--body-tag",
            "-b",
            help="Tag to place in notes body; option can be used more than once.",
        ),
    ] = None,
    metaprop: Annotated[  # noqa: ARG001
        list[str] | None,
        typer.Option(
            "--metaprop",
            "-p",
            callback=_opt_metaprop_callback,
            help="Frontmatter metadata property to apply to notes; option can be used more than once.",
        ),
    ] = None,
    message_filter: Annotated[
        list[str] | None,
        typer.Option(
            "--filter",
            "-f",
            help="Export messages matching filter criteria; option can be used more than once.",
        ),
    ] = None,
    filter_cond: Annotated[
        FilterJoinCond,
        typer.Option(
            "--filter-cond",
            help="How to join multiple filter conditions (and/or).",
        ),
    ] = FilterJoinCond.cond_and,
    filter_test: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--filter-test",
            help="Test filter and search criteria and exit.",
        ),
    ] = False,
    message_search: Annotated[
        list[str] | None,
        typer.Option(
            "--search",
            "-s",
            help="Export messages matching search criteria; option can be used more than once.",
        ),
    ] = None,
    search_cond: Annotated[
        FilterJoinCond,
        typer.Option(
            "--search-cond",
            help="How to join multiple search conditions (and/or).",
        ),
    ] = FilterJoinCond.cond_and,
    place_tags_frontmatter: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            help="Place --tag tags in note frontmatter.",
        ),
    ] = False,
    place_tags_body_end: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            help="Place body tags at the end of the note instead of the top.",
        ),
    ] = False,
    use_manifest: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--manifest/--no-manifest",
            help="Write a manifest of exported messages and use to skip duplicates.",
        ),
    ] = True,
    limit: Annotated[
        int | None,
        typer.Option(
            min=1,
            help="Number of messages to export.",
        ),
    ] = None,
    batch_size: Annotated[
        int,
        typer.Option(
            min=GRAPH_API_BATCH_MIN,
            max=GRAPH_API_BATCH_MAX,
            help="Number of messages to to fetch in each Graph API request.",
        ),
    ] = GRAPH_API_BATCH_MAX,
    mkdir: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            help="Export to a directory with folder name.",
        ),
    ] = False,
    dry_run: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--dry-run",
            help="Dry run; do not write any files.",
        ),
    ] = False,
) -> None:
    """Export messages from mailbox FOLDER to VAULT."""
    # Options common to all commands
    ctx_obj: CommandContext = ctx.obj

    assert ctx_obj is not None  # noqa: S101
    assert isinstance(ctx_obj, CommandContext)  # noqa: S101

    common_opts = ctx_obj.common_opts
    config_file_export_opts: dict[str, Any] = ctx_obj.config_file_opts.get("export", {}) if ctx_obj.config_file_opts else {}
    config_file_tags: list[str] = config_file_export_opts.get("tag", [])
    config_file_body_tags: list[str] = config_file_export_opts.get("body_tag", [])
    config_file_metaprop: dict[str, Any] = config_file_export_opts.get("metaprop", {})
    config_file_filters: list[str] = config_file_export_opts.get("filter", [])
    config_file_search: list[str] = config_file_export_opts.get("search", [])

    note_tags = config_file_tags + (tag or [])
    note_body_tags = config_file_body_tags + (body_tag or [])
    note_metaprop = _merge_metaprops(config_file_metaprop, ctx_obj.export_metaprops)

    filter_query, search_query = _build_filter_query(
        config_file_filters + (message_filter or []),
        filter_cond,
        config_file_search + (message_search or []),
        search_cond,
        filter_test=filter_test,
    )

    try:
        graph_config = get_graph_config(common_opts)

        export_opts = ExportOptions(
            dry_run=dry_run,
            filter_query=filter_query,
            filter_test=filter_test,
            folder_mkdir=mkdir,
            graph_config=graph_config,
            note_body_tags=note_body_tags or None,
            note_metaprop=note_metaprop or None,
            note_tags=note_tags or None,
            place_tags_body_end=place_tags_body_end,
            place_tags_frontmatter=place_tags_frontmatter,
            search_query=search_query,
            vault_dir=vault_dir,
        )

        folders: dict[str, GraphFolderT] = {}

        # Fail early if any folder does not exist
        with sessions.Session() as session:
            for folder_path in folder_name:
                folder = graph_get_folder_by_path(session, graph_config, Path(folder_path), MAILBOX_ROOT)
                if not folder:
                    err_console.print(f"Error: '{folder_path}': No such mail folder.")
                    raise typer.Exit(code=1)

                folders[folder_path] = folder

        logger.info("Querying Graph API for %d messages at a time", batch_size)

        if not filter_test:
            console.print(f"Exporting messages from {len(folders)} folders...")

        with sessions.Session() as session:
            for folder in folders.values():
                _export_folder(
                    session,
                    graph_config,
                    folder,
                    limit or 0,
                    batch_size,
                    common_opts.cache_dir if use_manifest else None,
                    export_opts,
                )

    except requests.RequestException as err:
        resp_error_code, resp_error_message = None, None

        if err.response is not None:
            # See if Graph API returned an error message in the response body
            resp_json = err.response.json()
            resp_error = resp_json.get("error", {})
            resp_error_code, resp_error_message = resp_error.get("code"), resp_error.get("message")

        if resp_error_code and resp_error_message:
            err_console.print(f"Error: {resp_error_message} [{resp_error_code}]; {err}")
        else:
            err_console.print(f"Error: {err}")

        raise typer.Exit(code=1) from err
    except AuthRequiredError as err:
        err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
        raise typer.Exit(code=1) from err
    except ValueError as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err
