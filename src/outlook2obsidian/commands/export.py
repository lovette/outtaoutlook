"""outlook2obsidian export command."""

import base64
import binascii
import copy
import json
import logging
import mimetypes
import os
import re
import sys
import textwrap
import unicodedata
from collections.abc import Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from enum import StrEnum
from pathlib import Path
from types import UnionType
from typing import TYPE_CHECKING, Annotated, Any, get_args
from urllib.parse import unquote

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
    CommandContext,
    GraphApiConfig,
    GraphApiError,
    GraphApiReturnT,
    GraphAttachmentT,
    GraphAuthRequiredError,
    GraphFolderT,
    GraphMessageT,
    NoteMetaPropT,
    console,
    datetime_now,
    err_console,
    get_graph_config,
    get_hash_filename,
    graph_get_resource,
    graph_resolve_folder_paths,
    path_with_vars,
    print_graph_request_exception,
)

if TYPE_CHECKING:
    from collections.abc import Generator

app = typer.Typer(add_completion=False)
logger = logging.getLogger(APP_NAME)  # initialized in __main__


######################################################################
# Constants

TITLE_UNTITLED = "Untitled"
TITLE_MAXLEN = 255 - 25  # macOS limit minus a few characters for whatever `_unique_out_path` tacks on
MIN_PROGRESS_THRESHOLD = 10  # Hide progress if only a few messages
EXPORT_FOLDER_MSGFOLDERROOT = "msgfolderroot"

ATTACHMENT_DEFAULT_SUFFIX = ".bin"
ATTACHMENT_DEFAULT_NAME = "unnamed"  # suffix will be guessed
ATTACHMENTS_DEFAULT_FOLDER = "assets"
ATTACHMENTS_FOLDER_NOTE = "note"
ATTACHMENTS_FOLDER_VAULT = "vault"

# https://learn.microsoft.com/en-us/graph/api/attachment-get
ATTACHMENT_DATATYPE_FILE = "#microsoft.graph.fileAttachment"
ATTACHMENT_DATATYPE_ITEM = "#microsoft.graph.itemAttachment"
ATTACHMENT_DATATYPE_REFERENCE = "#microsoft.graph.referenceAttachment"

MESSAGE_DATATYPE_EVENT_MESSAGE = "#microsoft.graph.eventMessage"
MESSAGE_DATATYPE_EVENT_MESSAGE_REQUEST = "#microsoft.graph.eventMessageRequest"
MESSAGE_DATATYPE_EVENT_MESSAGE_RESPONSE = "#microsoft.graph.eventMessageResponse"

MESSAGE_DATATYPE_EVENT_MESSAGE_SET = {
    MESSAGE_DATATYPE_EVENT_MESSAGE,
    MESSAGE_DATATYPE_EVENT_MESSAGE_REQUEST,
    MESSAGE_DATATYPE_EVENT_MESSAGE_RESPONSE,
}

# These are banned outright: \/:
# Links do not work if title includes: #^[]|
# macOS is fine with: *"<>?
# Title cannot start with a period.
# \x00-\x1f\x7f = [:cntrl:]
RE_TITLE_IGNORECHARS = re.compile(r"[\[\]\/\\:|^#\x00-\x1f\x7f]")
RE_TITLE_WHITESPACE = re.compile(r"\s+")

# Simple detection if filter/search critera contain 'and' or 'or' conditions.
RE_AND_OR = re.compile(r"\s(and|or)\s", re.IGNORECASE)

# Tag and property variables in the format '{tag.propertyName}'
RE_TAG_VAR = re.compile(r"{[a-zA-Z.]+}")

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


class ExportError(GraphApiError):
    """Export failure."""


class FilterJoinCond(StrEnum):
    """Choices for joining filter conditions."""

    cond_and = "and"
    cond_or = "or"


class ExportStat(StrEnum):
    """Export statistic messages."""

    STAT_NOTE_CREATED = "Notes created"
    STAT_ATTACHMENT_LINKED = "Attachments linked"
    STAT_SKIPPED_DUPLICATE = "Skipped; Message already exported"
    STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED = "Skipped; Markdown conversion failed"
    STAT_SKIPPED_NO_MESSAGE_BODY = "Skipped; No message body"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS = "Skipped; No note content, just attachments (not being saved)"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS = "Skipped; No note content"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN = "Skipped; No note content, unknown reason"
    STAT_SKIPPED_EVENT_MESSAGE = "Skipped; Calendar event"


@dataclass
class ExportOptions:
    """Export command options."""

    attachments_folder: str | Path  # `ATTACHMENTS_FOLDER_*`, folder name or absolute path.
    attachments_mkdir: bool  # Create mail folder hierarchy in attachments directory.
    dry_run: bool  # Whether to perform a dry run without making any changes.
    filter_query: str | None  # --filter criteria from command line and config file.
    filter_test: bool  # Whether to perform a filter test without exporting.
    notes_mkdir: bool  # Create mail folder hierarchy in notes directory.
    note_body_tags: list[str] | None  # --body-tag tags from command line and config file.
    note_metaprop: NoteMetaPropT | None  # --metaprop properties from command line and config file.
    note_tags: list[str] | None  # --tag tags from command line and config file.
    notes_dir: Path  # Directory to save notes (set internally.)
    place_tags_body_end: bool  # True to place body tags at the end of the note instead the top.
    place_tags_frontmatter: bool  # True to place non-body tags in frontmatter.
    search_query: str | None  # --search criteria from command line and config file.
    vault_dir: Path  # Root vault directory (set by user.)
    with_attachments: bool  # Whether to download and link attachments to notes.


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


def _opt_attachments_folder_callback(attachments_folder: str) -> str:
    """Handle `--attachments-folder` command option.

    Args:
        attachments_folder (str): Command line option value.

    Returns:
        str: Special keyword, valid folder name or absolute path.
    """
    # Special keywords.
    if attachments_folder.upper() == ATTACHMENTS_FOLDER_VAULT or attachments_folder.startswith(ATTACHMENTS_FOLDER_NOTE):
        return attachments_folder

    # Name of subfolder such as "assets" or "attachments".
    if len(Path(attachments_folder).parts) == 1:
        if attachments_folder != _safe_title(attachments_folder, ""):
            errmsg = f"'{attachments_folder}' is not a valid attachments directory name."
            raise typer.BadParameter(errmsg)

        return attachments_folder

    # Absolute or relative path to a directory optionally with ~ or $ENV variables.
    attachments_dir_path = path_with_vars(attachments_folder)
    if attachments_dir_path.is_dir():
        return str(attachments_dir_path)

    errmsg = f"Directory '{attachments_folder}' does not exist."
    raise typer.BadParameter(errmsg)


def _opt_metaprop_callback(
    ctx: typer.Context,
    param: typer.CallbackParam,  # noqa: ARG001
    meta: list[str] | None,
) -> list[str] | None:
    """Handle `--metaprop` command option.

    Note:
        The values of `default_map` are used as the default values for command options.
        This function is called even if option is not specified.
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


def _safe_title(subject: str, default: str) -> str:
    """Clean up title so it is safe to be used as a file name.

    Args:
        subject (str): Message subject.
        default (str): Default title if subject is empty.

    Returns:
        str
    """
    subject = subject.replace(":", " ")  # Common separator so we'll replace it with a safe character.
    subject = RE_TITLE_IGNORECHARS.sub("", subject)
    subject = RE_TITLE_WHITESPACE.sub(" ", subject)

    if len(subject) > TITLE_MAXLEN:
        # Truncate text while preserving whole words if possible
        subject = textwrap.shorten(subject, width=TITLE_MAXLEN, placeholder="")

    subject = subject.strip(" .")

    return subject or default


def _unique_out_path(p: Path, date_hint: datetime | None = None) -> Path:
    """Create a Path with unique file name, optionally using a date hint.

    Args:
        p (Path): Existing path.
        date_hint (datetime | None): Optional datetime to include in the filename for uniqueness.

    Raises:
        ExportError: Raised if too many copies of the file already exist.

    Returns:
        Path: Unique path with a non-conflicting file name.
    """
    n = 1
    max_paths = 10000
    unique_path = p

    # Keep in mind we need to keep the final filename within TITLE_MAXLEN characters.

    while unique_path.exists() and n < max_paths:
        unique_path = p.with_stem(f"{p.stem} ({n})")
        n += 1

    if n == max_paths and date_hint:
        # Try using the date hint to create a unique filename.
        n = 1
        unique_path = p.with_stem(f"{p.stem} {date_hint:%Y%m%d%H%M%S}")
        while unique_path.exists() and n < max_paths:
            unique_path = p.with_stem(f"{p.stem} {date_hint:%Y%m%d%H%M%S} ({n})")
        n += 1

    if not unique_path.exists():
        return unique_path

    errmsg = f"{n} copies of '{p}' already exist?!"
    raise ExportError(errmsg)


def _coerce_datetime(dt: str | None) -> datetime | None:
    """Convert a datetime string in ISO format into a `datetime` object.

    Args:
        dt (str | None): Datetime string in ISO format.

    Returns:
        datetime | None: Datetime object if conversion is successful, otherwise None.
    """
    if not dt:
        return None

    try:
        return datetime.fromisoformat(dt)
    except ValueError:
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


def _expand_vars_dict(d: NoteMetaPropT, prop_vars: dict[str, str]) -> NoteMetaPropT:
    """Expand dynamic variable references in a dictionary of metadata properties.

    Args:
        d (NoteMetaPropT): Dictionary of properties to expand.
        prop_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        NoteMetaPropT: Updated dictionary with message-related placeholders expanded.
    """

    def _varrepl(m: re.Match) -> str:
        var_name: str = m.group(0)
        return prop_vars.get(var_name, var_name)

    if not d:
        return {}

    d = d.copy()  # don't modify caller

    for key, value in d.items():
        if isinstance(value, str):
            d[key] = RE_TAG_VAR.sub(_varrepl, value)

    return d


def _expand_vars_list(lst: list[str], prop_vars: dict[str, str]) -> list[str]:
    """Expand dynamic variable references in a list of tags.

    Args:
        lst (list[str]): List of tags or metadata properties to expand.
        prop_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        list[str]: New list with dynamic placeholders expanded.
    """

    def _varrepl(m: re.Match) -> str:
        var_name: str = m.group(0)
        return prop_vars.get(var_name, var_name)

    if not lst:
        return []

    return [RE_TAG_VAR.sub(_varrepl, _tag) for _tag in lst]


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

    tags = _expand_vars_list(tags, tag_vars)

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

    metaprop = _expand_vars_dict(metaprop, prop_vars)

    for name, value in metaprop.items():
        if isinstance(value, str) and "{folder.folderPath.parts}" in value:
            # Frontmatter can include list types so we can simply pass it along
            path_parts = Path(folder["folderPath"]).parts
            metaprop[name] = path_parts[0] if len(path_parts) == 1 else path_parts

    return metaprop


def _init_note_frontmatter(metaprop: NoteMetaPropT | None, prop_vars: dict[str, str]) -> FrontmatterT:
    """Initialize note frontmatter with expanded metadata properties.

    Args:
        metaprop (NoteMetaPropT | None): Dictionary of metadata properties to expand.
        prop_vars (dict[str, str]): Dictionary of replacement patterns.

    Returns:
        FrontmatterT: Initialized frontmatter dictionary.
    """

    def is_float(value: str) -> bool:
        try:
            float(value)
        except ValueError:
            return False
        return True

    frontmatter: FrontmatterT = {}

    metaprop = _expand_vars_dict(metaprop or {}, prop_vars)

    # Properties support some scalar types
    for key, value in metaprop.items():
        if not isinstance(value, str):
            frontmatter[key] = value
        elif value.isdecimal():
            frontmatter[key] = int(value)
        elif "." in value and is_float(value):
            frontmatter[key] = Decimal(value)
        elif value.lower() in ("true", "false"):
            frontmatter[key] = value.lower() == "true"
        else:
            frontmatter[key] = value

    return frontmatter


def _append_note_tags(
    note_content: str,
    export_opts: ExportOptions,
    tag_vars: dict[str, str],
    frontmatter: FrontmatterT,
) -> str:
    """Append tags to note content and/or frontmatter.

    Args:
        note_content (str): Note content.
        export_opts (ExportOptions): Export options.
        tag_vars (dict[str, str]): Dictionary of replacement patterns.
        frontmatter (FrontmatterT): Frontmatter dictionary. Frontmatter tags will be added here if applicable.

    Returns:
        str
    """
    body_tags = _expand_vars_list(export_opts.note_body_tags or [], tag_vars)
    frontmatter_tags = _expand_vars_list(export_opts.note_tags or [], tag_vars)

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


def _list_messages(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    url: str,
    params: dict[str, Any] | None = None,
) -> GraphApiReturnT:
    """Get a list of messages.

    Args:
        session (sessions.Session): Requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        url (str): Entity URL.
        params (dict, optional): Query parameters for the GET request. Defaults to None.

    Returns:
        GraphApiReturnT: JSON response from the Microsoft Graph API.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    headers = {
        # https://learn.microsoft.com/en-us/graph/outlook-immutable-id
        "Prefer": 'IdType="ImmutableId"',
    }

    return graph_get_resource(session, graph_config, url, params=params, headers=headers)


def _list_attachments(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
) -> list[GraphAttachmentT]:
    """List properties for all message attachments.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.

    Returns:
        list[GraphAttachmentT]

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    # Best I can tell, '/attachments' returns all attachments and has no 'nextLink'
    url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments"
    data = graph_get_resource(session, graph_config, url)
    return data.get("value", [])


def _get_attachment_content(  # noqa: C901
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    attachment: GraphAttachmentT,
    attachment_name_prefix: str,
) -> tuple[str | None, str | None]:
    """Get attachment content and ensure 'name' property is appropriate and valid.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.
        attachment (GraphAttachmentT): Attachment properties.
        attachment_name_prefix (str): Prefix to prepend to attachment filenames. Must be filename safe.

    Returns:
        tuple[str, str]: Attachment content bytes and content type.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _warn_attachment(name: Path, errmsg: str, level: int = logging.WARNING) -> None:
        _log_message_notice(message, f'attachment("{name}"): {errmsg}', level)

    def _get_item_expanded_content() -> tuple[dict[str, Any], str | None]:
        # Get properties of the attached contact, event, or message.
        url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments/{attachment.get('id')}"
        params = {"$expand": "microsoft.graph.itemattachment/item"}
        data = graph_get_resource(session, graph_config, url, params=params)
        item = data.get("item", {})
        return item, item.get("@odata.type")

    def _get_item_value() -> tuple[str | None, str | None]:
        # Get raw attachment content for a contact, event or message.
        url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments/{attachment.get('id')}/$value"
        data = graph_get_resource(session, graph_config, url)
        content, content_type = data.get("contentBytes"), data.get("contentType")
        return content, content_type

    def _get_item_content(name: Path) -> tuple[str | None, str | None]:
        # Attachment is a contact, event or message.
        _, item_data_type = _get_item_expanded_content()
        content, content_type = _get_item_value()

        if item_data_type == "#microsoft.graph.message":
            if content_type != "text/plain":
                _warn_attachment(name, f"Content type for '{item_data_type}' is '{content_type}', expected 'text/plain'")
            content_type = "message/rfc822"  # guess_extension() -> '.eml'
        elif item_data_type == "#microsoft.graph.contact":
            if not content_type or not content_type.endswith("vcard"):
                _warn_attachment(name, f"Content type for '{item_data_type}' is '{content_type}', expected 'text/vcard'")
            content_type = "text/x-vcard"  # guess_extension() -> '.vcf'
        else:
            _warn_attachment(name, f"Unsupported item data type '{item_data_type}'")

        return content, content_type

    name = Path(_safe_title(attachment.get("name", ""), ATTACHMENT_DEFAULT_NAME))
    data_suffix = name.suffix or None

    # We require every attachment to have a suffix.
    # Lowercase to keep consistent with guess_extension() result.
    name = name.with_suffix(data_suffix.lower()) if data_suffix else Path(f"{name.stem}{ATTACHMENT_DEFAULT_SUFFIX}")

    data_type: str | None = attachment.get("@odata.type")
    content: str | None = attachment.get("contentBytes")
    content_type: str | None = attachment.get("contentType")

    if data_type == ATTACHMENT_DATATYPE_ITEM:
        if content is not None:
            _warn_attachment(name, f"@data.type '{data_type}' has unexpected 'contentBytes' value")
        content, content_type = _get_item_content(name)
    elif data_type == ATTACHMENT_DATATYPE_FILE:
        if content is None:
            _warn_attachment(name, f"@data.type '{data_type}' is missing 'contentBytes' value")
    else:
        _warn_attachment(name, f"@data.type '{data_type}' is not supported")

    if content is None:
        _warn_attachment(name, "No content", logging.ERROR)
        return None, None

    if not data_suffix and content_type:
        guess_suffix = mimetypes.guess_extension(content_type, strict=False)

        if guess_suffix:
            name = name.with_suffix(guess_suffix)
        else:
            _warn_attachment(name, f"Content type '{content_type}' is not assigned a standard suffix")

    attachment["name"] = f"{attachment_name_prefix} {name}"

    return content, content_type


def _save_attachment(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    attachment: GraphAttachmentT,
    attachment_name_prefix: str,
    note_date: datetime,
    export_opts: ExportOptions,
) -> Path | None:
    """Save a message attachment to the local filesystem.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.
        attachment (GraphAttachmentT): Attachment properties.
        attachment_name_prefix (str): Prefix to prepend to attachment filenames. Must be filename safe.
        note_date (datetime): The date of the note.
        export_opts (ExportOptions): Export options.

    Returns:
        Path | None: Path to saved attachment or None if save failed for a non-fatal reason.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    content, content_type = _get_attachment_content(
        session,
        graph_config,
        message,
        attachment,
        attachment_name_prefix,
    )

    if content is None:
        return None

    assert isinstance(export_opts.attachments_folder, Path)  # noqa: S101
    assert export_opts.attachments_folder.is_absolute()  # noqa: S101

    attachment_path = _unique_out_path(export_opts.attachments_folder / Path(attachment["name"]), note_date)

    if export_opts.dry_run:
        logger.debug("Wrote attachment to '%s' (%s) (dry run)", attachment_path.name, content_type)
    elif content_type and content_type.startswith(("text/", "message/")):
        attachment_path.write_text(content, encoding="utf-8")
        logger.debug("Wrote attachment to '%s' (%s)", attachment_path.name, content_type)
    else:
        try:
            with attachment_path.open("wb") as f:
                f.write(base64.b64decode(content))
            logger.debug("Wrote attachment to '%s' (%s)", attachment_path.name, content_type)
        except binascii.Error as err:
            _log_message_notice(message, f'attachment("{attachment_path.name}"): b64decode failed: {err}', logging.ERROR)
            return None

    if attachment_path.is_file():
        os.utime(attachment_path, (note_date.timestamp(), note_date.timestamp()))

    return attachment_path


def _save_attachments(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    note_path: Path,
    note_date: datetime,
    export_opts: ExportOptions,
) -> list[str]:
    """Save message attachments to the local filesystem.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.
        note_path (Path): The path to the note file.
        note_date (datetime): The date of the note.
        export_opts (ExportOptions): Export options.

    Returns:
        list[str]: Paths relative to vault or URIs (file://) for absolute paths.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _attachment_rel_links(abs_paths: list[Path], note_dir: Path, vault_dir: Path) -> list[str]:
        """Convert absolute paths into paths relative to the note directory or the vault.

        For links without a path, Obsidian seems to search:

        1. Note folder
        2. Note subfolders
        3. Vault root
        4. Vault subfolders, alphabetically

        Outside the note directory and vault, use a file URI.
        """
        always_rel = False
        rel_paths: list[str] = []

        for abs_path in abs_paths:
            assert abs_path.is_absolute()  # noqa: S101 ; as_uri fails otherwise
            if always_rel:
                try:
                    rel_path = str(abs_path.relative_to(vault_dir, walk_up=True))
                except ValueError:
                    rel_path = abs_path.as_uri()
            elif abs_path.is_relative_to(note_dir) or abs_path.is_relative_to(vault_dir):
                rel_path = abs_path.name
            else:
                rel_path = abs_path.as_uri()

            rel_paths.append(rel_path)

        return rel_paths

    attachment_paths: list[Path] = []

    for attachment in _list_attachments(session, graph_config, message):
        attachment_path = _save_attachment(
            session,
            graph_config,
            message,
            attachment,
            note_path.stem,
            note_date,
            export_opts,
        )

        if attachment_path:
            attachment_paths.append(attachment_path)

    return _attachment_rel_links(attachment_paths, note_path.parent, export_opts.vault_dir)


def _export_message(  # noqa: C901, PLR0912, PLR0915
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> None:
    """Export a message.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): The message.
        export_opts (ExportOptions): Export options.
        export_stats (ExportStatsDict): Export statistics dictionary to update.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _fix_dates(message: GraphMessageT, is_draft: bool) -> None:  # noqa: FBT001
        dt_now = datetime_now()

        for field in ("createdDateTime", "lastModifiedDateTime", "receivedDateTime", "sentDateTime"):
            value = message.get(field)
            dt = _coerce_datetime(value)

            if not dt:
                if field not in message:
                    _log_message_notice(message, f"field({field}): missing date field")
                elif is_draft and field in {"sentDateTime", "receivedDateTime"}:
                    pass  # Drafts aren't expected to have sent or received dates, we'll set something below
                elif value:
                    _log_message_notice(message, f"field({field}): unexpected date format: '{value}'")
                else:
                    _log_message_notice(message, f"field({field}): no value")

            # Ensure all fields are set to a valid datetime
            message[field] = dt or dt_now

        # Outlook does crazy things with dates, so we do our best to help them make some kind of sense.

        message["sentDateTime"] = min(message["sentDateTime"], message["receivedDateTime"])
        message["receivedDateTime"] = max(message["sentDateTime"], message["receivedDateTime"])

        message["createdDateTime"] = min(
            message["createdDateTime"],
            message["lastModifiedDateTime"],
            message["sentDateTime"],
            message["receivedDateTime"],
        )

        message["lastModifiedDateTime"] = max(
            message["createdDateTime"],
            message["lastModifiedDateTime"],
            message["sentDateTime"],
            message["receivedDateTime"],
        )

    def _rel_path_link(rel_path_or_uri: str, *, inline: bool) -> str:
        if rel_path_or_uri.startswith("file://"):
            # Links external to the vault are normal markdown links
            abs_path = Path(rel_path_or_uri[7:])
            return f"[{unquote(abs_path.name)}]({rel_path_or_uri})"

        return f"![[{rel_path_or_uri}]]" if inline else f"[[{rel_path_or_uri}]]"

    subject: str = message.get("subject") or TITLE_UNTITLED
    title = _safe_title(subject, TITLE_UNTITLED)
    is_draft: bool | None = message.get("isDraft")
    body_content: str | None = message.get("body", {}).get("content")
    content_type: str | None = message.get("body", {}).get("contentType")
    has_attachments: bool | None = message.get("hasAttachments")
    dt_now = datetime_now()
    note_body = body_content
    data_type: str | None = message.get("@odata.type")
    is_calendar_related = data_type and data_type in MESSAGE_DATATYPE_EVENT_MESSAGE_SET

    _fix_dates(message, bool(is_draft))

    note_date: datetime = message["receivedDateTime"] or message["sentDateTime"]
    if is_draft:
        note_date = message["lastModifiedDateTime"] or message["createdDateTime"]

    if body_content:
        if is_calendar_related:
            pass  # Calendar requests and responses are skipped
        elif content_type == "html":
            note_body = _convert_to_markdown(body_content)
        elif content_type != "text":
            _log_message_notice(message, f"field('contentType'): unexpected type: '{content_type}'")

    if not note_body or is_calendar_related:
        skip_reason, level = None, logging.INFO

        if is_calendar_related:
            skip_reason, level = ExportStat.STAT_SKIPPED_EVENT_MESSAGE, logging.DEBUG
        elif body_content is None:
            skip_reason, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING
        elif note_body is None:
            skip_reason, level = ExportStat.STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED, logging.WARNING
        elif has_attachments and export_opts.with_attachments:
            pass  # export attachments
        elif not has_attachments:
            skip_reason, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO
        elif not export_opts.with_attachments:
            skip_reason, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO
        else:
            skip_reason, level = ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN, logging.ERROR

        if skip_reason:
            _exportstat_incr(export_stats, skip_reason)
            _log_message_notice(message, skip_reason.value, level)
            return  # don't write empty notes

        note_body = ""  # content will be links to attachments

    tag_vars = {}

    # Dynamic variables that can be referenced in tags.
    # {now.[datetime|date|time|year]}
    # {message.[date|sent|received|created|modified].[datetime|date|time|year]}
    for var_name, dt in (
        ("now", dt_now),
        ("message.date", note_date),
        ("message.created", message["createdDateTime"]),
        ("message.modified", message["lastModifiedDateTime"]),
        ("message.received", message["receivedDateTime"]),
        ("message.sent", message["sentDateTime"]),
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

    note_content: list[str] = []
    attachment_links: list[str] = []
    note_path = _unique_out_path(export_opts.notes_dir / Path(title + ".md"), note_date)

    # Save attachments and get links to them
    if has_attachments and export_opts.with_attachments:
        attachment_paths = _save_attachments(
            session,
            graph_config,
            message,
            note_path,
            note_date,
            export_opts,
        )

        inline = len(attachment_paths) == 1 and not note_body  # Note consists of a single inline link
        attachment_links = [_rel_path_link(rel_path_or_uri, inline=inline) for rel_path_or_uri in attachment_paths]

        _exportstat_incr(export_stats, ExportStat.STAT_ATTACHMENT_LINKED, len(attachment_paths))

    frontmatter = _init_note_frontmatter(export_opts.note_metaprop, metaprop_vars)

    note_body = _append_note_tags(note_body, export_opts, tag_vars, frontmatter)

    if frontmatter:
        # Frontmatter goes first
        note_content = [
            "---",
            yaml.safe_dump(frontmatter, sort_keys=False).strip(),
            "---",
            *note_content,
        ]

    if attachment_links:
        note_content.extend(attachment_links)
        note_content.append("")  # blank line

    # Note body goes last
    note_content.append(note_body)

    if not export_opts.dry_run:
        note_text = "\n".join(note_content).strip() + "\n"
        note_path.write_text(note_text, encoding="utf-8")
        os.utime(note_path, (note_date.timestamp(), note_date.timestamp()))
        logger.debug("Wrote note to '%s'", note_path.name)
    else:
        logger.debug("Wrote note to '%s' (dry run)", note_path.name)

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
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> tuple[int, str | None]:
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
        exported_ids (set[str]): The set of message IDs that have already been exported. Modified in-place.
        export_opts (ExportOptions): Options for exporting messages.
        export_stats (ExportStats): Export statistics dictionary to update.

    Returns:
        tuple[int, str | None]: [Number of messages exported in this batch (including dups),
            Link to next batch of messages or None if there are no more batches]

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    data = first_batch_data or _list_messages(session, graph_config, url, params=params)

    messages = data.get("value", [])
    next_url = data.get("@odata.nextLink")

    for message in messages:
        if message["id"] not in exported_ids:
            _export_message(
                session,
                graph_config,
                message,
                export_opts,
                export_stats,
            )

            exported_ids.add(message["id"])
        else:
            _exportstat_incr(export_stats, ExportStat.STAT_SKIPPED_DUPLICATE)

        progress.update(progress_task, advance=1)

        batch_limit -= 1
        if batch_limit <= 0:
            break

    return len(messages), next_url


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
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _mkdir(dir_path: Path, name: str, *, parents: bool = False, exist_ok: bool = True) -> None:
        if not dir_path.is_dir():
            if not export_opts.dry_run:
                try:
                    dir_path.mkdir(parents=parents, exist_ok=exist_ok)
                except OSError as err:
                    err_console.print(f"Failed to create {name} directory: {err}")
                    raise

                logger.info("Created %s directory '%s'", name, dir_path)
            else:
                logger.info("Created %s directory '%s' (dry run)", name, dir_path)

    def _write_manifest() -> None:
        if manifest_file_path:
            manifest_file_path.write_text(
                json.dumps(
                    {
                        "date": datetime_now().isoformat(),
                        "folderId": folder["id"],
                        "folderPath": folder["folderPath"],
                        "messageImmutableIds": list(manifest_ids),
                    },
                ),
            )

            logger.debug("Wrote export manifest: '%s'", manifest_file_path)

    folder_message_count = folder["totalItemCount"]

    if not export_opts.filter_test:
        console.print()
        console.print(f"[bold]Exporting folder '{folder['folderPath']}':")

    if not folder_message_count:
        console.print("Folder is empty")
        return

    manifest_ids: set[str] = set()
    manifest_file_name = get_hash_filename(folder["id"], "export_manifest", ".json")
    manifest_file_path = (manifest_dir / manifest_file_name) if manifest_dir else None

    if manifest_file_path and manifest_file_path.is_file():
        logger.info("Loading export manifest: '%s'", manifest_file_path)
        manifest_data = json.loads(manifest_file_path.read_text())
        manifest_ids = set(manifest_data.get("messageImmutableIds", []))

    exported_total_count = len(manifest_ids)
    notes_dir = export_opts.notes_dir
    attachments_dir_path = Path(export_opts.attachments_folder)  # assume it's an absolute path to start
    attachments_mkdir = export_opts.attachments_mkdir

    if export_opts.notes_mkdir:
        notes_dir = notes_dir / Path(folder["folderPath"])

    if isinstance(export_opts.attachments_folder, Path):
        assert attachments_dir_path.is_dir()  # noqa: S101 ; caller should have verified!
        assert attachments_dir_path.is_absolute()  # noqa: S101 ; caller should have verified!
    elif export_opts.attachments_folder.upper() == ATTACHMENTS_FOLDER_VAULT:
        attachments_dir_path = export_opts.vault_dir
        attachments_mkdir = False
    elif export_opts.attachments_folder.upper().startswith(ATTACHMENTS_FOLDER_NOTE):  # allow plural "NOTES"
        attachments_dir_path = notes_dir
        attachments_mkdir = False
    elif not attachments_dir_path.is_absolute():
        attachments_dir_path = notes_dir / _safe_title(export_opts.attachments_folder, ATTACHMENTS_DEFAULT_FOLDER)
        attachments_mkdir = False

    if attachments_mkdir:
        attachments_dir_path = attachments_dir_path / Path(folder["folderPath"])

    if not export_opts.filter_test:
        console.print(f"Saving notes to '{notes_dir}'")
        _mkdir(notes_dir, "notes", parents=export_opts.notes_mkdir)

        if export_opts.with_attachments:
            console.print(f"Saving attachments to '{attachments_dir_path}'")
            _mkdir(attachments_dir_path, "attachments", parents=attachments_mkdir)

    # Dynamic variables that can be referenced in tags and metadata properties.
    folder_tag_vars = {
        "{folder.displayName}": folder["displayName"],
        "{folder.folderPath}": folder["folderPath"],
    }

    # Tweak export options without modifying caller
    export_opts = copy.replace(
        export_opts,
        notes_dir=notes_dir,
        attachments_folder=attachments_dir_path,
        note_tags=_expand_folder_tags(folder, export_opts.note_tags, folder_tag_vars),
        note_body_tags=_expand_folder_tags(folder, export_opts.note_body_tags, folder_tag_vars),
        note_metaprop=_expand_folder_metaprop(folder, export_opts.note_metaprop, folder_tag_vars),
    )

    export_stats: ExportStatsT = {
        # Start with stats that will always be output
        ExportStat.STAT_NOTE_CREATED: 0,
    }

    if export_opts.with_attachments:
        export_stats[ExportStat.STAT_ATTACHMENT_LINKED] = 0

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
        first_batch_data = _list_messages(session, graph_config, next_url, params=params)
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
        if export_opts.filter_query:
            msg = "Folder contains no messages matching 'filter' criteria"
        elif export_opts.search_query:
            msg = "Folder contains no messages matching 'search' criteria"
        else:
            msg = f"Folder contains {folder_message_count} messages but Graph API result was empty"

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
            try:
                batch_total, next_url = _export_folder_batch(
                    session,
                    graph_config,
                    progress,
                    progress_task,
                    first_batch_data,
                    next_url,
                    params,
                    min(batch_size, export_limit - exported_total_count),
                    manifest_ids,
                    export_opts,
                    export_stats,
                )

            except requests.RequestException, GraphApiError, ValueError, OSError:
                # Try to keep the manifest up-to-date even if an error occurs.
                _write_manifest()
                raise
            else:
                # Doing this on every batch is inefficient for large folders but
                # ensures the manifest is mostly up-to-date if a truly fatal error occurs.
                _write_manifest()

            exported_total_count += batch_total
            params = None  # included in 'nextLink' automatically
            first_batch_data = None

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
        return s if (s.startswith('"') and s.endswith('"')) else f'"{s.replace('"', '\\"')}"'

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


def _validate_config_file_collection(
    config_opt: str,
    collection: list[Any] | dict[str, Any],
    collection_type: type,
    item_type: type | UnionType,
) -> None:
    if collection is not None:
        if not isinstance(collection, collection_type):
            errmsg = (
                f"Configuration file error: Expected '{config_opt}' to be a {collection_type.__name__}, "
                f"got a {type(collection).__name__} instead"
            )
            raise typer.BadParameter(errmsg)

        items = collection if isinstance(collection, list) else collection.values()

        if isinstance(item_type, UnionType):
            item_types = get_args(item_type)  # str | int | bool
            item_types_str = ", ".join(t.__name__ for t in item_types)  # "str, int, bool"

            for item in items:
                if not any(isinstance(item, t) for t in item_types):
                    errmsg = (
                        f"Configuration file error: '{config_opt}' must be a {collection_type.__name__} "
                        f"of types {item_types_str}; found a '{type(item).__name__}' instead"
                    )
                    raise typer.BadParameter(errmsg)
        else:
            for item in items:
                if not isinstance(item, item_type):
                    errmsg = (
                        f"Configuration file error: '{config_opt}' must be a {collection_type.__name__} "
                        f"of type {item_type.__name__}; found a '{type(item).__name__}' instead"
                    )
                    raise typer.BadParameter(errmsg)


######################################################################
# Commands


@app.command()
def export(  # noqa: PLR0913, PLR0917
    ctx: typer.Context,
    folder_name: Annotated[
        list[str],
        typer.Argument(
            hidden=True,
            metavar=f"'{EXPORT_FOLDER_MSGFOLDERROOT}'|FOLDER ...",
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
    not_folder: Annotated[
        list[str] | None,
        typer.Option(
            "--not-folder",
            "-n",
            help="Exclude folder from export; option can be used more than once.",
            metavar="FOLDER",
        ),
    ] = None,
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
    notes_mkdir: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            help="Create mail folder hierarchy in notes directory.",
        ),
    ] = False,
    with_attachments: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--attachments/--no-attachments",
            help="Download and link attachments to note.",
        ),
    ] = True,
    attachments_folder: Annotated[
        str,
        typer.Option(
            hidden=False,
            exists=False,
            dir_okay=True,
            file_okay=False,
            writable=True,
            metavar=f"'{ATTACHMENTS_FOLDER_VAULT}'|'{ATTACHMENTS_FOLDER_NOTE}'|NAME|DIRECTORY",
            help="Store attachments in the vault, same folder as notes, a subfolder of notes, or another directory.",
            callback=_opt_attachments_folder_callback,
        ),
    ] = ATTACHMENTS_DEFAULT_FOLDER,
    attachments_mkdir: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            help="Create mail folder hierarchy in attachments directory when outside the vault.",
        ),
    ] = False,
    maxdepth: Annotated[
        int,
        typer.Option(
            "--maxdepth",
            min=1,
            help="Maximum depth of folder tree to export.",
        ),
    ] = 1,
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
    config_file_not_folder: list[str] = config_file_export_opts.get("not_folder", [])

    # Simple YAML syntax changes can result in values being interpreted as the wrong type!
    _validate_config_file_collection("tag", config_file_tags, list, str)
    _validate_config_file_collection("body_tag", config_file_body_tags, list, str)
    _validate_config_file_collection("metaprop", config_file_metaprop, dict, str | int | bool)
    _validate_config_file_collection("filter", config_file_filters, list, str)
    _validate_config_file_collection("search", config_file_search, list, str)
    _validate_config_file_collection("not_folder", config_file_not_folder, list, str)

    note_tags = config_file_tags + (tag or [])
    note_body_tags = config_file_body_tags + (body_tag or [])
    note_metaprop = _merge_metaprops(config_file_metaprop, ctx_obj.export_metaprops)
    not_folder = config_file_not_folder + (not_folder or [])
    not_folder_set = {nf.lower() for nf in not_folder}

    filter_query, search_query = _build_filter_query(
        config_file_filters + (message_filter or []),
        filter_cond,
        config_file_search + (message_search or []),
        search_cond,
        filter_test=filter_test,
    )

    vault_dir = vault_dir.absolute()

    try:
        graph_config = get_graph_config(common_opts)

        export_opts = ExportOptions(
            attachments_folder=attachments_folder,
            attachments_mkdir=attachments_mkdir,
            dry_run=dry_run,
            filter_query=filter_query,
            filter_test=filter_test,
            notes_mkdir=notes_mkdir,
            note_body_tags=note_body_tags or None,
            note_metaprop=note_metaprop or None,
            note_tags=note_tags or None,
            notes_dir=vault_dir,
            place_tags_body_end=place_tags_body_end,
            place_tags_frontmatter=place_tags_frontmatter,
            search_query=search_query,
            vault_dir=vault_dir,
            with_attachments=with_attachments,
        )

        # Are we exporting the entire mailbox?
        folder_is_msgfolderroot = EXPORT_FOLDER_MSGFOLDERROOT in {f.lower() for f in folder_name}

        with sessions.Session() as session:
            folders = graph_resolve_folder_paths(
                session,
                graph_config,
                MAILBOX_ROOT if folder_is_msgfolderroot else folder_name,
                maxdepth=maxdepth,
            )

        # Remove excluded folders
        folders = [f for f in folders if f["displayName"].lower() not in not_folder_set]

        if not filter_test:
            logger.debug("Querying Graph API for %d messages at a time", batch_size)
            console.print(f"Exporting messages from {len(folders)} folders...")

        with sessions.Session() as session:
            for folder in folders:
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
        print_graph_request_exception(err)
        raise typer.Exit(code=1) from err
    except GraphAuthRequiredError as err:
        err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
        raise typer.Exit(code=1) from err
    except (GraphApiError, ValueError, OSError) as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err
