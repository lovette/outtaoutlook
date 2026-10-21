"""outtaoutlook export command."""

import base64
import binascii
import concurrent.futures
import copy
import json
import logging
import mimetypes
import os
import sys
import threading
from collections import defaultdict
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Self, TypedDict

import requests
import typer
from requests import sessions
from rich.progress import BarColumn, Progress, ProgressColumn, SpinnerColumn, Task, TaskID, TextColumn
from rich.text import Text

from outtaoutlook.shared import (
    APP_NAME,
    AUTH_ERROR_USECOMMAND,
    GRAPH_API_URL,
    GRAPH_MESSAGE_FIELDS_SET,
    TITLE_UNTITLED,
    ExportAbortedError,
    ExportError,
    ExportFolderContext,
    ExportOptions,
    ExportStat,
    ExportStatsT,
    GraphApiConfig,
    GraphApiReturnT,
    GraphAttachmentT,
    GraphAuthRequiredError,
    GraphFolderT,
    GraphMessageT,
    OutputFormatChoice,
    abort_workers_event,
    console,
    datetime_now,
    exportstat_incr,
    get_hash_filename,
    graph_get_resource,
    log_message_notice,
    print_graph_request_exception,
    safe_file_name,
    unique_file_path,
)

from .output_eml import ExportFormatEml
from .output_html import ExportFormatHtml
from .output_json import ExportFormatJson
from .output_md import ExportFormatMarkdown
from .shared import ATTACHMENTS_DEFAULT_FOLDER, ATTACHMENTS_FOLDER_FILE, ATTACHMENTS_FOLDER_ROOT, MAX_EXPORT_WORKERS

if TYPE_CHECKING:
    from collections.abc import Generator

    from .output_protocol import OutputFormatModule

######################################################################
# Global variables

logger = logging.getLogger(APP_NAME)  # initialized in __main__

# Thread-local storage for per-thread data.
thread_local_data = threading.local()

# Each thread is assigned a unique counter ID to track its order of execution.
thread_id_counter_lock = threading.Lock()
thread_id_counter = 0

# Keep track of whether the first thread spawned has received its first batch of messages.
first_batch_complete_event = threading.Event()


######################################################################
# Constants

MIN_PROGRESS_THRESHOLD = 10  # Hide progress if only a few messages
FIRST_BATCH_COMPLETE_EVENT_WAIT = 60  # seconds; just a failsafe, I don't expect there to be any deadlocking :)
FOLDER_ICON = "\N{OPEN FILE FOLDER}\N{NARROW NO-BREAK SPACE}"

ATTACHMENT_DEFAULT_SUFFIX = ".bin"
ATTACHMENT_DEFAULT_NAME = "unnamed"  # suffix will be guessed

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


EXPECTED_MESSAGE_FIELDS_SET = GRAPH_MESSAGE_FIELDS_SET | {"@odata.etag"}
MESSAGE_DATE_FIELDS = ["createdDateTime", "lastModifiedDateTime", "sentDateTime", "receivedDateTime"]

EXPORT_MODULES: dict[OutputFormatChoice, type[OutputFormatModule]] = {
    OutputFormatChoice.eml: ExportFormatEml,
    OutputFormatChoice.md: ExportFormatMarkdown,
    OutputFormatChoice.html: ExportFormatHtml,
    OutputFormatChoice.json: ExportFormatJson,
}


######################################################################
# Types and classes


class FilterJoinCond(StrEnum):
    """Choices for joining filter conditions."""

    cond_and = "and"
    cond_or = "or"


@dataclass
class ExportResult:
    """Results of folder export."""

    aborted: bool = False  # Whether the export was aborted.
    messages: defaultdict[int, list[str]] = field(default_factory=lambda: defaultdict(list))  # {level:[message]}
    stats: ExportStatsT | None = None  # Export statistics keyed by ExportStat.


class RemainingItemsColumn(ProgressColumn):
    """Renders remaining count/total, e.g. ' 90/100 remaining'."""

    def render(self: Self, task: Task) -> Text:  # noqa: D102
        # Prevent math bugs if total isn't set yet
        total = f"{int(task.total):,}" if task.total is not None else "--"
        remaining = f"{int(task.remaining):,}" if task.remaining is not None else "--"

        return Text(f"{remaining}/{total}", style="progress.remaining")


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


def _export_mkdir(
    dir_path: Path,
    name: str,
    *,
    parents: bool = False,
    exist_ok: bool = True,
    export_opts: ExportOptions,
) -> None:
    """Create an export related directory if it doesn't exist.

    Args:
        dir_path (Path): Directory path.
        name (str): Purpose of the directory (used for logging.)
        export_opts (ExportOptions): Export options.
        parents (bool, optional): Whether to create parent directories if they don't exist. Defaults to False.
        exist_ok (bool, optional): Whether to ignore the error if the directory already exists. Defaults to True.
    """
    if not dir_path.is_dir():
        if not export_opts.dry_run:
            try:
                dir_path.mkdir(parents=parents, exist_ok=exist_ok)
            except OSError as err:
                logger.error("Failed to create %s directory: %s", name, str(err))  # noqa: TRY400
                raise

            logger.info("Created %s directory '%s'", name, dir_path)
        else:
            logger.info("Created %s directory '%s' (dry run)", name, dir_path)


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
        log_message_notice(message, f'attachment("{name}"): {errmsg}', level)

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
                _warn_attachment(
                    name,
                    f"Content type for '{item_data_type}' is '{content_type}', expected 'text/plain'",
                    logging.DEBUG,
                )
            content_type = "message/rfc822"  # guess_extension() -> '.eml'
        elif item_data_type == "#microsoft.graph.contact":
            if not content_type or not content_type.endswith("vcard"):
                _warn_attachment(
                    name,
                    f"Content type for '{item_data_type}' is '{content_type}', expected 'text/vcard'",
                    logging.DEBUG,
                )
            content_type = "text/x-vcard"  # guess_extension() -> '.vcf'
        else:
            _warn_attachment(name, f"Unsupported item data type '{item_data_type}'")

        return content, content_type

    name = Path(safe_file_name(attachment.get("name", ""), ATTACHMENT_DEFAULT_NAME))
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
    file_date: datetime,
    export_opts: ExportOptions,
) -> Path | None:
    """Save a message attachment to the local filesystem.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.
        attachment (GraphAttachmentT): Attachment properties.
        attachment_name_prefix (str): Prefix to prepend to attachment filenames. Must be filename safe.
        file_date (datetime): The date of the file.
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

    # Create attachments directory the first time an attachment is saved.
    _export_mkdir(
        export_opts.attachments_folder,
        "attachments",
        parents=export_opts.attachments_mkdir,
        export_opts=export_opts,
    )

    attachment_path = unique_file_path(export_opts.attachments_folder / Path(attachment["name"]), file_date)

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
            log_message_notice(message, f'attachment("{attachment_path.name}"): b64decode failed: {err}', logging.ERROR)
            return None

    if attachment_path.is_file():
        os.utime(attachment_path, (file_date.timestamp(), file_date.timestamp()))

    return attachment_path


def _save_attachments(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    attachment_name_prefix: str,
    file_date: datetime,
    export_opts: ExportOptions,
) -> list[Path]:
    """Save message attachments to the local filesystem.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): Message properties.
        attachment_name_prefix (str): The prefix for attachment file names.
        file_date (datetime): The date of the file.
        export_opts (ExportOptions): Export options.

    Returns:
        list[Path]: Paths relative to root output directory or URIs (file://) for absolute paths.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    attachment_paths: list[Path] = []

    for attachment in _list_attachments(session, graph_config, message):
        attachment_path = _save_attachment(
            session,
            graph_config,
            message,
            attachment,
            attachment_name_prefix,
            file_date,
            export_opts,
        )

        if attachment_path:
            attachment_paths.append(attachment_path)

    return attachment_paths


def _export_message(  # noqa: C901, PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    message: GraphMessageT,
    output_module: OutputFormatModule,
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> int:
    """Export a message.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        message (GraphMessageT): The message.
        output_module (OutputFormatModule): The output module to use for exporting.
        export_opts (ExportOptions): Export options.
        export_stats (ExportStatsT): Export statistics dictionary to update.

    Returns:
        int: The number of exported items. May be 0 if messages are batched by output module.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _fix_dates(message: GraphMessageT, is_draft: bool) -> None:  # noqa: FBT001
        dt_now = datetime_now()

        for date_field in ("createdDateTime", "lastModifiedDateTime", "receivedDateTime", "sentDateTime"):
            value = message.get(date_field)
            dt = _coerce_datetime(value)

            if not dt:
                if date_field not in message:
                    log_message_notice(message, f"field({date_field}): missing date field")
                elif is_draft and date_field in {"sentDateTime", "receivedDateTime"}:
                    pass  # Drafts aren't expected to have sent or received dates, we'll set something below
                elif value:
                    log_message_notice(message, f"field({date_field}): unexpected date format: '{value}'")
                else:
                    log_message_notice(message, f"field({date_field}): no value")

            # Ensure all fields are set to a valid datetime
            message[date_field] = dt or dt_now

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

    message_id: str | None = message.get("id")
    subject: str = message.get("subject") or TITLE_UNTITLED
    title = safe_file_name(subject, TITLE_UNTITLED)
    is_draft: bool | None = message.get("isDraft")
    has_attachments: bool | None = message.get("hasAttachments")
    data_type: str | None = message.get("@odata.type")
    is_calendar_related = data_type and data_type in MESSAGE_DATATYPE_EVENT_MESSAGE_SET
    attachment_paths: list[Path] | None = None

    if not message_id:
        # I don't expect this to ever happen, but just in case.
        log_message_notice(message, "'id' field is not set!", logging.ERROR)
        return 0

    _fix_dates(message, bool(is_draft))

    file_date: datetime = message["receivedDateTime"] or message["sentDateTime"]
    if is_draft:
        file_date = message["lastModifiedDateTime"] or message["createdDateTime"]

    message["subject"] = subject
    message["fileTitle"] = title  # custom property
    message["messageDateTime"] = file_date  # custom property

    if is_calendar_related:
        # Calendar requests and responses are skipped
        skip_reason, level = ExportStat.STAT_SKIPPED_EVENT_MESSAGE, logging.DEBUG
        exportstat_incr(export_stats, skip_reason)
        log_message_notice(message, skip_reason.value, level)
        return 0  # don't write empty files

    file_path = output_module.get_file_path(message, export_opts)

    # Save attachments
    if has_attachments and export_opts.with_attachments:
        attachment_paths = _save_attachments(
            session,
            graph_config,
            message,
            file_path.stem,
            file_date,
            export_opts,
        )

        exportstat_incr(export_stats, ExportStat.STAT_ATTACHMENT_SAVED, len(attachment_paths))

    # Export the message in specified format.
    return output_module.export_message(
        message,
        file_path,
        attachment_paths,
        export_opts,
        export_stats,
    )


def _export_folder_batch(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    progress: Progress,
    progress_task_id: TaskID,
    first_batch_data: GraphApiReturnT | None,
    url: str,
    params: dict[str, str | int] | None,
    batch_limit: int,
    exported_ids: set[str],
    output_module: OutputFormatModule,
    export_opts: ExportOptions,
    export_stats: ExportStatsT,
) -> tuple[int, str | None]:
    """Export a batch of messages from a specific folder.

    Args:
        session (sessions.Session): The requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        progress (rich.Progress): The progress bar to update during the export process.
        progress_task_id (TaskID): The task ID for updating the progress bar.
        first_batch_data (GraphApiReturnT | None): First batch of messages or None to query `url`.
        url (str): The URL for the API request.
        params (dict[str, str | int] | None): The query parameters for the API request.
        batch_limit (int): The maximum number of messages to export in this batch.
        exported_ids (set[str]): The set of message IDs that have already been exported. Modified in-place.
        output_module (OutputFormatModule): The output module to use for exporting.
        export_opts (ExportOptions): Options for exporting messages.
        export_stats (ExportStatsT): Export statistics dictionary to update.

    Returns:
        tuple[int, str | None]: [Number of messages exported in this batch (including dups),
            Link to next batch of messages or None if there are no more batches]

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    progress.update(progress_task_id, status="fetching messages", refresh=True)

    data = first_batch_data or _list_messages(session, graph_config, url, params=params)

    messages = data.get("value", [])
    next_url = data.get("@odata.nextLink")

    progress.update(progress_task_id, status="writing files", refresh=True)

    for message in messages:
        if abort_workers_event.is_set():
            raise ExportAbortedError

        n_exported = 1  # count processed even if in manifest

        if message["id"] not in exported_ids:
            n_exported = _export_message(
                session,
                graph_config,
                message,
                output_module,
                export_opts,
                export_stats,
            )

            exported_ids.add(message["id"])
        else:
            exportstat_incr(export_stats, ExportStat.STAT_SKIPPED_DUPLICATE)

        if n_exported:
            progress.update(progress_task_id, advance=n_exported)

        batch_limit -= 1
        if batch_limit <= 0:
            break

    return len(messages), next_url


def _export_folder(  # noqa: C901, PLR0912, PLR0915
    export_context: ExportFolderContext,
    output_module: OutputFormatModule,
) -> ExportResult:
    """Export messages from a specific folder to the output directory.

    Args:
        export_context (ExportFolderContext): Context for the folder being exported.
        output_module (OutputFormatModule): The output module to use for exporting.

    Returns:
        ExportResult: The result of the export operation, including statistics and any filter test message.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

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

    def _emit_log(errmsg: str, level: int = logging.INFO) -> None:
        logger.log(level, 'Folder("%s"): %s', folder["folderPath"], errmsg)

    folder = export_context.folder
    session = export_context.session
    progress = export_context.progress
    progress_task_id = export_context.progress_task_id
    graph_config = export_context.graph_config
    export_opts = export_context.export_opts
    export_stats = export_context.export_stats  # update context as progress is made

    folder_message_count = folder["totalItemCount"]
    export_result = ExportResult(stats=export_stats)

    if not folder_message_count:
        progress.update(progress_task_id, total=0, completed=0, status="empty folder", visible=False)
        export_result.messages[logging.INFO].append("Folder is empty")
        return export_result

    progress.update(progress_task_id, status="evaluating", refresh=True)

    manifest_ids: set[str] = set()
    manifest_file_name = get_hash_filename(folder["id"], "export_manifest", ".json")
    manifest_file_path = (export_opts.manifest_dir / manifest_file_name) if export_opts.manifest_dir else None

    if manifest_file_path and manifest_file_path.is_file():
        logger.info("Loading export manifest: '%s'", manifest_file_path)
        manifest_data = json.loads(manifest_file_path.read_text())
        manifest_ids = set(manifest_data.get("messageImmutableIds", []))

    export_limit = export_opts.files_limit
    exported_total_count = len(manifest_ids)
    files_dir = export_opts.files_dir
    attachments_dir_path = Path(export_opts.attachments_folder)  # assume it's an absolute path to start
    attachments_mkdir = export_opts.attachments_mkdir

    if export_opts.files_mkdir:
        files_dir = files_dir / (
            Path(folder["folderPath"])
            if not export_opts.mkdir_maxdepth
            else Path(*Path(folder["folderPath"]).parts[-export_opts.mkdir_maxdepth :])
        )

    if isinstance(export_opts.attachments_folder, Path):
        assert attachments_dir_path.is_dir()  # noqa: S101 ; caller should have verified!
        assert attachments_dir_path.is_absolute()  # noqa: S101 ; caller should have verified!
    elif export_opts.attachments_folder.upper() == ATTACHMENTS_FOLDER_ROOT:
        attachments_dir_path = export_opts.root_out_dir
        attachments_mkdir = False
    elif export_opts.attachments_folder.upper().startswith(ATTACHMENTS_FOLDER_FILE):  # allow plural "FILES"
        attachments_dir_path = files_dir
        attachments_mkdir = False
    elif not attachments_dir_path.is_absolute():
        attachments_dir_path = files_dir / safe_file_name(export_opts.attachments_folder, ATTACHMENTS_DEFAULT_FOLDER)
        attachments_mkdir = False

    if attachments_mkdir:
        attachments_dir_path = attachments_dir_path / (
            Path(folder["folderPath"])
            if not export_opts.mkdir_maxdepth
            else Path(*Path(folder["folderPath"]).parts[-export_opts.mkdir_maxdepth :])
        )

    if not export_opts.filter_test:
        _emit_log(f"Saving files to '{files_dir}'")
        _export_mkdir(files_dir, "files", parents=export_opts.files_mkdir, export_opts=export_opts)

        if export_opts.with_attachments:
            _emit_log(f"Saving attachments to '{attachments_dir_path}'")
            # Attachments directory will be created the first time an attachment is saved.

    # Tweak export options without modifying caller
    export_opts = copy.replace(
        export_opts,
        files_dir=files_dir,
        attachments_mkdir=attachments_mkdir,
        attachments_folder=attachments_dir_path,
    )

    export_opts = output_module.set_export_opts(export_opts)

    if export_opts.with_attachments:
        export_stats[ExportStat.STAT_ATTACHMENT_SAVED] = 0

    # https://learn.microsoft.com/en-us/graph/query-parameters
    params: dict[str, str | int] | None = {
        "$count": "true",
    }

    if export_opts.filter_query:
        params["$top"] = 1  # We're only interested in the message $count
        params["$filter"] = export_opts.filter_query
        exported_total_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

    elif export_opts.search_query:
        params["$top"] = 1  # We're only interested in the message $count
        params["$search"] = export_opts.search_query
        exported_total_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

    else:
        # The Graph API imposes rules when using $orderby and $filter together.
        # $search and $orderby cannot be used together.
        # Orderby multiple fields so order is more deterministic.
        params["$top"] = export_opts.batch_size
        params["$orderby"] = "lastModifiedDateTime,internetMessageId,subject"

    if exported_total_count and not export_opts.filter_test:
        params["$skip"] = exported_total_count  # skip already exported messages
        exportstat_incr(export_stats, ExportStat.STAT_SKIPPED_DUPLICATE, incr=exported_total_count)

    output_module.update_list_messages_params(params)

    if export_opts.filter_test:
        params["$select"] = "id"  # Might as well make the response shorter; this *overrides* the output module!

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
            progress.print(
                f"[red]Error: Failed to list messages in folder '{folder['folderPath']}', "
                f"possibly due to '{field}' criteria syntax errors.",
            )

            _emit_log("Aborting due to filter/search criteria error", logging.DEBUG)

        # If the first batch query fails due to filter or search criteria,
        # there's no reason for other workers to continue.
        raise

    # Now that we know that the first batch of messages has been
    # successfully retrieved, we can release the other workers.
    _release_workers()

    first_batch_messages = first_batch_data.get("value", [])

    # I'm guessing '@odata.count' is not provided when search results are less than '$top'.
    export_message_count: int = first_batch_data.get("@odata.count") or len(first_batch_messages)

    if export_opts.filter_test:
        # Report the filtered message count next to task progress.
        progress.update(
            progress_task_id,
            total=1,
            completed=1,
            status=f"{export_message_count} messages",
        )

        return export_result

    if not export_message_count:
        if export_opts.filter_query:
            msg, level = "No messages match 'filter' criteria", logging.INFO
        elif export_opts.search_query:
            msg, level = "No messages match 'search' criteria", logging.INFO
        else:
            msg, level = f"Contains {folder_message_count} messages but Graph API result was empty", logging.ERROR

        export_result.messages[level].append(msg)

        return export_result

    if not export_limit:
        export_limit = export_message_count

    status_message = ["Exporting"]

    if export_limit == folder_message_count:
        status_message.append(f"all {folder_message_count}")
    elif export_opts.filter_query or export_opts.search_query:
        status_message.append(f"{export_message_count} (filtered) of {folder_message_count}")
    else:
        status_message.append(f"{export_limit} of {export_message_count}")

    status_message.append("messages")

    if manifest_file_path:
        status_message.append("(skipping dups)")
    if export_opts.dry_run:
        status_message.append("(dry run)")

    _emit_log(" ".join(status_message))

    progress.update(
        progress_task_id,
        total=min(export_limit, export_message_count),
        completed=exported_total_count,
    )

    if True:
        while next_url is not None and exported_total_count < export_limit:
            if abort_workers_event.is_set():
                raise ExportAbortedError

            try:
                batch_total, next_url = _export_folder_batch(
                    session,
                    graph_config,
                    progress,
                    progress_task_id,
                    first_batch_data,
                    next_url,  # ignored when `first_batch_data` is set
                    params,
                    min(export_opts.batch_size, export_limit - exported_total_count),
                    manifest_ids,
                    output_module,
                    export_opts,
                    export_stats,
                )

            except requests.RequestException, ExportError, ValueError, OSError:
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

    # Let output module finish up any messages it may have queued.
    n_exported = output_module.last_batch_complete()

    progress.update(progress_task_id, advance=n_exported, status="done")

    return export_result


def _signal_abort_workers() -> None:
    """Set event that signals all workers to stop exporting messages."""
    if not abort_workers_event.is_set():
        logger.debug("Setting all workers abort flag")
        abort_workers_event.set()


def _release_workers() -> None:
    if thread_local_data.thread_id == 1 and not first_batch_complete_event.is_set():
        logger.debug("Releasing workers")
        first_batch_complete_event.set()


def _export_folder_worker(  # noqa: C901
    folder: GraphFolderT,
    *,
    progress: Progress,
    progress_task_id: TaskID,
    graph_config: GraphApiConfig,
    export_opts: ExportOptions,
) -> ExportResult:
    """Export folder worker thread entry point.

    Args:
        folder (GraphFolderT): Folder to export.
        progress (Progress): Progress bar.
        progress_task_id (TaskID): The task ID for the progress tracker.
        graph_config (GraphApiConfig): Graph API configuration.
        export_opts (ExportOptions): Options for exporting messages.

    Raises:
        Exceptions are not captured and will bubble up to the main thread by `future.result`.

    Returns:
        ExportResult: The result of the export operation, including statistics and any status message.
    """

    def _emit_log(errmsg: str, level: int = logging.INFO) -> None:
        logger.log(level, 'Folder("%s"): %s', folder["folderPath"], errmsg)

    def _get_thread_id() -> int:
        global thread_id_counter  # noqa: PLW0603

        if not getattr(thread_local_data, "thread_id", None):
            with thread_id_counter_lock:
                thread_id_counter += 1
                thread_local_data.thread_id = thread_id_counter

            new_name = f"Thread #{thread_local_data.thread_id}"
            logger.debug("Setting thread name to '%s'", new_name)
            threading.current_thread().name = new_name

        return thread_local_data.thread_id

    def _wait_first_batch_release() -> bool:
        if thread_id > 1 and not first_batch_complete_event.is_set():
            _emit_log("Waiting for first batch release", logging.DEBUG)

            if not first_batch_complete_event.wait(timeout=FIRST_BATCH_COMPLETE_EVENT_WAIT):
                _emit_log("Timed out waiting for release.", logging.DEBUG)
                _signal_abort_workers()
                return False

        return True

    thread_id = _get_thread_id()

    # Wait for the first batch to pass its checkpoint if this is not the primary thread.
    if not _wait_first_batch_release():
        return ExportResult(aborted=True)

    # If one worker aborted, all the others abort too.
    if abort_workers_event.is_set():
        return ExportResult(aborted=True)

    _emit_log("Working on this folder", logging.DEBUG)

    # Reset task start time to get true elapsed time for this folder.
    progress.reset(progress_task_id, start=True)

    output_module_type = EXPORT_MODULES[export_opts.out_format]

    export_stats: ExportStatsT = {
        # Start with stats that will always be output
        ExportStat.STAT_FILE_CREATED: 0,
    }

    export_result = ExportResult(stats=export_stats)

    try:
        with sessions.Session() as session:
            export_context = ExportFolderContext(
                folder=folder,
                session=session,
                graph_config=graph_config,
                progress=progress,
                progress_task_id=progress_task_id,
                export_opts=export_opts,
                export_stats=export_stats,  # updated as progress is made
            )

            return _export_folder(export_context, output_module_type(export_context))

    except requests.RequestException as err:
        print_graph_request_exception(progress.console, err)
        _emit_log("Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    except GraphAuthRequiredError as err:
        # This should have triggered in main thread when calling `get_graph_config`...
        logger.error("%s %s", str(err), AUTH_ERROR_USECOMMAND)  # noqa: TRY400
        _emit_log("Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    except ExportAbortedError:
        # Whichever thread signaled the abort already logged the reason...
        export_result.aborted = True
    except (ExportError, ValueError, OSError) as err:
        logger.error("%s", str(err))  # noqa: TRY400
        _emit_log("Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    finally:
        # Primary thread must always release workers!
        _release_workers()

    # The failure of one worker brings everything to a halt.
    if export_result.aborted:
        _signal_abort_workers()

    return export_result


def export_folders(  # noqa: C901, PLR0912, PLR0915
    folders: list[GraphFolderT],
    graph_config: GraphApiConfig,
    export_opts: ExportOptions,
    max_threads: int | None,
) -> None:
    """Export a list of folders, each in its own thread.

    Args:
        folders (list[GraphFolderT]): List of folder objects.
        graph_config (GraphApiConfig): Graph API configuration.
        export_opts (ExportOptions): Options for exporting messages.
        max_threads (int or None): Maximum size of thread pool or None to decide based on hardware.
    """

    def _print_export_results(results: dict[str, tuple[GraphFolderT, ExportResult | None]]) -> None:  # noqa: C901
        # Filter export_results to items with messages or stats
        results = {
            folder_id: (folder, export_result)
            for folder_id, (folder, export_result) in results.items()
            if export_result and (export_result.messages or export_result.stats)
        }

        if not results:
            return

        console.print()
        console.print("[bold]Results:")

        for folder, export_result in results.values():
            if export_result and (export_result.messages or export_result.stats):
                console.print()
                console.print(f"[bold]'{folder['folderPath']}'[/bold]:")

                if export_result.aborted:
                    console.print("[red]- Export did not complete[/red]")

                for level, messages in export_result.messages.items():
                    for message in messages:
                        if level == logging.ERROR:
                            console.print(f"[red]- {message}[/red]")
                        elif level == logging.WARNING:
                            console.print(f"[yellow]- {message}[/yellow]")
                        else:
                            console.print(f"- {message}")

                if export_result.stats:
                    for stat, stat_count in export_result.stats.items():
                        console.print(f"- {stat}: {stat_count:,}")

    # Keyword arguments to `_export_folder_worker`.
    class _WorkerArgs(TypedDict):
        progress: Progress
        progress_task_id: TaskID
        graph_config: GraphApiConfig
        export_opts: ExportOptions

    show_progress = True

    if export_opts.filter_test:
        progress_cols = [
            TextColumn("[bold]{task.description}"),
            SpinnerColumn(finished_text=""),
            TextColumn("{task.fields[status]}"),  # our custom field
        ]
    else:
        progress_cols = [
            TextColumn("{task.description}"),
            BarColumn(),
            SpinnerColumn(),
            RemainingItemsColumn(),  # TextColumn("{task.remaining:g}"),
            TextColumn("{task.fields[status]}"),  # our custom field
        ]

    export_results: dict[str, tuple[GraphFolderT, ExportResult | None]] = {}  # {id:(folder, result)}

    if max_threads and max_threads >= MAX_EXPORT_WORKERS:
        max_threads = None  # let ThreadPoolExecutor decide

    with (
        concurrent.futures.ThreadPoolExecutor(max_threads) as executor,
        _logger_friendly_progress(*progress_cols, console=console, logger=logger, disable=not show_progress) as progress,
    ):
        folder_task_id: dict[str, tuple[GraphFolderT, TaskID]] = {}  # {id: (folder,task_id)}
        workers: dict[concurrent.futures.Future[ExportResult], GraphFolderT] = {}
        future_task_id: dict[concurrent.futures.Future[ExportResult], TaskID] = {}

        logger.debug("Worker thread pool size: %d", executor._max_workers)  # noqa: SLF001

        for folder in folders:
            # Add tasks to progress bar *before* starting.
            # This keeps the task bars in hierarchical order and ensures
            # console output doesn't interrupt the progress display.
            progress_task_id = progress.add_task(
                f"{FOLDER_ICON}{folder['folderPath']}",
                start=False,
                total=None,
                visible=True,
                status="pending",
            )

            worker_args: _WorkerArgs = {
                "progress": progress,
                "progress_task_id": progress_task_id,
                "graph_config": graph_config,
                "export_opts": export_opts,
            }

            future = executor.submit(_export_folder_worker, folder, **worker_args)

            workers[future] = folder
            future_task_id[future] = progress_task_id
            folder_task_id[folder["id"]] = (folder, progress_task_id)
            export_results[folder["id"]] = (folder, None)

        shutdown_pending = False
        progress_tasks: dict[TaskID, Task] = {}
        progress_tasks = {task.id: task for task in progress.tasks}

        # We loop through workers sequentially instead of using `as_completed` because it hangs after we call `shutdown` :/
        # https://tiewkeehui.cc/blog/python-thread-pool-executor/
        # https://stackoverflow.com/questions/66923690/does-concurrent-futures-as-completed-yield-for-cancelled-futures

        logger.debug("Starting worker threads.")

        # Start timer now so we can turn off the spinner if aborted.
        for progress_task_id in progress.task_ids:
            progress.start_task(progress_task_id)

        for worker, folder in workers.items():
            progress_task_id = future_task_id[worker]

            if worker.cancelled():
                logger.debug('Folder("%s"): Future was cancelled', folder["folderPath"])
                progress.update(progress_task_id, status="Cancelled")
                continue

            worker_exception = False

            try:
                export_result = worker.result()
            except concurrent.futures.CancelledError:
                # This will never happen because we skip cancelled futures before calling result()
                worker_exception = True
            else:
                export_results[folder["id"]] = (folder, export_result)

                result_type = [
                    "aborted" if export_result.aborted else None,
                    "stats" if export_result.stats else None,
                    "messages" if export_result.messages else None,
                ]

                logger.debug(
                    'Folder("%s"): Worker result is available (%s)',
                    folder["folderPath"],
                    ",".join(filter(None, result_type)),
                )

                if export_result.aborted:
                    task = progress_tasks[progress_task_id]
                    progress.update(progress_task_id, completed=task.total, status="cancelled")

            if worker_exception and not abort_workers_event.is_set():
                logger.debug('Folder("%s"): Worker resulted in an exception.', folder["folderPath"])
                _signal_abort_workers()

            if abort_workers_event.is_set() and not shutdown_pending:
                # Cancel any workers that have not run yet
                logger.debug('Folder("%s"): Abort signal is set, shutting down worker threads.', folder["folderPath"])
                executor.shutdown(wait=False, cancel_futures=True)
                shutdown_pending = True

    logger.debug("All worker threads finished.")

    _print_export_results(export_results)

    if abort_workers_event.is_set():
        raise typer.Exit(code=1)
