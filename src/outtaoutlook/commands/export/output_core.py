"""outtaoutlook export folder worker core."""

import base64
import binascii
import itertools
import json
import logging
import mimetypes
import os
import threading
from abc import ABC, abstractmethod
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from functools import cached_property
from pathlib import Path
from typing import TYPE_CHECKING, Any, Self

import requests
from requests import Session

from outtaoutlook.graphapi import (
    GRAPH_API_URL,
    GRAPH_MESSAGE_FIELDS_SET,
    GraphApiConfig,
    GraphApiReturnT,
    GraphAttachmentT,
    GraphFolderT,
    GraphMessageT,
    graph_get_resource,
    log_message_notice,
)
from outtaoutlook.shared import (
    ExportCancelledError,
    ExportError,
    datetime_now,
    get_hash_filename,
    logger,
    safe_file_name,
    unique_file_path,
)

from .shared import (
    ATTACHMENTS_DEFAULT_FOLDER,
    ATTACHMENTS_FOLDER_FILE,
    ATTACHMENTS_FOLDER_ROOT,
    ExportOptions,
    json_serializer,
)

if TYPE_CHECKING:
    from collections.abc import Sequence

    from rich.progress import Progress, Task, TaskID


######################################################################
# Constants


TITLE_UNTITLED = "Untitled"
MIN_PROGRESS_THRESHOLD = 10  # Hide progress if only a few messages
LIST_MESSAGES_SUCCESS_EVENT_WAIT = 60  # seconds; just a failsafe, I don't expect there to be any deadlocking :)

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


######################################################################
# Types and classes


class ExportStat(StrEnum):
    """Export statistic messages."""

    # Results are output in the order they are defined here.
    STAT_MESSAGE_EXPORTED = "Messages exported"
    STAT_ATTACHMENT_SAVED = "Attachments saved"
    STAT_SKIPPED_DUPLICATE = "Skipped; Message already exported"
    STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED = "Skipped; Markdown conversion failed"
    STAT_SKIPPED_NO_MESSAGE_BODY = "Skipped; No message body"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS = "Skipped; No file content, just attachments (not being saved)"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS = "Skipped; No file content"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN = "Skipped; No file content, unknown reason"
    STAT_SKIPPED_EVENT_MESSAGE = "Skipped; Calendar event"


ExportStatsT = dict[ExportStat, int]


@dataclass
class ExportResult:
    """Results of folder export."""

    failed: bool = False  # Whether the export failed.
    cancelled: bool = False  # Whether the export was cancelled.
    messages: defaultdict[int, list[str]] = field(default_factory=lambda: defaultdict(list))  # {level:[message]}

    def add_message(self, message: str, level: int = logging.INFO) -> None:
        """Add a message to the export result.

        Args:
            message (str): Message text.
            level (int): Logging level (e.g., logging.INFO, logging.WARNING, logging.ERROR).
        """
        self.messages[level].append(message)


######################################################################
# Export folder worker core


class OutputWorkerCore(ABC):
    """Export folder worker core."""

    # Keep track of whether the first worker has received its first batch of messages.
    _messages_success_event = threading.Event()

    # Any worker can signal for all other workers to cancel exporting messages.
    _workers_cancelled_event = threading.Event()

    def __init__(
        self,
        folder: GraphFolderT,
        graph_config: GraphApiConfig,
        progress: Progress,
        progress_task_id: TaskID,
        export_opts: ExportOptions,
    ) -> None:
        """Initialize the export folder worker.

        Args:
            folder (GraphFolderT): Folder object.
            graph_config (GraphApiConfig): Graph API configuration.
            progress (Progress): Progress bar object.
            progress_task_id (TaskID): Progress bar task ID for this folder.
            export_opts (ExportOptions): Export options.
        """
        self.thread_id: int | None = None
        self.folder = folder
        self.session: Session | None = None
        self.graph_config = graph_config
        self.progress = progress
        self.progress_task_id = progress_task_id
        self.export_opts = export_opts
        self.export_stats: ExportStatsT = {}
        self.export_result = ExportResult()
        self.manifest_ids: set[str] = set()
        self.write_messages_status = "writing files"

    ######################################################################
    # Helper methods

    @staticmethod
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

    def _mkdir(
        self,
        dir_path: Path,
        purpose: str,
        *,
        parents: bool = False,
        exist_ok: bool = True,
    ) -> None:
        """Create an export related directory if it doesn't exist.

        Args:
            dir_path (Path): Directory path.
            purpose (str): Purpose of the directory (used for logging.)
            parents (bool, optional): Whether to create parent directories if they don't exist. Defaults to False.
            exist_ok (bool, optional): Whether to ignore the error if the directory already exists. Defaults to True.
        """
        if not dir_path.is_dir():
            if not self.export_opts.dry_run:
                try:
                    dir_path.mkdir(parents=parents, exist_ok=exist_ok)
                except OSError as err:
                    logger.error("Failed to create %s directory: %s", purpose, str(err))
                    raise

                logger.info("Created %s directory '%s'", purpose, dir_path)
            else:
                logger.info("Created %s directory '%s' (dry run)", purpose, dir_path)

    def _list_messages(
        self,
        url: str,
        params: dict[str, Any] | None = None,
    ) -> GraphApiReturnT:
        """Get a list of messages.

        Args:
            url (str): Entity URL.
            params (dict, optional): Query parameters for the GET request. Defaults to None.

        Returns:
            GraphApiReturnT: JSON response from the Microsoft Graph API.

        Raises:
            ExportError if an export error occurred.
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        assert self.session is not None  # noqa: S101

        headers = {
            # https://learn.microsoft.com/en-us/graph/outlook-immutable-id
            "Prefer": 'IdType="ImmutableId"',
        }

        self.progress_update(status="fetching messages", refresh=True)

        try:
            data = graph_get_resource(
                self.session,
                self.graph_config,
                url,
                params=params,
                headers=headers,
            )

        except requests.HTTPError:
            if self.export_opts.filter_query:
                field = "filter"
            elif self.export_opts.search_query:
                field = "search"
            else:
                field = None

            if field:
                # Try to be somewhat helpful
                self.progress.print(
                    f"[red]Error: Failed to list messages in folder '{self.folder_path}', "
                    f"possibly due to '{field}' criteria syntax errors.",
                )

                self.emit_log("Cancelling export due to filter/search criteria error", logging.DEBUG)

            raise
        else:
            # Now that we know that the first batch of messages has been
            # successfully retrieved, we can release the other workers.
            self.release_workers()

        return data

    def _list_attachments(
        self,
        message: GraphMessageT,
    ) -> list[GraphAttachmentT]:
        """List properties for all message attachments.

        Args:
            message (GraphMessageT): Message properties.

        Returns:
            list[GraphAttachmentT]

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        # Best I can tell, '/attachments' returns all attachments and has no 'nextLink'
        assert self.session is not None  # noqa: S101
        url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments"
        data = graph_get_resource(self.session, self.graph_config, url)
        return data.get("value", [])

    def _get_attachment_content(  # noqa: C901, PLR0915
        self,
        message: GraphMessageT,
        attachment: GraphAttachmentT,
        attachment_name_prefix: str,
    ) -> tuple[str | None, str | None]:
        """Get attachment content and ensure 'name' property is appropriate and valid.

        Args:
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
            assert self.session is not None  # noqa: S101
            url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments/{attachment.get('id')}"
            params = {"$expand": "microsoft.graph.itemattachment/item"}
            data = graph_get_resource(self.session, self.graph_config, url, params=params)
            item = data.get("item", {})
            return item, item.get("@odata.type")

        def _get_item_value() -> tuple[str | None, str | None]:
            # Get raw attachment content for a contact, event or message.
            assert self.session is not None  # noqa: S101
            url = f"{GRAPH_API_URL}/me/messages/{message.get('id')}/attachments/{attachment.get('id')}/$value"
            data = graph_get_resource(self.session, self.graph_config, url)
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

    def _save_attachment(
        self,
        message: GraphMessageT,
        attachment: GraphAttachmentT,
        attachment_name_prefix: str,
        file_date: datetime,
    ) -> Path | None:
        """Save a message attachment to the local filesystem.

        Args:
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
        content, content_type = self._get_attachment_content(
            message,
            attachment,
            attachment_name_prefix,
        )

        if content is None:
            return None

        assert isinstance(self.export_opts.attachments_folder, Path)  # noqa: S101
        assert self.export_opts.attachments_folder.is_absolute()  # noqa: S101

        # Create attachments directory the first time an attachment is saved.
        self._mkdir(
            self.export_opts.attachments_folder,
            "attachments",
            parents=self.export_opts.attachments_mkdir,
        )

        attachment_path = unique_file_path(self.export_opts.attachments_folder / Path(attachment["name"]), file_date)

        if self.export_opts.dry_run:
            self.emit_log(f"Wrote attachment to '{attachment_path.name}' ({content_type}) (dry run)", logging.DEBUG)
        elif content_type and content_type.startswith(("text/", "message/")):
            attachment_path.write_text(content, encoding="utf-8")
            self.emit_log(f"Wrote attachment to '{attachment_path.name}' ({content_type})", logging.DEBUG)
        else:
            try:
                with attachment_path.open("wb") as f:
                    f.write(base64.b64decode(content))
                self.emit_log(f"Wrote attachment to '{attachment_path.name}' ({content_type})", logging.DEBUG)
            except binascii.Error as err:
                log_message_notice(message, f'attachment("{attachment_path.name}"): b64decode failed: {err}', logging.ERROR)
                return None

        if attachment_path.is_file():
            os.utime(attachment_path, (file_date.timestamp(), file_date.timestamp()))

        return attachment_path

    def _save_attachments(
        self,
        message: GraphMessageT,
        attachment_name_prefix: str,
        file_date: datetime,
    ) -> list[Path]:
        """Save message attachments to the local filesystem.

        Args:
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

        for attachment in self._list_attachments(message):
            self.raise_if_cancelled()

            attachment_path = self._save_attachment(
                message,
                attachment,
                attachment_name_prefix,
                file_date,
            )

            if attachment_path:
                attachment_paths.append(attachment_path)

        return attachment_paths

    def _export_write_message(  # noqa: C901
        self,
        message: GraphMessageT,
    ) -> int:
        """Prepares message and writes a message to a file (internal use only).

        Args:
            message (GraphMessageT): The message.

        Returns:
            int: The number of exported items. May be 0 if messages are processed in a batch.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _fix_dates(message: GraphMessageT, is_draft: bool) -> None:  # noqa: FBT001
            dt_now = datetime_now()

            for date_field in ("createdDateTime", "lastModifiedDateTime", "receivedDateTime", "sentDateTime"):
                value = message.get(date_field)
                dt = self._coerce_datetime(value)

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
            return 1

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
            self.increxportstat(skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return 1  # don't write empty files

        file_path = self.get_file_path(message)

        # Save attachments
        if has_attachments and self.export_opts.with_attachments:
            attachment_paths = self._save_attachments(
                message,
                file_path.stem,
                file_date,
            )

            self.increxportstat(ExportStat.STAT_ATTACHMENT_SAVED, len(attachment_paths))

        # Export the message in specified format.
        return self.export_write_message(
            message,
            file_path,
            attachment_paths,
        )

    def _export_folder_batch(
        self,
        first_batch_data: GraphApiReturnT | None,
        url: str,
        params: dict[str, str | int] | None,
        batch_limit: int,
    ) -> tuple[int, int, int, str | None]:
        """Export a batch of messages from a specific folder.

        Args:
            first_batch_data (GraphApiReturnT | None): First batch of messages or None to query `url`.
            url (str): The URL for the API request.
            params (dict[str, str | int] | None): The query parameters for the API request.
            batch_limit (int): The maximum number of messages to export in this batch.

        Returns:
            tuple[int, int, int, str | None]: [
                Number of messages processed in this batch,
                Number of duplicate messages in this batch,
                Number of messages actually written in this batch,
                Link to next batch of messages or None if there are no more batches
                ]

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        data = first_batch_data or self._list_messages(url, params=params)

        messages: list[GraphMessageT] = data.get("value", [])
        next_url: str | None = data.get("@odata.nextLink")

        self.progress_update(status=self.write_messages_status, refresh=True)

        dup_total = 0
        batch_total = 0
        written_total = 0

        for message in itertools.islice(messages, batch_limit):
            self.raise_if_cancelled()

            if message["id"] not in self.manifest_ids:
                # n_advance will be:
                # - 0 if no progress is made (e.g. message queued)
                # - 1 if a single message is written or skipped (e.g. calendar event)
                # - 20 if a batch of queued messages is written
                n_advance = self._export_write_message(message)
                written_total += n_advance

                self.manifest_ids.add(message["id"])

                # We dont't advance progress for dups because when we want to export 50 messages
                # and it takes 322 attempts due to duplicates, we show progress like "322/50".
                # The downside is when there are mostly duplicates, the progress bar doesn't move :/
                if n_advance:
                    self.progress_update(advance=n_advance)
            else:
                dup_total += 1
                self.increxportstat(ExportStat.STAT_SKIPPED_DUPLICATE)

            batch_total += 1

        return batch_total, dup_total, written_total, next_url

    def emit_log(self, errmsg: str | Sequence[str | None], level: int = logging.INFO) -> None:
        """Emit a log message related to the current folder.

        Args:
            errmsg (str | Sequence[str | None]): The error message to log or a sequence of error messages to join.
            level (int, optional): The logging level. Defaults to logging.INFO.
        """
        if not isinstance(errmsg, str):
            errmsg = " ".join(filter(None, errmsg)).strip()  # join non-None parts into a string
        if errmsg:
            logger.log(level, 'folder("%s"): %s', self.folder_path, errmsg)

    def set_failed(self, reason: str, *, raise_now: bool = False) -> None:
        """Mark the export as failed and update progress bar task status.

        Also releases workers waiting to be released.

        Args:
            reason (str): The reason for the failure.
            raise_now (bool, optional): True to raise ExportCancelledError immediately. Defaults to False.
        """
        reason = f"Export failed: {reason}"
        self.emit_log(reason, logging.DEBUG)

        # Show user in result summary
        self.export_result.add_message(reason)

        self.export_result.failed = True
        self.progress_update(status="failed")

        # If one worker fails, we bring everything to an end.
        self.cancel_workers(reason, raise_now=raise_now)

    def set_cancelled(self, reason: str, *, raise_now: bool = False) -> None:
        """Mark the export as cancelled and update progress bar task status.

        Args:
            reason (str): The reason for the cancellation.
            raise_now (bool, optional): True to raise ExportCancelledError immediately. Defaults to False.
        """
        reason = f"Export cancelled: {reason}"
        self.emit_log(reason, logging.DEBUG)

        self.export_result.cancelled = True
        self.progress_update(status="cancelled")

        if raise_now:
            raise ExportCancelledError

    def cancel_workers(self, reason: str | None = None, *, raise_now: bool = False) -> None:
        """Set event that signals all workers to stop exporting messages.

        Also releases workers waiting to be released.

        Args:
            reason (str | None): The reason for cancelling the workers.
            raise_now (bool, optional): True to raise ExportCancelledError immediately. Defaults to False.

        Raises:
            ExportCancelledError
        """
        if not self.workers_cancelled:
            reason = f": {reason}" if reason else "."
            self.emit_log(f"All workers cancelled{reason}", logging.DEBUG)

        self._workers_cancelled_event.set()

        # Ensure all workers are released.
        self.release_workers()

        if raise_now:
            raise ExportCancelledError

    def raise_if_cancelled(self) -> None:
        """If workers have been cancelled, set our export status to cancelled and raise ExportCancelledError."""
        if self.workers_cancelled:
            self.set_cancelled("All workers have been cancelled.", raise_now=True)

    def release_workers(self) -> None:
        """Signal to other workers that they can proceed with their work."""
        if not self._messages_success_event.is_set():
            logger.debug("Releasing other workers")
        self._messages_success_event.set()

    def _wait_for_release_workers(self) -> None:
        """Wait for first worker to release all the rest.

        Raises:
            ExportCancelledError: If the workers have been cancelled while waiting.
        """
        assert self.thread_id is not None  # noqa: S101

        if self.thread_id > 1 and not self._messages_success_event.is_set():
            self.emit_log("Waiting to be released", logging.DEBUG)

            if not self._messages_success_event.wait(timeout=LIST_MESSAGES_SUCCESS_EVENT_WAIT):
                self.set_failed("Timed out waiting for first worker to release workers.", raise_now=True)

        # If any worker signals a cancellation, all work comes to a stop.
        self.raise_if_cancelled()

    @cached_property
    def folder_id(self) -> str:
        """Get folder Outlook identifier.

        Identifiers are 120 characters in length and look like this:
        'AAMkAGIzZTczZTJhLTNiM2UtN...CAQAo-r499r1UTpQZ9f71nCKHAALmwWqBAAA='

        Returns:
            str
        """
        return self.folder["id"]

    @cached_property
    def folder_name(self) -> str:
        """Get folder name.

        Returns:
            str
        """
        return self.folder["displayName"]

    @cached_property
    def folder_path(self) -> Path:
        """Get folder path.

        Returns:
            Path
        """
        return Path(self.folder["folderPath"])

    @cached_property
    def progress_task(self) -> Task:
        """Get progress task associated with the folder.

        Returns:
            Task
        """
        progress_tasks = {task.id: task for task in self.progress.tasks}
        return progress_tasks[self.progress_task_id]

    @property
    def workers_cancelled(self) -> bool:
        """Get workers cancelled signal status.

        Returns:
            bool
        """
        return self._workers_cancelled_event.is_set()

    @property
    def has_results(self) -> bool:
        """Return True if export has any results to report.

        Returns:
            bool
        """
        return bool(self.export_result.messages) or bool(self.export_stats)

    @cached_property
    def manifest_path(self) -> Path | None:
        """Get manifest file path, or None if not using a manifest.

        Returns:
            Path | None
        """
        if not self.export_opts.manifest_dir:
            return None
        return self.export_opts.manifest_dir / get_hash_filename(self.folder_id, "export_manifest", ".json")

    def _read_manifest(self) -> dict[str, str] | None:
        """Read export manifest file if using a manifest."""
        manifest_data = None

        if self.manifest_path and self.manifest_path.is_file():
            manifest_data = json.loads(self.manifest_path.read_text())
            self.manifest_ids = set(manifest_data.get("messageImmutableIds", []))
            self.emit_log(
                f"Read {len(self.manifest_ids):,} message IDs from manifest: '{self.manifest_path}'",
                logging.DEBUG,
            )

        return manifest_data

    def _write_manifest(self, *, emitlog: bool = False) -> None:
        """Write export manifest file if using a manifest.

        Args:
            emitlog (bool, optional): Whether to emit a log message after writing the manifest. Defaults to False.
        """
        if self.manifest_path and self.manifest_ids:
            self.manifest_path.write_text(
                json.dumps(
                    {
                        "date": datetime_now().isoformat(),
                        "folderId": self.folder_id,
                        "folderPath": self.folder_path,
                        "filterQuery": self.export_opts.filter_query,
                        "searchQuery": self.export_opts.search_query,
                        "messageImmutableIds": list(self.manifest_ids),
                    },
                    default=json_serializer,
                ),
            )

            if emitlog:
                self.emit_log(
                    f"Wrote {len(self.manifest_ids):,} message IDs to manifest: '{self.manifest_path}'",
                    logging.DEBUG,
                )

    def increxportstat(self, stat: ExportStat, incr: int = 1) -> None:
        """Increment specified export stat.

        Args:
            stat (ExportStat): Stat to increment.
            incr (int, optional): Increment amount. Defaults to 1.
        """
        self.export_stats[stat] = self.export_stats.get(stat, 0) + incr

    def write_output_file(
        self,
        file_path: Path,
        file_text: str | bytes,
        file_date: datetime | None,
    ) -> None:
        """Write output file and set its modification date.

        Args:
            file_path (Path): Output file path.
            file_text (str | bytes): File content. If str, it will be written as text; if bytes, as binary.
            file_date (datetime | None): Modification date. If None, the file's modification date will not be set.
        """
        if not self.export_opts.dry_run:
            if isinstance(file_text, str):
                file_path.write_text(file_text, encoding="utf-8")
            else:
                file_path.write_bytes(file_text)

            if file_date:
                os.utime(file_path, (file_date.timestamp(), file_date.timestamp()))

            self.emit_log(f"Wrote file to '{file_path.name}'", logging.DEBUG)
        else:
            self.emit_log(f"Wrote file to '{file_path.name}' (dry run)", logging.DEBUG)

    def prepare_export(self) -> None:
        """Prepare to export messages.

        The default implementation creates the files directory.
        """
        self.emit_log(f"Saving files to '{self.export_opts.files_dir}'")
        self._mkdir(self.export_opts.files_dir, "files", parents=self.export_opts.files_mkdir)

        # Note: Attachments directory will be created the first time an attachment is saved.
        if self.export_opts.with_attachments:
            self.emit_log(f"Saving attachments to '{self.export_opts.attachments_folder}'")

    def progress_update(  # noqa: PLR0913
        self,
        *,
        total: float | None = None,
        completed: float | None = None,
        advance: float | None = None,
        visible: bool | None = None,
        refresh: bool = False,
        status: str | None = None,
    ) -> None:
        """Update folder export progress.

        Args:
            total (float, optional): Updates task.total if not None.
            completed (float, optional): Updates task.completed if not None.
            advance (float, optional): Add a value to task.completed if not None.
            visible (bool, optional): Set visible flag if not None.
            refresh (bool): Force a refresh of progress information. Default is False.
            status (str, optional): Updates task status if not None.
        """
        kwargs = {}
        if status:
            kwargs["status"] = status

        self.progress.update(
            self.progress_task_id,
            total=total,
            completed=completed,
            advance=advance,
            visible=visible,
            refresh=refresh,
            **kwargs,
        )

    def export_loop(
        self,
        first_batch_data: GraphApiReturnT | None,
        export_limit: int,
        params: dict[str, str | int] | None,
    ) -> None:
        """Internal loop for exporting messages.

        All messages are processed in this loop (which includes `finalize_export` at the end).

        Args:
            first_batch_data (GraphApiReturnT | None): First batch of messages.
            export_limit (int): Maximum number of messages to export.
            params (dict[str, str | int] | None): Parameters for the listing messages.

        Raises:
            ExportError if an export error occurred.
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        assert params is not None  # noqa: S101

        next_url = "..."  # ignored when `first_batch_data` is set
        n_batch = 0
        export_total = 0
        dupes_total = 0
        batch_size = int(params.get("$top", 0)) or export_limit

        self.emit_log(f"Fetching {batch_size} messages at a time", logging.DEBUG)

        while next_url is not None and export_total < export_limit:
            self.raise_if_cancelled()

            n_batch += 1

            batch_total, batch_dups, batch_written, next_url = self._export_folder_batch(
                first_batch_data,
                next_url,
                params,
                min(batch_size, export_limit - export_total),
            )

            self.emit_log(
                [
                    f"Batch #{n_batch}:",
                    f"{batch_total:,} messages,",
                    f"{batch_dups:,} duplicates,",
                    f"{batch_written:,} written",
                ],
                logging.DEBUG,
            )

            # Doing this on every batch is inefficient for large folders but
            # ensures the manifest is mostly up-to-date if a truly fatal error occurs.
            if batch_written:
                self._write_manifest()

            export_total += batch_total
            dupes_total += batch_dups
            params = None  # included in 'nextLink' automatically
            first_batch_data = None

        # Finish up any messages that may have been queued.
        n_written = self.finalize_export()

        self.emit_log(
            [
                "Export summary:",
                f"{n_batch} batches,",
                f"{export_total:,} messages,",
                f"{dupes_total:,} duplicates",
            ],
            logging.DEBUG,
        )

        # Update the progress to indicate that the export is complete.
        # We're not using 'advance=n_written' because I like the progress bar to be 100%
        # even when a lot of the messages were duplicates (previously exported.)
        # Otherwise, when 222 of 322 messages are exported with 100 duplicates,
        # the progress bar stops at 222/322 :/
        self.progress_update(completed=self.progress_task.total, status="done")

        # The manifest isn't final until after `finalize_export`!
        if n_written:
            self._write_manifest(emitlog=True)

    @property
    def exported_message_count(self) -> int:
        """Returns count of non-duplicate messages exported.

        Returns:
            int
        """
        return self.export_stats.get(ExportStat.STAT_MESSAGE_EXPORTED, 0)

    ######################################################################
    # Abstract methods

    @abstractmethod
    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
    ) -> int:
        """Write a message to a file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

    ######################################################################
    # Public functions

    def export(  # noqa: C901, PLR0912, PLR0915
        self,
        thread_id: int,
        session: Session,
    ) -> None:
        """Export messages from a specific folder to the output directory.

        Results of the export operation can be found in the `export_results` and `export_stats` properties.

        Args:
            thread_id (int): Thread ID of the worker.
            session (Session): Session to use to pool requests.

        Raises:
            ExportCancelledError if export was cancelled.
            ExportError if a non-cancellable export error occurred.
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        self.thread_id = thread_id
        self.session = session

        self.update_export_opts()

        # Wait for the first thread/worker to get its first batch of messages.
        # This allows the first worker to encounter any issues with
        # the Graph API before other workers proceed and encounter the same.
        self._wait_for_release_workers()

        self.emit_log("Working on this folder...", logging.DEBUG)

        # Reset task start time to get true elapsed time for this folder.
        self.progress.reset(self.progress_task_id, start=True)

        folder_message_count = self.folder["totalItemCount"]

        if not folder_message_count:
            self.progress_update(total=0, completed=0, status="empty folder", visible=False)
            self.export_result.add_message("Folder is empty")
            return

        self.progress_update(status="evaluating", refresh=True)

        # Ignore manifest when doing filter test
        manifest_data = self._read_manifest() if not self.export_opts.filter_test else None

        export_limit = self.export_opts.files_limit  # 0=no limit
        manifest_id_count = len(self.manifest_ids)

        # https://learn.microsoft.com/en-us/graph/query-parameters
        params: dict[str, str | int] | None = {
            "$count": "true",
            "$top": min(export_limit, self.export_opts.batch_size) or self.export_opts.batch_size,
        }

        if self.export_opts.filter_query:
            params["$filter"] = self.export_opts.filter_query
            manifest_id_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

        elif self.export_opts.search_query:
            params["$search"] = self.export_opts.search_query
            manifest_id_count = 0  # Manifest can be used to find dups but we but can't "skip" ahead

        else:
            # The Graph API imposes rules when using $orderby and $filter together.
            # $search and $orderby cannot be used together.
            # Orderby multiple fields so order is more deterministic.
            params["$orderby"] = "lastModifiedDateTime,internetMessageId,subject"

        if (
            manifest_data
            and manifest_id_count
            and self.export_opts.filter_query == manifest_data["filterQuery"]
            and self.export_opts.search_query == manifest_data["searchQuery"]
        ):
            # We can skip previously exported messages when filtering (or not)
            # is kept consistent between runs! Otherwise, messages that should
            # be included are certainly being skipped.
            params["$skip"] = manifest_id_count

        self.update_messages_params(params)

        if self.export_opts.filter_test:
            # This *overrides* everyone!
            params["$top"] = 1  # We're only interested in the message $count
            params["$select"] = "id"  # Might as well make the response shorter

        # Sort to ensure consistent ordering for caching
        params = dict(sorted(params.items()))

        # https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
        # https://learn.microsoft.com/en-us/graph/api/message-get
        next_url = f"{GRAPH_API_URL}/me/mailFolders/{self.folder_id}/messages"

        # Request the first batch now so we can get a count of the results.
        first_batch_data = self._list_messages(next_url, params=params)

        first_batch_messages = first_batch_data.get("value", [])

        # I'm guessing '@odata.count' is not provided when search results are less than '$top'.
        messages_count: int = first_batch_data.get("@odata.count") or len(first_batch_messages)

        if self.export_opts.filter_test:
            # Report the filtered message count next to task progress.
            self.progress_update(total=1, completed=1, status=f"{messages_count} messages")
            return

        if not messages_count:
            if self.export_opts.filter_query:
                msg, level = "No messages match 'filter' criteria", logging.INFO
            elif self.export_opts.search_query:
                msg, level = "No messages match 'search' criteria", logging.INFO
            else:
                msg, level = f"Contains {folder_message_count} messages but Graph API result was empty", logging.ERROR

            self.export_result.add_message(msg, level)

            return

        # When filtering or searching, consider all messages remaining.
        messages_remaining = (
            messages_count
            if self.export_opts.filter_query or self.export_opts.search_query
            else max(0, messages_count - manifest_id_count)
        )

        if not messages_remaining:
            self.emit_log("No remaining messages", logging.DEBUG)  # log stream
            self.export_result.add_message("All messages have been exported.")  # reported in results
            self.progress_update(total=0, completed=0, status="done")
            return

        export_limit = min(export_limit, messages_remaining) if export_limit else messages_remaining

        self.progress_update(total=export_limit, completed=0)

        self.emit_log(
            [
                "Exporting",
                f"all {folder_message_count:,}" if export_limit == folder_message_count else None,
                f"{export_limit:,} of {messages_count:,}" if export_limit < folder_message_count else None,
                "(filtered)" if self.export_opts.filter_query or self.export_opts.search_query else None,
                "messages",
                "(skipping dups)" if self.manifest_path else None,
                "(dry run)" if self.export_opts.dry_run else None,
            ],
        )

        self.prepare_export()

        try:
            self.export_loop(
                first_batch_data,
                export_limit,
                params,
            )
        except requests.RequestException, ExportError, ValueError, OSError:
            # Try to keep the manifest up-to-date even if an error occurs.
            self._write_manifest()
            raise

        if not self.exported_message_count:
            self.emit_log("No messages were exported.")

    ######################################################################
    # Common overridden methods

    def update_export_opts(self) -> None:
        """Set and optionally update export options."""
        files_dir = self.export_opts.files_dir
        attachments_folder = self.export_opts.attachments_folder
        attachments_dir_path = Path(attachments_folder)  # assume it's an absolute path to start
        attachments_mkdir = self.export_opts.attachments_mkdir

        # Determine the folder path to use when creating subdirectories for files and attachments.
        # For example, if path is "A/B/C" and mkdir_maxdepth is 2, folder_path_mkdir will be "B/C".
        folder_path_mkdir = (
            self.folder_path
            if not self.export_opts.mkdir_maxdepth
            else Path(*self.folder_path.parts[-self.export_opts.mkdir_maxdepth :])
        )

        if self.export_opts.files_mkdir:
            files_dir = files_dir / folder_path_mkdir

        if isinstance(attachments_folder, Path):
            assert attachments_dir_path.is_dir()  # noqa: S101 ; caller should have verified!
            assert attachments_dir_path.is_absolute()  # noqa: S101 ; caller should have verified!
        elif attachments_folder.upper() == ATTACHMENTS_FOLDER_ROOT:
            attachments_dir_path = self.export_opts.root_out_dir
            attachments_mkdir = False
        elif attachments_folder.upper().startswith(ATTACHMENTS_FOLDER_FILE):  # allow plural "FILES"
            attachments_dir_path = files_dir
            attachments_mkdir = False
        elif not attachments_dir_path.is_absolute():
            attachments_dir_path = files_dir / safe_file_name(attachments_folder, ATTACHMENTS_DEFAULT_FOLDER)
            attachments_mkdir = False

        if attachments_mkdir:
            attachments_dir_path = attachments_dir_path / folder_path_mkdir

        self.export_opts.files_dir = files_dir
        self.export_opts.attachments_mkdir = attachments_mkdir
        self.export_opts.attachments_folder = attachments_dir_path

    def update_messages_params(self: Self, params: dict[str, str | int]) -> None:  # noqa: B027
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """

    @abstractmethod
    def get_file_path(
        self: Self,
        message: GraphMessageT,
    ) -> Path:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.

        Returns:
            Path
        """
        title = message["fileTitle"]
        file_date = message["messageDateTime"]

        return unique_file_path(self.export_opts.files_dir / Path(title + ".tbd"), file_date)

    def finalize_export(self: Self) -> int:
        """Called after the last batch of messages for a folder is complete.."""
        return 0
