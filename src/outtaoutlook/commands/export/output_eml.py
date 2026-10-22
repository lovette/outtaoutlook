"""Export messages as raw .eml files."""

import base64
import binascii
import logging
from dataclasses import dataclass
from typing import TYPE_CHECKING, Self

from outtaoutlook.graphapi import (
    GRAPH_API_URL,
    GraphMessageT,
    graph_get_resource_batch,
    log_message_notice,
)
from outtaoutlook.shared import (
    logger,
)

from .output_core import ExportStat, OutputWorkerCore

if TYPE_CHECKING:
    from pathlib import Path

######################################################################
# Constants


# Message fields we reference for exporting.
# We need subject for log messages and dates for file name and timestamp.
SELECT_MESSAGE_FIELDS = ",".join(
    sorted(
        [
            "id",
            "subject",
            "isDraft",
            "hasAttachments",
            "createdDateTime",
            "lastModifiedDateTime",
            "sentDateTime",
            "receivedDateTime",
        ],
    ),
)

REQUESTS_PER_BATCH = 20  # Graph API limit


######################################################################
# Types and classes


@dataclass
class MessageQueueItemT:
    """Items in the message export queue."""

    message: GraphMessageT
    file_path: Path
    attachment_paths: list[Path] | None


######################################################################
# Public module class


class OutputWorkerEml(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for EML output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

        # Private instance properties
        self._queued_messages: dict[str, MessageQueueItemT] = {}  # {message-id: MessageQueueItemT}

        # If there are multiple threads making requests, we need Graph API to execute
        # batched individual requests in sequence so we don't hit concurrency limits.
        self._sequence_batch = self.export_opts.max_workers > 1

        self._batch_request_id_start = 1

    def update_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        # The raw value is fetched using Graph API with the ID.
        # Subject is required for log messages.
        params["$select"] = SELECT_MESSAGE_FIELDS

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
        return super().get_file_path(message).with_suffix(".eml")

    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
    ) -> int:
        """Export a message as a raw .eml file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        message_id: str | None = message.get("id")
        assert message_id is not None  # noqa: S101

        # Instead of making one '/message/{id}/$value' request for each message,
        # we batch them together and submit a single request for multiple messages.

        self._queued_messages[message_id] = MessageQueueItemT(
            message=message,
            file_path=file_path,
            attachment_paths=attachment_paths,
        )

        if len(self._queued_messages) >= REQUESTS_PER_BATCH:
            return self._flush_queue()
        return 0

    def finalize_export(self: Self) -> int:
        """Called after the last batch of messages for a folder is complete.."""
        if len(self._queued_messages):
            return self._flush_queue()
        return 0

    ######################################################################
    # Internal methods

    def _flush_queue(self) -> int:

        def _batch_request_id(idx: int) -> str:
            """Outlook message IDs are too long to use as batch request IDs so we use a shorter one."""
            return f"{self.folder_name}-{self._batch_request_id_start + idx}"

        assert len(self._queued_messages) > 0  # noqa: S101
        assert self.session is not None  # noqa: S101

        logger.debug("Processing %d queued messages", len(self._queued_messages))

        url_by_message_id = {
            message_id: f"{GRAPH_API_URL}/me/messages/{message_id}/$value" for message_id in self._queued_messages
        }

        # Create a deterministic cache key.
        cache_key = {
            "url": "/me/messages/.../$value",
            "ids": sorted(self._queued_messages),
        }

        self.progress_update(status="fetching messages", refresh=True)

        # Submit a single Graph API request that returns $value for all the messages in our batch.
        responses = graph_get_resource_batch(
            self.session,
            self.graph_config,
            url_by_message_id,
            sequence=self._sequence_batch,
            cache_key=cache_key,
            short_key_factory=_batch_request_id,
        )

        self.progress_update(status="writing files", refresh=True)

        assert len(responses) == len(url_by_message_id)  # noqa: S101
        n_messages = len(responses)

        for request_id, response in responses.items():
            self.raise_if_cancelled()

            item = self._queued_messages.pop(request_id)  # dequeue

            content, content_type = response.get("body"), response.get("headers", {}).get("Content-Type")

            if content_type != "text/plain":
                log_message_notice(item.message, f"$batch/messages/id/$value: Unexpected content type: '{content_type}'")
                content = None

            if content:
                try:
                    content = base64.b64decode(content).decode("utf-8")
                except binascii.Error as err:
                    log_message_notice(
                        item.message,
                        f"$batch/messages/id/$value: Failed to decode message content: {err}",
                        logging.ERROR,
                    )

            n_exported = self._real_export_message(item, content)
            if n_exported:
                self.manifest_ids.add(item.message["id"])

        self._batch_request_id_start += n_messages

        return n_messages

    def _real_export_message(self: Self, item: MessageQueueItemT, content: str | None) -> int:
        """Export a message as a raw .eml file.

        Args:
            item (MessageQueueItemT): The message queue item containing all necessary information.
            content (str | None): The raw content of the message, or None if not available.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            self.increxportstat(skip_reason)
            log_message_notice(item.message, skip_reason.value, level)
            return 1

        file_date = item.message["messageDateTime"]
        has_attachments = item.message.get("hasAttachments")

        if not content:
            if content is None:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)
            if has_attachments and self.export_opts.with_attachments:
                content = ""  # content will be links to attachments
            if not has_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO)
            if not self.export_opts.with_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO)

        self.write_output_file(item.file_path, content, file_date)

        self.increxportstat(ExportStat.STAT_FILE_CREATED)

        return 1
