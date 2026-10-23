"""Export queued messages as files."""

import base64
import binascii
import hashlib
import logging
from abc import abstractmethod
from dataclasses import dataclass
from typing import TYPE_CHECKING, Self

from outtaoutlook.graphapi import (
    GRAPH_API_URL,
    GraphMessageT,
    graph_get_resource_batch,
    log_message_notice,
)

from .output_core import OutputWorkerCore

if TYPE_CHECKING:
    from pathlib import Path

######################################################################
# Constants


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


class OutputWorkerQueued(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for queued output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

        # Private instance properties
        self._queued_messages: dict[str, MessageQueueItemT] = {}  # {message-id: MessageQueueItemT}

        # If there are multiple threads making requests, we need Graph API to execute
        # batched individual requests in sequence so we don't hit concurrency limits.
        self._sequence_batch = self.export_opts.max_workers > 1

    ######################################################################
    # Method overrides

    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
    ) -> int:
        """Export a message as a file.

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
            return self.write_queued_messages()

        return 0

    def finalize_export(self: Self) -> int:
        """Called after the last batch of messages for a folder is complete.."""
        if len(self._queued_messages):
            return self.write_queued_messages()

        return 0

    ######################################################################
    # Queue methods

    def write_queued_messages(self) -> int:
        """Write queued messages to file.

        Returns:
            int: The number of files written.
        """

        def _batch_request_id(message_id: str, idx: int) -> str:  # noqa: ARG001
            # Outlook message immutable IDs are 68 characters in length.
            # The Graph API truncates them at some point so we get duplicate key errors.
            # The short key should be deterministic so cached responses can be correctly matched.
            sha256_hex = hashlib.sha256(message_id.encode("utf-8")).hexdigest()
            return sha256_hex[:32]

        assert len(self._queued_messages) > 0  # noqa: S101
        assert self.session is not None  # noqa: S101

        self.emit_log(f"Processing {len(self._queued_messages)} queued messages", logging.DEBUG)

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

        self.progress_update(status=self.write_messages_status, refresh=True)

        assert len(responses) == len(url_by_message_id)  # noqa: S101
        n_messages = len(responses)

        for request_id in list(self._queued_messages):
            self.raise_if_cancelled()

            item = self._queued_messages.pop(request_id)  # dequeue
            response = responses[request_id]

            content, content_type = response.get("body"), response.get("headers", {}).get("Content-Type")

            if content_type != "text/plain":
                log_message_notice(item.message, f"$batch/messages/id/$value: Unexpected content type: '{content_type}'")
                content = None

            if content:
                try:
                    content = base64.b64decode(content)
                except binascii.Error as err:
                    log_message_notice(
                        item.message,
                        f"$batch/messages/id/$value: Failed to b64decode message content: {err}",
                        logging.ERROR,
                    )
                    content = None

            n_exported = self.write_queued_message(
                item.message,
                item.file_path,
                item.attachment_paths,
                content,
            )

            if n_exported:
                self.manifest_ids.add(item.message["id"])

        return n_messages

    ######################################################################
    # Abstract methods

    @abstractmethod
    def write_queued_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
        content: bytes | None,
    ) -> int:
        """Export a message as a file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): The path to the output file.
            attachment_paths (list[Path] | None): List of paths to attachments, or None if there are no attachments.
            content (bytes | None): The raw content of the message, or None if not available.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
