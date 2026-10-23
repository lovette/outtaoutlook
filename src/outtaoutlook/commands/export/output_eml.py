"""Export messages as raw .eml files."""

import logging
from typing import TYPE_CHECKING, Self

from outtaoutlook.graphapi import GraphMessageT, log_message_notice

from .output_core import ExportStat
from .output_queued import OutputWorkerQueued

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


######################################################################
# Public module class


class OutputWorkerEml(OutputWorkerQueued):
    """Implements the `OutputWorkerQueued` for .eml output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

    ######################################################################
    # Method overrides

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

    def write_queued_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,  # noqa: ARG002
        content: bytes | None,
    ) -> int:
        """Export a message as a raw .eml file.

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

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            self.increxportstat(skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return 1

        file_date = message["messageDateTime"]

        if content is None:
            return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)

        self.write_output_file(file_path, content, file_date)

        self.increxportstat(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1
