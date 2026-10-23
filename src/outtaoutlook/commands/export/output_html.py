"""Export messages as html files."""

import logging
from typing import TYPE_CHECKING, Self

from outtaoutlook.graphapi import GraphMessageT, log_message_notice

from .output_core import ExportStat, OutputWorkerCore

if TYPE_CHECKING:
    from pathlib import Path

######################################################################
# Constants

# Message fields we require for exporting.
SELECT_MESSAGE_FIELDS = ",".join(
    sorted(
        [
            "id",
            "subject",
            "body",
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


class OutputWorkerHtml(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for HTML output."""

    def update_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
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
        return super().get_file_path(message).with_suffix(".html")

    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,  # noqa: ARG002
    ) -> int:
        """Export a message as a HTML file.

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

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            self.increxportstat(skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return 1

        file_date = message["messageDateTime"]
        body_content: str | None = message.get("body", {}).get("content")
        content_type: str | None = message.get("body", {}).get("contentType")
        has_attachments: bool | None = message.get("hasAttachments")

        if content_type == "html":
            file_path = file_path.with_suffix(".html")
        elif content_type == "text":
            file_path = file_path.with_suffix(".txt")
        else:
            log_message_notice(message, f"field('contentType'): unexpected type: '{content_type}'")

        if not body_content:
            if body_content is None:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)
            if has_attachments and self.export_opts.with_attachments:
                body_content = ""  # content will be links to attachments
            if not has_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO)
            if not self.export_opts.with_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO)

        self.write_output_file(file_path, body_content, file_date)

        self.increxportstat(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1
