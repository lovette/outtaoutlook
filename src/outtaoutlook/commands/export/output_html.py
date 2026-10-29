"""Export messages as html files."""

import logging
from typing import TYPE_CHECKING, Self

from .output_core import ExportStat, OutputWorkerCore
from .shared import DEFAULT_SELECT_MESSAGE_FIELDS

if TYPE_CHECKING:
    from pathlib import Path

    from outtaoutlook.graphapi import GraphMessageT

######################################################################
# Constants


# Message fields we require for exporting.
SELECT_MESSAGE_FIELDS = ",".join(
    sorted(
        (
            *DEFAULT_SELECT_MESSAGE_FIELDS,
            "body",
        ),
    ),
)


######################################################################
# Public module class


class OutputWorkerHtml(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for HTML output."""

    def update_export_opts(self) -> None:
        """Set and optionally update export options for messages in the current folder."""
        super().update_export_opts()

        if self.export_opts.with_attachments is None:
            self.export_opts.with_attachments = True

    def update_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        params["$select"] = SELECT_MESSAGE_FIELDS

    def get_file_path(
        self: Self,
        message: GraphMessageT,
        suffix: str | None = None,  # noqa: ARG002
    ) -> Path | None:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.
            suffix (str | None): Optional file suffix (e.g., ".md", ".eml").

        Returns:
            Path | None
        """
        _, content_type = self._get_message_content(message)

        return super().get_file_path(message, ".html" if content_type == "html" else ".txt")

    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path | None,
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
            self.export_stats.incr(skip_reason)
            self.emit_log_w_message(message, skip_reason.value, level)
            return 1

        assert file_path is not None  # noqa: S101

        file_date = message["messageDateTime"]
        body_content, content_type = self._get_message_content(message)
        has_attachments: bool | None = message.get("hasAttachments")

        if content_type not in ("html", "text"):
            self.emit_log_w_message(message, f"field('contentType'): unexpected type: '{content_type}'")

        if not body_content:
            if body_content is None:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)
            if has_attachments and self.export_opts.with_attachments:
                body_content = ""  # content will be links to attachments
            if not has_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO)
            if not self.export_opts.with_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO)

        if self.export_opts.prime_cache:
            return 1  # only prime the cache, do not write files

        self.write_output_file(file_path, body_content, file_date)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1

    ######################################################################
    # Internal methods

    def _get_message_content(self, message: GraphMessageT) -> tuple[str | None, str | None]:
        """Get the body content and content type of a message.

        Args:
            message (GraphMessageT): The message.

        Returns:
            tuple[str, str | None]: (Body content, content type)
        """
        body = message.get("body", {})
        return body.get("content"), body.get("contentType")
