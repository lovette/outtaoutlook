"""Export messages as html files."""

import logging
from pathlib import Path
from typing import Self

from outtaoutlook.graphapi import GraphMessageT, log_message_notice
from outtaoutlook.shared import (
    unique_file_path,
)

from .shared import (
    ExportFolderContext,
    ExportOptions,
    ExportStat,
    ExportStatsT,
    exportstat_incr,
    write_output_file,
)

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


class ExportFormatHtml:
    """Implements the `OutputFormatModule` protocol for HTML output."""

    def __init__(self, export_context: ExportFolderContext) -> None:
        """Object constructor.

        Args:
            export_context (ExportFolderContext): Context for the folder being exported.
        """
        self.export_context = export_context

    def set_export_opts(self: Self, export_opts: ExportOptions) -> ExportOptions:
        """Set and optionally update export options for messages in the current folder.

        Args:
            export_opts (ExportOptions): Base export options.

        Returns:
            ExportOptions: Export options to use for messages in this folder.
        """
        return export_opts

    def update_list_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        params["$select"] = SELECT_MESSAGE_FIELDS

    def get_file_path(
        self: Self,
        message: GraphMessageT,
        export_opts: ExportOptions,
    ) -> Path:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.
            export_opts (ExportOptions): Export options.

        Returns:
            Path
        """
        title = message["fileTitle"]
        file_date = message["messageDateTime"]

        return unique_file_path(export_opts.files_dir / Path(title + ".html"), file_date)

    def export_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,  # noqa: ARG002
        export_opts: ExportOptions,
        export_stats: ExportStatsT,
    ) -> int:
        """Export a message as a HTML file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.
            export_opts (ExportOptions): Export options.
            export_stats (ExportStatsT): Export statistics dictionary to update.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, htmlDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            exportstat_incr(export_stats, skip_reason)
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
            if has_attachments and export_opts.with_attachments:
                body_content = ""  # content will be links to attachments
            if not has_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO)
            if not export_opts.with_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO)

        write_output_file(file_path, body_content, file_date, export_opts)

        exportstat_incr(export_stats, ExportStat.STAT_FILE_CREATED)

        return 1

    def last_batch_complete(self: Self) -> int:
        """Called after the last batch of messages for a folder is complete.."""
        return 0
