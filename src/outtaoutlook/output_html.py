"""Export messages as html files."""

import logging
from pathlib import Path

from outtaoutlook.shared import (
    ExportOptions,
    ExportStat,
    ExportStatsT,
    GraphFolderT,
    GraphMessageT,
    exportstat_incr,
    log_message_notice,
    unique_file_path,
    write_output_file,
)

######################################################################
# Public module class


class ExportFormatHtml:
    """Implements the `OutputFormatModule` protocol for HTML output."""

    @staticmethod
    def get_folder_export_opts(
        folder: GraphFolderT,  # noqa: ARG004
        export_opts: ExportOptions,
    ) -> ExportOptions:
        """Generate export options for a specific folder.

        Args:
            folder (GraphFolderT): The folder being exported.
            export_opts (ExportOptions): Base export options.

        Returns:
            ExportOptions: A copy of export options to use for messages in this folder.
        """
        return export_opts

    @staticmethod
    def update_list_messages_params(params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """

    @staticmethod
    def get_file_path(
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

    @staticmethod
    def export_message(
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,  # noqa: ARG004
        export_opts: ExportOptions,
        export_stats: ExportStatsT,
    ) -> bool:
        """Export a message as a HTML file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.
            export_opts (ExportOptions): Export options.
            export_stats (ExportStatsT): Export statistics dictionary to update.

        Returns:
            bool: True if the message was successfully exported, False otherwise.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, htmlDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _skip_message(skip_reason: ExportStat, level: int) -> bool:
            exportstat_incr(export_stats, skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return False

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

        return True
