"""Export messages as JSON files."""

import json
from datetime import datetime
from pathlib import Path

from outtaoutlook.shared import (
    ExportOptions,
    ExportStat,
    ExportStatsT,
    GraphFolderT,
    GraphMessageT,
    exportstat_incr,
    unique_file_path,
    write_output_file,
)

######################################################################
# Helpers


def _json_serial(obj: object) -> str:
    """JSON serializer for objects not serializable by `json.dumps`.

    Args:
        obj (Any): _description_

    Raises:
        TypeError: _description_

    Returns:
        str: _description_
    """
    if isinstance(obj, datetime):
        return obj.isoformat()

    msg = f"Object of type {type(obj)} is not JSON serializable"
    raise TypeError(msg)


######################################################################
# Public module class


class ExportFormatJson:
    """Implements the `OutputFormatModule` protocol for JSON output."""

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
        del params["$select"]  # We want all the fields

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

        return unique_file_path(export_opts.files_dir / Path(title + ".json"), file_date)

    @staticmethod
    def export_message(
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,  # noqa: ARG004
        export_opts: ExportOptions,
        export_stats: ExportStatsT,
    ) -> bool:
        """Export a message as a JSON file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.
            export_opts (ExportOptions): Export options.
            export_stats (ExportStatsT): Export statistics dictionary to update.

        Returns:
            bool: True if the message was successfully exported, False otherwise.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        file_date = message["messageDateTime"]

        message_copy = message.copy()

        # These don't need to be exported
        message_copy.pop("@odata.etag", None)
        message_copy.pop("fileTitle", None)

        file_text = json.dumps(message_copy, sort_keys=True, default=_json_serial)

        write_output_file(file_path, file_text, file_date, export_opts)

        exportstat_incr(export_stats, ExportStat.STAT_FILE_CREATED)

        return True
