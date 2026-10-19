"""Protocol for output format modules."""

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from pathlib import Path

    from outtaoutlook.shared import ExportOptions, ExportStatsT, GraphFolderT, GraphMessageT


class OutputFormatModule(Protocol):
    """Protocol for output format modules."""

    @staticmethod
    def get_folder_export_opts(
        folder: GraphFolderT,
        export_opts: ExportOptions,
    ) -> ExportOptions:
        """Generate export options for a specific folder.

        Args:
            folder (GraphFolderT): The folder being exported.
            export_opts (ExportOptions): Base export options.

        Returns:
            ExportOptions: A copy of export options to use for messages in this folder.
        """
        ...

    @staticmethod
    def update_list_messages_params(params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        ...

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
        ...

    @staticmethod
    def export_message(
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
        export_opts: ExportOptions,
        export_stats: ExportStatsT,
    ) -> bool:
        """Export a message as an output file.

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
        ...
