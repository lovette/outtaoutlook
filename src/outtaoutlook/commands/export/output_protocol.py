"""Protocol for output format modules."""

from typing import TYPE_CHECKING, Protocol, Self

if TYPE_CHECKING:
    from pathlib import Path

    from outtaoutlook.shared import GraphMessageT

    from .shared import ExportFolderContext, ExportOptions, ExportStatsT


######################################################################
# Classes and types


class OutputFormatModule(Protocol):
    """Protocol for output format modules."""

    # Instance variables all implementations must have.
    export_context: ExportFolderContext

    def __init__(self, export_context: ExportFolderContext) -> None:
        """Object constructor.

        Args:
            export_context (ExportFolderContext): Context for the folder being exported.
        """
        ...

    def set_export_opts(self: Self, export_opts: ExportOptions) -> ExportOptions:
        """Set and optionally update export options for messages in the current folder.

        Args:
            export_opts (ExportOptions): Base export options.

        Returns:
            ExportOptions: Export options to use for messages in this folder.
        """
        ...

    def update_list_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        ...

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
        ...

    def export_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
        export_opts: ExportOptions,
        export_stats: ExportStatsT,
    ) -> int:
        """Export a message as an output file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.
            export_opts (ExportOptions): Export options.
            export_stats (ExportStatsT): Export statistics dictionary to update.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        ...

    def last_batch_complete(self: Self) -> int:
        """Called after the last batch of messages for a folder is complete.."""
        ...
