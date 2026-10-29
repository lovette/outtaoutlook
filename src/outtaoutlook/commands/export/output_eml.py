"""Export messages as raw .eml files."""

from typing import TYPE_CHECKING, Self

from .output_core import ExportStat
from .output_queued import OutputWorkerQueued
from .shared import DEFAULT_SELECT_MESSAGE_FIELDS

if TYPE_CHECKING:
    import mailbox
    from pathlib import Path

    from outtaoutlook.graphapi import GraphMessageT


######################################################################
# Constants


# Message fields we reference for exporting.
SELECT_MESSAGE_FIELDS = ",".join(sorted(DEFAULT_SELECT_MESSAGE_FIELDS))


######################################################################
# Public module class


class OutputWorkerEml(OutputWorkerQueued):
    """Implements the `OutputWorkerQueued` for .eml output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

    ######################################################################
    # Method overrides

    def update_export_opts(self) -> None:
        """Set and optionally update export options for messages in the current folder."""
        super().update_export_opts()

        self.convert_meta_props_to_xheaders()

        if self.export_opts.with_attachments is None:
            # Only export attachments when explicitly asked
            self.export_opts.with_attachments = False

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
        suffix: str | None = None,  # noqa: ARG002
    ) -> Path | None:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.
            suffix (str | None): Optional file suffix (e.g., ".md", ".eml").

        Returns:
            Path | None
        """
        return super().get_file_path(message, ".eml")

    def write_queued_message(
        self: Self,
        message: GraphMessageT,
        mb_message: mailbox.Message,
        file_path: Path | None,
    ) -> int:
        """Export a message as a raw .eml file.

        Args:
            message (GraphMessageT): The message.
            mb_message (mailbox.Message): The mailbox message object.
            file_path (Path | None): The path to the output file.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """
        assert file_path is not None  # noqa: S101

        file_date = message["messageDateTime"]

        self.write_output_file(file_path, mb_message.as_bytes(), file_date)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1
