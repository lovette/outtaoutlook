"""Export messages as JSON files."""

import json
from typing import TYPE_CHECKING, Self

from .output_core import ExportStat, OutputWorkerCore
from .shared import json_serializer

if TYPE_CHECKING:
    from pathlib import Path

    from outtaoutlook.graphapi import GraphMessageT


######################################################################
# Public module class


class OutputWorkerJson(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for JSON output."""

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
        params.pop("$select", None)  # We want all the fields

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
        return super().get_file_path(message, ".json")

    def export_write_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path | None,
        attachment_paths: list[Path] | None,  # noqa: ARG002
    ) -> int:
        """Export a message as a JSON file.

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
        assert file_path is not None  # noqa: S101

        file_date = message["messageDateTime"]

        message_copy = message.copy()

        # These don't need to be exported
        message_copy.pop("@odata.etag", None)
        message_copy.pop("fileTitle", None)

        file_text = json.dumps(message_copy, sort_keys=True, default=json_serializer)

        if self.export_opts.prime_cache:
            return 1  # only prime the cache, do not write files

        self.write_output_file(file_path, file_text, file_date)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1
