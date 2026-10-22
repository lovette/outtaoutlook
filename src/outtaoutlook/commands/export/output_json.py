"""Export messages as JSON files."""

import json
from datetime import datetime
from typing import TYPE_CHECKING, Self

from .output_core import OutputWorkerCore
from .shared import (
    ExportStat,
    exportstat_incr,
    write_output_file,
)

if TYPE_CHECKING:
    from pathlib import Path

    from outtaoutlook.graphapi import GraphMessageT

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


class OutputWorkerJson(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for JSON output."""

    def update_list_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        del params["$select"]  # We want all the fields

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
        return super().get_file_path(message).with_suffix(".json")

    def export_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,
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
        file_date = message["messageDateTime"]

        message_copy = message.copy()

        # These don't need to be exported
        message_copy.pop("@odata.etag", None)
        message_copy.pop("fileTitle", None)

        file_text = json.dumps(message_copy, sort_keys=True, default=_json_serial)

        write_output_file(file_path, file_text, file_date, self.export_opts)

        exportstat_incr(self.export_stats, ExportStat.STAT_FILE_CREATED)

        return 1
