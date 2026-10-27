"""Shared helpers and constants for the export command."""

from dataclasses import dataclass
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from outtaoutlook.shared import FileMetaPropT

######################################################################
# Constants


MAX_EXPORT_WORKERS = 32  # maximum ThreadPoolExecutor will use

ATTACHMENTS_FOLDER_FILE = "same"
ATTACHMENTS_FOLDER_ROOT = "root"

ATTACHMENTS_DEFAULT_FOLDER = "assets"

# Message fields we reference for exporting.
# We need subject for log messages and dates for file name and timestamp.
DEFAULT_SELECT_MESSAGE_FIELDS = (
    "id",
    "subject",
    "isDraft",
    "hasAttachments",
    "createdDateTime",
    "lastModifiedDateTime",
    "sentDateTime",
    "receivedDateTime",
)


######################################################################
# Types and classes


class OutputFormatChoice(StrEnum):
    """Choices for output file format."""

    eml = "eml"
    html = "html"
    json = "json"
    mbox = "mbox"
    maildir = "maildir"
    md = "md"


@dataclass
class ExportOptions:
    """Export command options."""

    out_format: OutputFormatChoice  # Output file format.
    attachments_folder: str | Path  # `ATTACHMENTS_FOLDER_*`, folder name or absolute path.
    attachments_mkdir: bool  # Create mail folder hierarchy in attachments directory.
    batch_size: int  # Number of messages to process in a batch.
    dry_run: bool  # Whether to perform a dry run without making any changes.
    filter_query: str | None  # --filter criteria from command line and config file.
    filter_test: bool  # Whether to perform a filter test without exporting.
    manifest_dir: Path | None  # Directory to write the manifest file.
    hier_shallow: bool  # Use only the leaf of folder tree when creating mail folder hierarchy.
    md_body_tags: list[str] | None  # --body-tag tags from command line and config file.
    meta_props: FileMetaPropT | None  # --metaprop properties from command line and config file.
    md_tags: list[str] | None  # --tag tags from command line and config file.
    files_dir: Path  # Directory to save files (set internally.)
    export_limit: int  # Maximum number of files to export (per folder).
    hier_flat: bool  # Write all files to root directory, do not preserve mail folder hierarchy.
    md_place_tags_body_end: bool  # True to place body tags at the end of the file instead the top.
    md_place_tags_frontmatter: bool  # True to place non-body tags in frontmatter.
    search_query: str | None  # --search criteria from command line and config file.
    root_out_dir: Path  # Root output directory (set by user.)
    with_attachments: bool  # Whether to download and link attachments to files.
    max_workers: int  # Maximum number of workers being used for exporting.
    assume_yes: bool  # Automatically confirm all prompts.
    prime_cache: bool  # Whether to prime the cache without exporting files.
    verbose: int  # Verbosity level (VERBOSE_LEVEL_*)


######################################################################
# Public functions


def json_serializer(obj: object) -> str:
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
    if isinstance(obj, Path):
        return str(obj)

    msg = f"Object of type {type(obj)} is not JSON serializable"
    raise TypeError(msg)
