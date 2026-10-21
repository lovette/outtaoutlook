"""Shared helpers and constants for the export command."""

import os
import threading
from collections import defaultdict
from dataclasses import dataclass, field
from enum import StrEnum
from typing import TYPE_CHECKING

from outtaoutlook.shared import logger

if TYPE_CHECKING:
    from datetime import datetime
    from pathlib import Path

    from requests import sessions
    from rich.progress import Progress, TaskID

    from outtaoutlook.shared import FileMetaPropT, GraphApiConfig, GraphFolderT

######################################################################
# Constants

MAX_EXPORT_WORKERS = 32  # maximum ThreadPoolExecutor will use

ATTACHMENTS_FOLDER_FILE = "same"
ATTACHMENTS_FOLDER_ROOT = "root"

ATTACHMENTS_DEFAULT_FOLDER = "assets"


######################################################################
# Types and classes


class OutputFormatChoice(StrEnum):
    """Choices for output file format."""

    eml = "eml"
    html = "html"
    json = "json"
    md = "md"


class ExportStat(StrEnum):
    """Export statistic messages."""

    STAT_FILE_CREATED = "Files created"
    STAT_ATTACHMENT_SAVED = "Attachments saved"
    STAT_SKIPPED_DUPLICATE = "Skipped; Message already exported"
    STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED = "Skipped; Markdown conversion failed"
    STAT_SKIPPED_NO_MESSAGE_BODY = "Skipped; No message body"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS = "Skipped; No file content, just attachments (not being saved)"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS = "Skipped; No file content"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN = "Skipped; No file content, unknown reason"
    STAT_SKIPPED_EVENT_MESSAGE = "Skipped; Calendar event"


ExportStatsT = dict[ExportStat, int]


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
    mkdir_maxdepth: int | None  # Maximum depth (from the leaf) of folder tree to create when creating mail folder hierarchy.
    md_body_tags: list[str] | None  # --body-tag tags from command line and config file.
    md_metaprop: FileMetaPropT | None  # --metaprop properties from command line and config file.
    md_tags: list[str] | None  # --tag tags from command line and config file.
    files_dir: Path  # Directory to save files (set internally.)
    files_limit: int  # Maximum number of files to export (per folder).
    files_mkdir: bool  # Create mail folder hierarchy in files directory.
    md_place_tags_body_end: bool  # True to place body tags at the end of the file instead the top.
    md_place_tags_frontmatter: bool  # True to place non-body tags in frontmatter.
    search_query: str | None  # --search criteria from command line and config file.
    root_out_dir: Path  # Root output directory (set by user.)
    with_attachments: bool  # Whether to download and link attachments to files.
    max_workers: int  # Maximum number of workers being used for exporting.


@dataclass
class ExportResult:
    """Results of folder export."""

    aborted: bool = False  # Whether the export was aborted.
    messages: defaultdict[int, list[str]] = field(default_factory=lambda: defaultdict(list))  # {level:[message]}
    stats: ExportStatsT | None = None  # Export statistics keyed by ExportStat.


@dataclass
class ExportFolderContext:
    """Context for exporting a folder."""

    folder: GraphFolderT
    session: sessions.Session
    graph_config: GraphApiConfig
    progress: Progress
    progress_task_id: TaskID
    export_opts: ExportOptions
    export_stats: ExportStatsT


######################################################################
# Global variables


# Thread-local storage for per-thread data.
thread_local_data = threading.local()

# Keep track of whether the first thread spawned has received its first batch of messages.
first_batch_complete_event = threading.Event()


######################################################################
# Public functions


def exportstat_incr(export_stats: ExportStatsT, stat: ExportStat, incr: int = 1) -> None:
    """Increment specified export stat.

    Args:
        export_stats (ExportStatsT): Export stats dictionary.
        stat (ExportStat): Stat to increment.
        incr (int, optional): Increment amount. Defaults to 1.
    """
    export_stats[stat] = export_stats.get(stat, 0) + incr


def write_output_file(
    file_path: Path,
    file_text: str,
    file_date: datetime | None,
    export_opts: ExportOptions,
) -> None:
    """Write output file and set its modification date.

    Args:
        file_path (Path): Output file path.
        file_text (str): File content.
        file_date (datetime | None): Modification date. If None, the file's modification date will not be set.
        export_opts (ExportOptions): Export options.
    """
    if not export_opts.dry_run:
        file_path.write_text(file_text, encoding="utf-8")

        if file_date:
            os.utime(file_path, (file_date.timestamp(), file_date.timestamp()))

        logger.debug("Wrote file to '%s'", file_path.name)
    else:
        logger.debug("Wrote file to '%s' (dry run)", file_path.name)


def _release_workers() -> None:
    if thread_local_data.thread_id == 1 and not first_batch_complete_event.is_set():
        logger.debug("Releasing workers")
        first_batch_complete_event.set()
