"""outtaoutlook shared variables and functions."""

import hashlib
import logging
import os
import re
import textwrap
import threading
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from rich.console import Console

######################################################################
# Constants


APP_NAME = "outtaoutlook"
TITLE_MAXLEN = 255 - 25  # macOS limit minus a few characters for whatever `unique_file_path` tacks on

# These are banned outright: \/:
# Links do not work if title includes: #^[]|
# macOS is fine with: *"<>?
# Title cannot start with a period.
# \x00-\x1f\x7f = [:cntrl:]
RE_TITLE_IGNORECHARS = re.compile(r"[\[\]\/\\:|^#\x00-\x1f\x7f]")
RE_TITLE_WHITESPACE = re.compile(r"\s+")

VERBOSE_LEVEL_ERROR = 0  # logging.ERROR
VERBOSE_LEVEL_WARNING = 1  # -v, logging.WARNING
VERBOSE_LEVEL_INFO = 2  # -vv, logging.INFO
VERBOSE_LEVEL_DEBUG = 3  # -vvv, logging.DEBUG
VERBOSE_LEVEL_MORE = 4  # -vvvv, logging.DEBUG


######################################################################
# Exceptions


class ExportError(Exception):
    """Base class for all export related errors."""


class ExportCancelledError(ExportError):
    """Raised when an export operation is cancelled."""

    def __init__(self, message: str = "Export operation stopped due to cancellation signal.") -> None:
        """Constructor."""
        super().__init__(message)


######################################################################
# Types and classes


FileMetaPropT = dict[str, str | int | Sequence[str | int]]


@dataclass
class CommonCommandOptions:
    """Common command options."""

    auth_file_path: Path | None = None
    cache_dir: Path | None = None
    graph_cache_enabled: bool = False
    verbose: int = VERBOSE_LEVEL_ERROR
    offline: bool = False


@dataclass
class CommandContext:
    """Additional context for command execution accessed through `ctx.obj`."""

    common_opts: CommonCommandOptions = field(init=False)

    config_file_path: Path | None = None  # set if --config-file is provided and is read
    config_file_opts: dict[str, Any] | None = None  # set if configuration file contains non-`default_map` options

    export_metaprops: FileMetaPropT | None = None  # set if export command has `--metaprop` options

    def __post_init__(self) -> None:  # noqa: D105
        self.common_opts = CommonCommandOptions()


######################################################################
# Globals


console = Console(soft_wrap=True)  # Enable soft wrap so newlines are not inserted to wrap lines wider than the terminal
err_console = Console(stderr=True, soft_wrap=True, style="red", highlight=False)
logger = logging.getLogger(APP_NAME)  # initialized in __main__

# There is a gap of time between a unique path name being generated
# and the file is actually created on disk.
# Therefore we need a way to reserve unique paths that is coordinated between all threads.
# We do this by keeping a set of unique paths and a lock to coordinate access.
# dict[path.parent]=(set(path.name), lock)
_unique_file_path_sets: dict[Path, tuple[set[str], threading.Lock]] = defaultdict(lambda: (set(), threading.Lock()))


######################################################################
# Public functions


def datetime_now() -> datetime:
    """Return current datetime in UTC.

    Returns:
        datetime
    """
    return datetime.now(tz=timezone.utc)  # noqa: UP017


def path_with_vars(p: str | Path, *, absolute: bool = True, resolve: bool = False) -> Path:
    """Expand user home directory reference (~) and shell variables ($ENV) in a path.

    Args:
        p (str | Path): A path.
        absolute (bool, optional): Whether to return an absolute path. Defaults to True.
        resolve (bool, optional): Whether to resolve symlinks and remove '..' segments. Defaults to False.

    Returns:
        Path: The expanded path, optionally absolute and resolved based on the parameters.
    """
    path = Path(os.path.expandvars(str(p))).expanduser()
    if resolve:
        path = path.resolve()
    return path.absolute() if absolute else path


def get_hash_filename(content: str, prefix: str, suffix: str) -> Path:
    """Generate a hash-based filename for the given content.

    Args:
        content (str): The content to hash.
        prefix (str): The prefix to prepend to the filename.
        suffix (str): The suffix to append to the filename (e.g. ".json")

    Returns:
        Path
    """
    # Truncating to 32 characters gives us a shorter filename with negligible collision risk.
    full_hash = hashlib.sha256(content.encode()).hexdigest()
    short_hash = full_hash[:32]
    return Path(f"{prefix}_{short_hash}{suffix}")


def safe_file_name(subject: str, default: str) -> str:
    """Clean up title so it is safe to be used as a file name.

    Args:
        subject (str): Message subject.
        default (str): Default title if subject is empty.

    Returns:
        str
    """
    subject = subject.replace(":", " ")  # Common separator so we'll replace it with a safe character.
    subject = RE_TITLE_IGNORECHARS.sub("", subject)
    subject = RE_TITLE_WHITESPACE.sub(" ", subject)

    if len(subject) > TITLE_MAXLEN:
        # Truncate text while preserving whole words if possible
        subject = textwrap.shorten(subject, width=TITLE_MAXLEN, placeholder="")

    subject = subject.strip(" .")

    return subject or default


def unique_file_path(p: Path, date_hint: datetime | None = None) -> Path:
    """Create a Path with unique file name, optionally using a date hint.

    Args:
        p (Path): Existing path.
        date_hint (datetime | None): Optional datetime to include in the filename for uniqueness.

    Raises:
        ExportError: Raised if too many copies of the file already exist.

    Returns:
        Path: Unique path with a non-conflicting file name.
    """

    def _get_unique_path(p_base: Path) -> Path | None:
        n = 1
        unique_path = p_base

        while n < max_paths:
            exists = (unique_path.name in path_names_set) or unique_path.exists()
            path_names_set.add(unique_path.name)
            if not exists:
                return unique_path
            unique_path = p_base.with_stem(f"{p_base.stem} ({n})")
            n += 1

        return None

    # Keep in mind we need to keep the final filename within TITLE_MAXLEN characters.

    max_paths = 10000
    path_names_set, path_names_lock = _unique_file_path_sets[p.parent]

    with path_names_lock:
        unique_path = _get_unique_path(p)
        if unique_path:
            return unique_path

    if date_hint:
        # Try using the date hint to create a unique filename.
        p = p.with_stem(f"{p.stem} {date_hint:%Y%m%d%H%M%S}")
        with path_names_lock:
            unique_path = _get_unique_path(p)
            if unique_path:
                return unique_path

    errmsg = f"{max_paths} copies of '{p}' already exist?!"
    raise ExportError(errmsg)


def format_seconds_dhms(seconds: int, *, with_sec: bool = True) -> tuple[str, datetime]:
    """Format a duration given in seconds into a human-readable string.

    Args:
        seconds (int): seconds
        with_sec (bool, optional): whether to include seconds in the formatted string. Defaults to True.

    Returns:
        tuple[str, datetime]: (human-readable duration, future datetime)
    """
    human_parts = []
    days, remainder = divmod(seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, secs = divmod(remainder, 60)

    if days > 0:
        human_parts.append(f"{days}d")
    if hours > 0:
        human_parts.append(f"{hours}h")
    if minutes > 0:
        human_parts.append(f"{minutes}m")
    if with_sec and secs > 0:
        human_parts.append(f"{secs}s")

    human = " ".join(human_parts) if human_parts else "0s"

    return human, datetime_now() + timedelta(seconds=seconds)
