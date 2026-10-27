"""Export messages as raw .mbox files."""

import logging
import mailbox
import os
import sys
from datetime import datetime, timezone
from email.utils import parseaddr
from pathlib import Path
from typing import Self

from outtaoutlook.graphapi import GraphMessageT, log_message_notice
from outtaoutlook.shared import VERBOSE_LEVEL_MORE, logger, safe_file_name

from .output_core import TITLE_UNTITLED, ExportStat
from .output_queued import OutputWorkerQueued

######################################################################
# Constants


# Message fields we reference for exporting.
# We need subject for log messages and dates for file name and timestamp.
SELECT_MESSAGE_FIELDS = ",".join(
    sorted(
        [
            "id",
            "subject",
            "isDraft",
            "hasAttachments",
            "createdDateTime",
            "lastModifiedDateTime",
            "sentDateTime",
            "receivedDateTime",
        ],
    ),
)

ALL_MAIL_MBOX = "All Mail"

######################################################################
# MboxWrapper class


class MboxWrapper:
    """Wraps `mailbox.mbox` to add lock and context management."""

    def __init__(self, path: Path, *, dry_run: bool) -> None:  # noqa: D107
        self.path = path
        self.dotlockpath = path.with_name(path.name + ".lock")
        self.dry_run = dry_run
        self.mbox: mailbox.mbox | None = None

    ######################################################################
    # mbox methods

    def lock(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("'%s': locked (dry run)", self.path.name)
        elif self.mbox is None:
            msg = "Failed to lock mbox: .mbox is not initialized"
            raise ValueError(msg)
        else:
            self.mbox.lock()

    def add(self, message: mailbox.mboxMessage) -> None:  # noqa: D102
        if self.dry_run:
            pass
        elif self.mbox is None:
            msg = "Failed to add message to mbox: .mbox is not initialized"
            raise ValueError(msg)
        else:
            self.mbox.add(message)

    def flush(self) -> None:  # noqa: D102
        if self.dry_run:
            # logger.debug("'%s': flushed (dry run)", self.path.name)
            pass
        elif self.mbox is None:
            msg = "Failed to flush mbox: .mbox is not initialized"
            raise ValueError(msg)
        else:
            self.mbox.flush()

    def unlock(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("'%s': unlocked (dry run)", self.path.name)
        elif self.mbox is None:
            msg = "Failed to unlock mbox: .mbox is not initialized"
            raise ValueError(msg)
        else:
            self.mbox.unlock()

    def close(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("'%s': closed (dry run)", self.path.name)
        elif self.mbox is None:
            msg = "Failed to close mbox: .mbox is not initialized"
            raise ValueError(msg)
        else:
            self.mbox.close()

    ######################################################################
    # Additional methods

    def is_dotlocked(self) -> bool:  # noqa: D102
        return self.dotlockpath.exists() and self.dotlockpath.is_file()

    def undotlock(self) -> bool:  # noqa: D102
        try:
            self.dotlockpath.unlink(missing_ok=True)
        except OSError:
            return False

        return True

    ######################################################################
    # Context manager methods

    def __enter__(self) -> Self:  # noqa: D105
        if self.dry_run:
            logger.debug("'%s': opened and locked (dry run)", self.path.name)
        else:
            mbox = mailbox.mbox(self.path)
            try:
                self.mbox = mbox
                self.lock()
            except mailbox.Error:
                self.mbox = None
                raise

        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:  # noqa: ANN001, D105
        if self.dry_run:
            logger.debug("'%s': closed (dry run)", self.path.name)
        elif self.mbox is not None:
            try:
                self.close()
            finally:
                self.mbox = None


######################################################################
# Public module class


class OutputWorkerMbox(OutputWorkerQueued):
    """Implements the `OutputWorkerQueued` for .mbox output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

        # Inherited properties
        self.write_messages_status = "appending to mbox"

        # Our properties
        self._mbox: MboxWrapper | None = None
        self.mbox_path: Path | None = None
        self.mbox_date: datetime | None = None

        # Stopping execution while debugging abandons the dotlock file.
        self.ignore_dotlock = "debugpy" in sys.modules

    ######################################################################
    # Method overrides

    def update_export_opts(self) -> None:
        """Set and optionally update export options for messages in the current folder."""
        super().update_export_opts()

        mbox_dir = self.export_opts.root_out_dir

        if not self.export_opts.hier_flat:
            if len(self.folder_path.parts) > 1 and not self.export_opts.hier_shallow:
                # If mail folder path is "A/B/C", set mbox path to "A/B/C.mbox".
                mbox_dir = mbox_dir / Path(*self.folder_path.parts[:-1])

            mbox_name = safe_file_name(self.folder_name, TITLE_UNTITLED)
        else:
            if self.export_opts.max_workers > 1:
                self.set_failed("Too many workers for a single mbox.", raise_now=True)

            mbox_name = ALL_MAIL_MBOX

        self.mbox_path = mbox_dir / (mbox_name + ".mbox")

        # We don't export attachments
        if self.export_opts.with_attachments:
            self.emit_log("Attachments will not be exported for .mbox output.")
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
    ) -> Path:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.

        Returns:
            Path
        """
        return super().get_file_path(message)

    def prepare_export(self) -> None:
        """Prepare to export messages."""
        assert self.mbox_path is not None  # noqa: S101

        if self.mbox_path.is_file():
            self.emit_log(f"Writing messages to existing mbox '{self.mbox_path}'", logging.INFO)
        else:
            self._mkdir(self.mbox_path.parent, "mbox parent", parents=not self.export_opts.hier_flat)
            self.emit_log(f"Writing messages to new mbox '{self.mbox_path}'", logging.INFO)

        # Note: Attachments directory will be created the first time an attachment is saved.
        if self.export_opts.with_attachments:
            self.emit_log(f"Saving attachments to '{self.export_opts.attachments_folder}'")

    def export_loop(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Internal loop for exporting messages."""
        assert self.mbox_path is not None  # noqa: S101

        if self.mbox_path.is_file():
            mbox_mtime = self.mbox_path.stat().st_mtime
            self.mbox_date = datetime.fromtimestamp(mbox_mtime, tz=timezone.utc)  # noqa: UP017

        assert self._mbox is None  # noqa: S101
        mbox = MboxWrapper(self.mbox_path, dry_run=self.export_opts.dry_run)

        if mbox.is_dotlocked():
            if not self.ignore_dotlock:
                errmsg = f"mbox is locked with a lock file: '{mbox.dotlockpath.name}'"
                self.set_failed(errmsg, raise_now=True)
            elif not mbox.undotlock():
                errmsg = f"mbox is locked but the lock file cannot be deleted: '{mbox.dotlockpath.name}'"
                self.set_failed(errmsg, raise_now=True)
            else:
                self.emit_log("mbox is locked with a lock file but it was successfully deleted", logging.INFO)

        try:
            with mbox:
                self._mbox = mbox
                super().export_loop(*args, **kwargs)

        except mailbox.ExternalClashError as err:
            errmsg = f"Failed to lock mbox '{self.mbox_path}': {err}"
            self.set_failed(errmsg, raise_now=True)
        except mailbox.Error as err:
            errmsg = f"Failed to create mbox '{self.mbox_path}': {err}"
            self.set_failed(errmsg, raise_now=True)
        finally:
            self._mbox = None

    def write_queued_messages(self) -> int:
        """Write queued messages to file.

        Returns:
            int: The number of files written.
        """
        assert self.mbox_path is not None  # noqa: S101

        if self._mbox is None:
            msg = "Mailbox is not initialized"
            raise ValueError(msg)

        n_written = super().write_queued_messages()

        # Flush after every batch of messages
        self._mbox.flush()

        # Update the modification and access times of the mbox file to match the most recent message date.
        if not self.export_opts.dry_run and self.mbox_date is not None:
            os.utime(self.mbox_path, (self.mbox_date.timestamp(), self.mbox_date.timestamp()))

        return n_written

    def write_queued_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,  # noqa: ARG002
        attachment_paths: list[Path] | None,  # noqa: ARG002
        content: bytes | None,
    ) -> int:
        """Export a message to an .mbox file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): The path to the output file.
            attachment_paths (list[Path] | None): List of paths to attachments, or None if there are no attachments.
            content (bytes | None): The raw content of the message, or None if not available.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            self.export_stats.incr(skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return 1

        if self._mbox is None:
            msg = "Failed to write message to mbox: .mbox is not initialized"
            raise ValueError(msg)

        if content is None:
            return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)

        file_date = message["messageDateTime"]

        mailboxmsg = mailbox.mboxMessage(mailbox.Message(content))

        # Determine the envelope sender so the message delimiter is most accurate.
        sender = mailboxmsg["Return-Path"] or mailboxmsg["Sender"] or mailboxmsg["From"]
        if sender:
            sender = parseaddr(sender)[1]

        mailboxmsg.set_from(sender or "MAILER-DAEMON", file_date.timetuple())

        self._mbox.add(mailboxmsg)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        self.emit_log(
            f"Added message to mbox: '{message['subject']}'" + (" (dry run)" if self.export_opts.dry_run else ""),
            VERBOSE_LEVEL_MORE,
        )

        self.mbox_date = max(file_date, self.mbox_date) if self.mbox_date else file_date

        return 1
