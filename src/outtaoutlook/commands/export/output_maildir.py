"""Export messages to Maildir."""

import hashlib
import logging
import mailbox
import os
import struct
import threading
from pathlib import Path
from typing import TYPE_CHECKING, Self

from outtaoutlook.graphapi import GraphMessageT, log_message_notice
from outtaoutlook.shared import VERBOSE_LEVEL_MORE, logger, safe_file_name

from .output_core import ExportStat
from .output_queued import OutputWorkerQueued

if TYPE_CHECKING:
    from datetime import datetime

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


######################################################################
# Globals


# If the root Maildir does not exist, let the first thread create it.
thread_maildir_lock = threading.Lock()


######################################################################
# MaildirWrapper class


class MaildirWrapper:
    """Wraps `mailbox.Maildir` to add lock and context management."""

    def __init__(  # noqa: D107
        self,
        path: Path,
        *,
        create: bool = True,
        dry_run: bool,
        maildir: mailbox.Maildir | None = None,
    ) -> None:
        self.path = path
        self._init_create = create
        self.dry_run = dry_run
        self.maildir: mailbox.Maildir | None = maildir
        self._uniq_count = 1

    ######################################################################
    # Maildir methods

    def lock(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("maildir('%s'): locked (dry run)", self.path.name)
        elif self.maildir is None:
            msg = "Failed to lock Maildir: .maildir is not initialized"
            raise ValueError(msg)
        else:
            self.maildir.lock()

    def add(self, message: mailbox.Message, dt: datetime, *, mdid: str | None = None) -> str:
        """Add message to Maildir.

        For dry runs, a message key is returned but a
        message file is not written so `get_message` will fail.

        Args:
            message (mailbox.Message): The message.
            dt (datetime): The date to use for the message file's timestamp.
            mdid (str | None): Optional ID to uniquely identify the message.

        Raises:
            ValueError: If the Maildir is not initialized.
            mailbox.ExternalClashError: If a message with the same name already exists.

        Returns:
            str: Unique message key.
        """
        # We override Maildir.add so we can dictate the file name and timestamp.
        ts = dt.timestamp()
        uniq_name = self._get_uniq_name(message, ts, mdid)

        if not self.dry_run:
            if self.maildir is None:
                errmsg = "Failed to add message to Maildir: .maildir is not initialized"
                raise ValueError(errmsg)

            # Newly delivered, unread messages are saved to "/new"
            add_path = Path(self.maildir._path) / "new" / uniq_name  # noqa: SLF001

            if add_path.exists():
                errmsg = f"Name clash with existing message: '{add_path.name}'"
                raise mailbox.ExternalClashError(errmsg)

            add_path.write_bytes(message.as_bytes())
            os.utime(add_path, (ts, ts))

        return uniq_name

    def add_folder(self, folder: str) -> MaildirWrapper:  # noqa: D102
        if self.dry_run:
            return MaildirWrapper(self.path / folder, dry_run=self.dry_run)

        if self.maildir is None:
            msg = "Failed to add folder to Maildir: .maildir is not initialized"
            raise ValueError(msg)

        maildir = self.maildir.add_folder(folder)

        return MaildirWrapper(path=self.path / folder, dry_run=self.dry_run, maildir=maildir)

    def get_folder(self, folder: str) -> MaildirWrapper:  # noqa: D102
        if self.dry_run:
            return MaildirWrapper(self.path / folder, dry_run=self.dry_run)

        if self.maildir is None:
            msg = "Failed to get folder from Maildir: .maildir is not initialized"
            raise ValueError(msg)

        maildir = self.maildir.get_folder(folder)

        return MaildirWrapper(path=self.path / folder, dry_run=self.dry_run, maildir=maildir)

    def get_message(self, key: str) -> mailbox.MaildirMessage | None:  # noqa: D102
        if self.dry_run:
            return None

        if self.maildir is None:
            msg = "Failed to get message from Maildir: .maildir is not initialized"
            raise ValueError(msg)

        return self.maildir.get_message(key)

    def flush(self) -> None:  # noqa: D102
        if self.dry_run:
            # logger.debug("maildir('%s'): flushed (dry run)", self.path.name)
            pass
        elif self.maildir is None:
            msg = "Failed to flush Maildir: .maildir is not initialized"
            raise ValueError(msg)
        else:
            self.maildir.flush()

    def unlock(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("maildir('%s'): unlocked (dry run)", self.path.name)
        elif self.maildir is None:
            msg = "Failed to unlock Maildir: .maildir is not initialized"
            raise ValueError(msg)
        else:
            self.maildir.unlock()

    def close(self) -> None:  # noqa: D102
        if self.dry_run:
            logger.debug("maildir('%s'): closed (dry run)", self.path.name)
        elif self.maildir is None:
            msg = "Failed to close Maildir: .maildir is not initialized"
            raise ValueError(msg)
        else:
            self.maildir.close()

    ######################################################################
    # Context manager methods

    def __enter__(self) -> Self:  # noqa: D105
        if self.dry_run:
            logger.debug("maildir('%s'): opened and locked (dry run)", self.path.name)
        elif self.maildir is None:
            self.init_maildir(create=self._init_create)

        if self.maildir is not None:
            self.lock()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:  # noqa: ANN001, D105
        if self.dry_run:
            logger.debug("maildir('%s'): closed (dry run)", self.path.name)
        elif self.maildir is not None:
            try:
                self.close()
            finally:
                self.maildir = None

    ######################################################################
    # Public methods

    def init_maildir(self, *, create: bool = True) -> Self:
        """Create wrapped Maildir object.

        Args:
            create (bool): Whether to create the Maildir if it does not exist. Defaults to True.

        Returns:
            Self

        Raises:
            mailbox.NoSuchMailboxError: If the Maildir does not exist and `create` is False.
            mailbox.Error: If the Maildir could not be created.
        """
        if not create or self.path.is_dir():
            self.maildir = mailbox.Maildir(self.path, create=False)
        elif self.dry_run:
            logger.debug("maildir('%s'): created (dry run)", self.path.name)
        else:
            self.maildir = mailbox.Maildir(self.path, create=create)

        return self

    ######################################################################
    # Internal methods

    def _get_uniq_name(self, message: mailbox.Message, ts: float, mdid: str | None = None) -> str:
        """Create a unique filename for adding a message to the Maildir."""
        # Maildir standard requires format "timestamp.uuid.hostname"
        # uuid and hostname can also contain periods.

        if not mdid:
            header_bytes = message.as_bytes().split(b"\n\n", 1)[0]

            hasher = hashlib.sha256()
            hasher.update(struct.pack(">I", self._uniq_count))  # allow adding same email twice
            hasher.update(header_bytes)
            mdid = hasher.hexdigest()[:32]

            self._uniq_count += 1

        return f"{int(ts)}.{mdid}.local"


######################################################################
# Public module class


class OutputWorkerMaildir(OutputWorkerQueued):
    """Implements the `OutputWorkerQueued` for Maildir output."""

    def __init__(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Object constructor."""
        super().__init__(*args, **kwargs)

        # Inherited properties
        self.write_messages_status = "writing to maildir"

        # Maildir for our folder.
        # Only set during the export loop.
        self.maildir: MaildirWrapper | None = None

    ######################################################################
    # Method overrides

    def update_export_opts(self) -> None:
        """Set and optionally update export options for messages in the current folder."""
        super().update_export_opts()

        # `self.export_opts.files_dir` is the folder path, in the traditional sense,
        # which we use to *find* the actual Maildir directory in which subfolders begin with "."

        # We don't export attachments
        if self.export_opts.with_attachments:
            self.emit_log("Attachments will not be exported for Maildir output.")
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
        assert self.maildir is None  # noqa: S101

        try:
            with thread_maildir_lock:
                self.maildir = self.create_maildir()

        except mailbox.NoSuchMailboxError as err:
            errmsg = f"Failed to create Maildir '{self.export_opts.files_dir}': {err}"
            self.set_failed(errmsg, raise_now=True)
        except mailbox.Error as err:
            errmsg = f"Maildir operation failed '{self.export_opts.files_dir}': {err}"
            self.set_failed(errmsg, raise_now=True)

        # Note: Attachments directory will be created the first time an attachment is saved.
        if self.export_opts.with_attachments:
            self.emit_log(f"Saving attachments to '{self.export_opts.attachments_folder}'")

    def export_loop(self, *args, **kwargs) -> None:  # noqa: ANN002, ANN003
        """Internal loop for exporting messages."""
        if self.maildir is None:
            msg = "Failed to write message to Maildir: .maildir is not initialized"
            raise ValueError(msg)

        try:
            with self.maildir:
                super().export_loop(*args, **kwargs)

        except mailbox.Error as err:
            errmsg = f"Maildir operation failed '{self.export_opts.files_dir}': {err}"
            self.set_failed(errmsg, raise_now=True)
        finally:
            self.maildir = None

    def write_queued_messages(self) -> int:
        """Write queued messages to file.

        Returns:
            int: The number of files written.
        """
        if self.maildir is None:
            msg = "Failed to write message to Maildir: .maildir is not initialized"
            raise ValueError(msg)

        n_written = super().write_queued_messages()

        # Flush after every batch of messages
        self.maildir.flush()

        return n_written

    def write_queued_message(
        self: Self,
        message: GraphMessageT,
        file_path: Path,  # noqa: ARG002
        attachment_paths: list[Path] | None,  # noqa: ARG002
        content: bytes | None,
    ) -> int:
        """Export a message to a Maildir.

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

        if self.maildir is None:
            msg = "Failed to write message to Maildir: .maildir is not initialized"
            raise ValueError(msg)

        if content is None:
            return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)

        file_date = message["messageDateTime"]

        md_message = mailbox.Message(mailbox.Message(content))

        try:
            self.maildir.add(md_message, file_date, mdid=message["ooid"])
        except mailbox.ExternalClashError:
            errmsg = f"Message already exists in Maildir: {message['ooid']}: '{message['subject']}'"
            self.emit_log(errmsg, logging.WARNING)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        self.emit_log(
            f"Added message to Maildir: '{message['subject']}'" + (" (dry run)" if self.export_opts.dry_run else ""),
            VERBOSE_LEVEL_MORE,
        )

        return 1

    ######################################################################
    # Internal methods

    def create_maildir(self) -> MaildirWrapper:
        """Create root Maildir directory and any subfolders.

        Returns:
            MaildirWrapper

        Raises:
            ExportError: Raised if a path part is not safe for use as a directory name.
        """
        if not self.get_folder_path(self.export_opts.root_out_dir):
            self.emit_log("Writing messages to a new Maildir", logging.INFO)
        elif not self.get_folder_path(self.export_opts.files_dir):
            self.emit_log("Writing messages to a new folder in an existing Maildir", logging.INFO)
        else:
            self.emit_log("Writing messages to an existing Maildir folder", logging.INFO)

        maildir = MaildirWrapper(self.export_opts.root_out_dir, dry_run=self.export_opts.dry_run)
        maildir.init_maildir(create=True)

        # Each "subfolder" is its own Maildir.
        if self.export_opts.files_dir != self.export_opts.root_out_dir:
            maildir = self.add_folder_path(maildir, self.export_opts.files_dir)

        return maildir

    def add_folder_path(self, maildir: MaildirWrapper, folder_path: Path) -> MaildirWrapper:
        """Add Maildir folders with the given (traditional) path.

        Args:
            maildir (MaildirWrapper): The root Maildir to start adding subfolders to.
            folder_path (Path): Full folder path, in the traditional sense. Must be relative to root Maildir.

        Returns:
            MaildirWrapper: Maildir of the leaf folder.

        Raises:
            ExportError: Raised if a path part is not safe for use as a directory name.
            mailbox.Error: General Maildir error occurs.
        """
        assert folder_path != self.export_opts.root_out_dir  # noqa: S101
        assert folder_path.is_relative_to(self.export_opts.root_out_dir)  # noqa: S101

        for subfolder in folder_path.relative_to(self.export_opts.root_out_dir).parts:
            if subfolder != safe_file_name(subfolder, ""):
                self.set_failed(f"'{subfolder}' cannot be used as a Maildir folder name.", raise_now=True)

            maildir = maildir.add_folder(subfolder)

        return maildir

    def get_folder_path(self, folder_path: Path) -> MaildirWrapper | None:
        """Find a Maildir corresponding to the given folder path.

        Args:
            folder_path (Path): Full folder path, in the traditional sense. Must be relative to root Maildir.

        Returns:
            MaildirWrapper | None: A MaildirWrapper if the folder exists, otherwise None.

        Raises:
            mailbox.Error: General Maildir error occurs.
        """
        try:
            maildir = MaildirWrapper(self.export_opts.root_out_dir, dry_run=self.export_opts.dry_run)
            maildir.init_maildir(create=False)

            if folder_path != self.export_opts.root_out_dir:
                assert folder_path.is_relative_to(self.export_opts.root_out_dir)  # noqa: S101
                for subfolder in folder_path.relative_to(maildir.path).parts:
                    maildir = maildir.get_folder(subfolder)

        except mailbox.NoSuchMailboxError:
            return None

        return maildir
