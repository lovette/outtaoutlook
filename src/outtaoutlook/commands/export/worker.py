"""outtaoutlook export command."""

import concurrent.futures
import copy
import logging
import sys
import threading
from contextlib import contextmanager
from typing import TYPE_CHECKING, Self, TypedDict

import requests
from requests import sessions
from rich.progress import BarColumn, Progress, ProgressColumn, SpinnerColumn, Task, TextColumn
from rich.text import Text

from outtaoutlook.graphapi import (
    AUTH_ERROR_USECOMMAND,
    GraphApiConfig,
    GraphAuthRequiredError,
    GraphFolderT,
    print_graph_request_exception,
)
from outtaoutlook.shared import (
    ExportCancelledError,
    ExportError,
    console,
    logger,
)

from .output_eml import OutputWorkerEml
from .output_html import OutputWorkerHtml
from .output_json import OutputWorkerJson
from .output_md import OutputWorkerMarkdown
from .shared import (
    MAX_EXPORT_WORKERS,
    ExportOptions,
    OutputFormatChoice,
)

if TYPE_CHECKING:
    from collections.abc import Generator, Sequence

    from .output_core import OutputWorkerCore


######################################################################
# Global variables


# Thread-local storage for per-thread data.
thread_local_data = threading.local()

# Each thread is assigned a unique counter ID to track its order of execution.
thread_id_counter_lock = threading.Lock()
thread_id_counter = 0


######################################################################
# Constants


FOLDER_ICON = "\N{OPEN FILE FOLDER}\N{NARROW NO-BREAK SPACE}"

EXPORT_MODULES: dict[OutputFormatChoice, type[OutputWorkerCore]] = {
    OutputFormatChoice.eml: OutputWorkerEml,
    OutputFormatChoice.md: OutputWorkerMarkdown,
    OutputFormatChoice.html: OutputWorkerHtml,
    OutputFormatChoice.json: OutputWorkerJson,
}


######################################################################
# Types and classes


class RemainingItemsColumn(ProgressColumn):
    """Renders remaining count/total, e.g. ' 90/100 remaining'."""

    def render(self: Self, task: Task) -> Text:  # noqa: D102
        # Prevent math bugs if total isn't set yet
        total = f"{int(task.total):,}" if task.total is not None else "--"
        remaining = f"{int(task.remaining):,}" if task.remaining is not None else "--"

        return Text(f"{remaining}/{total}", style="progress.remaining")


######################################################################
# Helpers


@contextmanager
def _logger_friendly_progress(*args, **kwargs) -> Generator[Progress]:  # noqa: ANN002, ANN003
    """Wrapper around rich.progress.Progress that redirects logger output through 'progress.console'."""
    # Emiting messages with standard Python logging breaks the Progress bar because logger handlers
    # uses the system stderr and Rich redirects stdout and stderr (by default but configurable) to the "console".
    # You can also use 'console.print', 'progress.print' and 'progress.log'.
    #
    # Logging breaks progress bar
    # https://github.com/Textualize/rich/issues/4235
    #
    # This context manager solution is modeled after this discussion comment:
    # https://github.com/Textualize/rich/discussions/1578#discussioncomment-5565014

    __logger = kwargs.pop("logger", logging.getLogger())
    __streamhandlers = [(handler, None) for handler in __logger.root.handlers if type(handler) is logging.StreamHandler]

    try:
        with Progress(*args, **kwargs) as progress:
            for i, (handler, _) in enumerate(__streamhandlers):
                prior_stderr = handler.setStream(sys.stderr)
                __streamhandlers[i] = handler, prior_stderr

            yield progress

    finally:
        for handler, prior_stderr in __streamhandlers:
            if prior_stderr is not None:
                handler.setStream(prior_stderr)


def _assign_thread_id() -> int:
    global thread_id_counter  # noqa: PLW0603

    if not getattr(thread_local_data, "thread_id", None):
        with thread_id_counter_lock:
            thread_id_counter += 1
            thread_local_data.thread_id = thread_id_counter

        new_name = f"Thread #{thread_local_data.thread_id}"
        # logger.debug("Setting thread name to '%s'", new_name)
        threading.current_thread().name = new_name

    return thread_local_data.thread_id


def _print_export_results(ows: Sequence[OutputWorkerCore]) -> None:
    ow_with_results = [ow for ow in ows if ow.has_results]
    if not ow_with_results:
        return  # nothing to report

    console.print()
    console.print("[bold]Results:")

    for ow in ow_with_results:
        console.print()
        console.print(f"[bold]'{ow.folder_path}'[/bold]:")

        if ow.export_result.failed or ow.export_result.cancelled:
            console.print("[red]- Export did not complete[/red]")

        for level, messages in ow.export_result.messages.items():
            for message in messages:
                if level == logging.ERROR:
                    console.print(f"[red]- {message}[/red]")
                elif level == logging.WARNING:
                    console.print(f"[yellow]- {message}[/yellow]")
                else:
                    console.print(f"- {message}")

        if ow.export_stats:
            for stat, stat_count in ow.export_stats.items():
                console.print(f"- {stat}: {stat_count:,}")


######################################################################
# Worker main


def _worker_main(
    ow: OutputWorkerCore,
) -> None:
    """Export folder worker thread entry point.

    Args:
        ow (OutputWorkerCore): An output worker instance.

    Raises:
        Expected exceptions are captured but anything else will bubble up to the main thread by `future.result`.

    Returns:
        None: This function does not return a value; it updates the export result within the output_worker instance.
    """
    # Do this first!
    thread_id = _assign_thread_id()

    try:
        with sessions.Session() as session:
            ow.export(thread_id, session)

    except requests.RequestException as err:
        print_graph_request_exception(ow.progress.console, err)
        ow.set_failed("Encountered a HTTP related exception.")
    except GraphAuthRequiredError as err:
        # This should have triggered in main thread when calling `get_graph_config`...
        logger.error("%s %s", str(err), AUTH_ERROR_USECOMMAND)
        ow.set_failed("Encountered an authentication related exception.")
    except ExportCancelledError:
        # Whichever thread signaled the cancellation already logged the reason...
        pass
    except ExportError as err:
        logger.error("%s", str(err))
        ow.set_failed("Encountered an export related exception.")
    except OSError as err:
        logger.error("%s", str(err))
        ow.set_failed("Encountered an OS related exception.")
    except ValueError as err:
        logger.error("%s", str(err))
        ow.set_failed("Encountered a general exception.")
    finally:
        # First thread must always release the other workers, even on error!
        ow.release_workers()


######################################################################
# Public functions


def start_workers(
    folders: list[GraphFolderT],
    graph_config: GraphApiConfig,
    export_opts: ExportOptions,
    max_workers: int | None,
) -> None:
    """Export a list of folders, each in its own thread.

    Args:
        folders (list[GraphFolderT]): List of folder objects.
        graph_config (GraphApiConfig): Graph API configuration.
        export_opts (ExportOptions): Options for exporting messages.
        max_workers (int or None): Maximum size of thread pool or None to decide based on hardware.
    """

    # Keyword arguments to `_worker_main`.
    class _WorkerArgs(TypedDict):
        ow: OutputWorkerCore

    output_worker_class = EXPORT_MODULES[export_opts.out_format]
    workers: dict[concurrent.futures.Future[None], OutputWorkerCore] = {}

    if export_opts.filter_test:
        progress_cols = [
            TextColumn("[bold]{task.description}"),
            SpinnerColumn(finished_text=""),
            TextColumn("{task.fields[status]}"),  # our custom field
        ]
    else:
        progress_cols = [
            TextColumn("{task.description}"),
            BarColumn(),
            SpinnerColumn(),
            RemainingItemsColumn(),  # TextColumn("{task.remaining:g}"),
            TextColumn("{task.fields[status]}"),  # our custom field
        ]

    if max_workers and max_workers >= MAX_EXPORT_WORKERS:
        max_workers = None  # let ThreadPoolExecutor decide

    with (
        concurrent.futures.ThreadPoolExecutor(max_workers) as executor,
        _logger_friendly_progress(
            *progress_cols,
            console=console,
            logger=logger,
        ) as progress,
    ):
        logger.debug("Starting workers; pool size is %d", executor._max_workers)  # noqa: SLF001

        for folder in folders:
            # Add tasks to progress bar *before* starting.
            # This keeps the task bars in hierarchical order and ensures
            # console output doesn't interrupt the progress display.
            progress_task_id = progress.add_task(
                f"{FOLDER_ICON}{folder['folderPath']}",
                start=False,
                total=None,
                visible=True,
                status="pending",
            )

            ow = output_worker_class(
                folder=folder,
                graph_config=graph_config,
                progress=progress,
                progress_task_id=progress_task_id,
                export_opts=copy.copy(export_opts),  # each gets their own copy
            )

            worker_args: _WorkerArgs = {
                "ow": ow,
            }

            future = executor.submit(_worker_main, **worker_args)

            workers[future] = ow

        shutdown_pending = False

        # Start timer now so we can turn off the spinner if export is cancelled.
        for progress_task_id in progress.task_ids:
            progress.start_task(progress_task_id)

        for future, ow in workers.items():
            # We are looping through workers sequentially instead of using `as_completed`
            # because it hangs after we call `shutdown` :/
            # https://tiewkeehui.cc/blog/python-thread-pool-executor/
            # https://stackoverflow.com/q/66923690

            try:
                if future.cancelled():
                    ow.set_cancelled("Future was cancelled")
                else:
                    # Wait for the worker to complete.
                    # Exceptions not handled by the worker will be re-raised here.
                    future.result()

            except concurrent.futures.CancelledError:
                # This will never happen because we skip cancelled futures before calling result()
                ow.set_cancelled("Worker result was a futures.CancelledError.")
            else:
                result_type = ", ".join(
                    filter(
                        None,
                        [
                            "failed" if ow.export_result.failed else None,
                            "cancelled" if ow.export_result.cancelled else None,
                            "stats" if ow.export_stats else None,
                            "messages" if ow.export_result.messages else None,
                        ],
                    ),
                )

                ow.emit_log(f"Worker result is available ({result_type})", logging.DEBUG)

            if ow.workers_cancelled and not shutdown_pending:
                # Cancel any workers that have not run yet
                ow.emit_log("Workers cancelled, shutting down executor pool", logging.DEBUG)
                executor.shutdown(wait=False, cancel_futures=True)
                shutdown_pending = True

    logger.debug("All workers are finished!")

    _print_export_results(list(workers.values()))
