"""outtaoutlook export command."""

import concurrent.futures
import logging
import sys
import threading
from contextlib import contextmanager
from typing import TYPE_CHECKING, Self, TypedDict

import requests
import typer
from requests import sessions
from rich.progress import BarColumn, Progress, ProgressColumn, SpinnerColumn, Task, TaskID, TextColumn
from rich.text import Text

from outtaoutlook.graphapi import (
    AUTH_ERROR_USECOMMAND,
    GraphApiConfig,
    GraphAuthRequiredError,
    GraphFolderT,
    print_graph_request_exception,
)
from outtaoutlook.shared import (
    ExportAbortedError,
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
    ExportResult,
    OutputFormatChoice,
    abort_workers_event,
    list_messages_success_event,
    signal_list_messages_success,
)

if TYPE_CHECKING:
    from collections.abc import Generator

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


LIST_MESSAGES_SUCCESS_EVENT_WAIT = 60  # seconds; just a failsafe, I don't expect there to be any deadlocking :)
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


def _signal_abort_workers() -> None:
    """Set event that signals all workers to stop exporting messages."""
    if not abort_workers_event.is_set():
        logger.debug("Setting all workers abort flag")
        abort_workers_event.set()


def _emit_log(folder: GraphFolderT, errmsg: str, level: int = logging.INFO) -> None:
    logger.log(level, 'Folder("%s"): %s', folder["folderPath"], errmsg)


def _assign_thread_id() -> int:
    global thread_id_counter  # noqa: PLW0603

    if not getattr(thread_local_data, "thread_id", None):
        with thread_id_counter_lock:
            thread_id_counter += 1
            thread_local_data.thread_id = thread_id_counter

        new_name = f"Thread #{thread_local_data.thread_id}"
        logger.debug("Setting thread name to '%s'", new_name)
        threading.current_thread().name = new_name

    return thread_local_data.thread_id


def _wait_first_batch_release(folder: GraphFolderT) -> bool:
    thread_id = _assign_thread_id()

    if thread_id > 1 and not list_messages_success_event.is_set():
        _emit_log(folder, "Waiting for first batch release", logging.DEBUG)

        if not list_messages_success_event.wait(timeout=LIST_MESSAGES_SUCCESS_EVENT_WAIT):
            _emit_log(folder, "Timed out waiting for release.", logging.DEBUG)
            _signal_abort_workers()
            return False

    return True


def _print_export_results(results: dict[str, tuple[GraphFolderT, ExportResult | None]]) -> None:  # noqa: C901
    # Filter export_results to items with messages or stats
    results = {
        folder_id: (folder, export_result)
        for folder_id, (folder, export_result) in results.items()
        if export_result and (export_result.messages or export_result.stats)
    }

    if not results:
        return

    console.print()
    console.print("[bold]Results:")

    for folder, export_result in results.values():
        if export_result and (export_result.messages or export_result.stats):
            console.print()
            console.print(f"[bold]'{folder['folderPath']}'[/bold]:")

            if export_result.aborted:
                console.print("[red]- Export did not complete[/red]")

            for level, messages in export_result.messages.items():
                for message in messages:
                    if level == logging.ERROR:
                        console.print(f"[red]- {message}[/red]")
                    elif level == logging.WARNING:
                        console.print(f"[yellow]- {message}[/yellow]")
                    else:
                        console.print(f"- {message}")

            if export_result.stats:
                for stat, stat_count in export_result.stats.items():
                    console.print(f"- {stat}: {stat_count:,}")


######################################################################
# Worker main


def _worker_main(
    folder: GraphFolderT,
    *,
    progress: Progress,
    progress_task_id: TaskID,
    graph_config: GraphApiConfig,
    export_opts: ExportOptions,
) -> ExportResult:
    """Export folder worker thread entry point.

    Args:
        folder (GraphFolderT): Folder to export.
        progress (Progress): Progress bar.
        progress_task_id (TaskID): The task ID for the progress tracker.
        graph_config (GraphApiConfig): Graph API configuration.
        export_opts (ExportOptions): Options for exporting messages.

    Raises:
        Exceptions are not captured and will bubble up to the main thread by `future.result`.

    Returns:
        ExportResult: The result of the export operation, including statistics and any status message.
    """
    # Do this first!
    _assign_thread_id()

    # Wait for the first/primary thread/worker to get its first batch of messages.
    # This allows the primary worker to encounter any issues with
    # the Graph API before other workers proceed and encounter the same.
    if not _wait_first_batch_release(folder):
        return ExportResult(aborted=True)

    # If the first batch of the first worker failed, or if any worker has since
    # signaled an abort, all workers should stop too.
    if abort_workers_event.is_set():
        return ExportResult(aborted=True)

    export_result = ExportResult()
    export_module_class = EXPORT_MODULES[export_opts.out_format]

    _emit_log(folder, "Working on this folder...", logging.DEBUG)

    # Reset task start time to get true elapsed time for this folder.
    progress.reset(progress_task_id, start=True)

    try:
        with sessions.Session() as session:
            export_module = export_module_class(
                thread_id=thread_local_data.thread_id,
                folder=folder,
                session=session,
                graph_config=graph_config,
                progress=progress,
                progress_task_id=progress_task_id,
                export_opts=export_opts,
            )

            # Assign stats to the result now so stats updated as progress is made
            # are reported even in the event an exception occurs.
            export_result.stats = export_module.export_stats

            return export_module.export()

    except requests.RequestException as err:
        print_graph_request_exception(progress.console, err)
        _emit_log(folder, "Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    except GraphAuthRequiredError as err:
        # This should have triggered in main thread when calling `get_graph_config`...
        logger.error("%s %s", str(err), AUTH_ERROR_USECOMMAND)
        _emit_log(folder, "Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    except ExportAbortedError:
        # Whichever thread signaled the abort already logged the reason...
        export_result.aborted = True
    except (ExportError, ValueError, OSError) as err:
        logger.error("%s", str(err))
        _emit_log(folder, "Encountered an exception.", logging.DEBUG)
        export_result.aborted = True
    finally:
        # Primary thread must always release the other workers!
        signal_list_messages_success(thread_local_data.thread_id)

    # The failure of one worker brings everything to a halt.
    if export_result.aborted:
        _signal_abort_workers()

    return export_result


######################################################################
# Public functions


def start_workers(  # noqa: C901, PLR0912, PLR0915
    folders: list[GraphFolderT],
    graph_config: GraphApiConfig,
    export_opts: ExportOptions,
    max_threads: int | None,
) -> None:
    """Export a list of folders, each in its own thread.

    Args:
        folders (list[GraphFolderT]): List of folder objects.
        graph_config (GraphApiConfig): Graph API configuration.
        export_opts (ExportOptions): Options for exporting messages.
        max_threads (int or None): Maximum size of thread pool or None to decide based on hardware.
    """

    # Keyword arguments to `_export_folder_worker`.
    class _WorkerArgs(TypedDict):
        progress: Progress
        progress_task_id: TaskID
        graph_config: GraphApiConfig
        export_opts: ExportOptions

    show_progress = True

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

    export_results: dict[str, tuple[GraphFolderT, ExportResult | None]] = {}  # {id:(folder, result)}

    if max_threads and max_threads >= MAX_EXPORT_WORKERS:
        max_threads = None  # let ThreadPoolExecutor decide

    with (
        concurrent.futures.ThreadPoolExecutor(max_threads) as executor,
        _logger_friendly_progress(*progress_cols, console=console, logger=logger, disable=not show_progress) as progress,
    ):
        folder_task_id: dict[str, tuple[GraphFolderT, TaskID]] = {}  # {id: (folder,task_id)}
        workers: dict[concurrent.futures.Future[ExportResult], GraphFolderT] = {}
        future_task_id: dict[concurrent.futures.Future[ExportResult], TaskID] = {}

        logger.debug("Worker thread pool size: %d", executor._max_workers)  # noqa: SLF001

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

            worker_args: _WorkerArgs = {
                "progress": progress,
                "progress_task_id": progress_task_id,
                "graph_config": graph_config,
                "export_opts": export_opts,
            }

            future = executor.submit(_worker_main, folder, **worker_args)

            workers[future] = folder
            future_task_id[future] = progress_task_id
            folder_task_id[folder["id"]] = (folder, progress_task_id)
            export_results[folder["id"]] = (folder, None)

        shutdown_pending = False
        progress_tasks: dict[TaskID, Task] = {}
        progress_tasks = {task.id: task for task in progress.tasks}

        # We loop through workers sequentially instead of using `as_completed` because it hangs after we call `shutdown` :/
        # https://tiewkeehui.cc/blog/python-thread-pool-executor/
        # https://stackoverflow.com/questions/66923690/does-concurrent-futures-as-completed-yield-for-cancelled-futures

        logger.debug("Starting worker threads.")

        # Start timer now so we can turn off the spinner if aborted.
        for progress_task_id in progress.task_ids:
            progress.start_task(progress_task_id)

        for worker, folder in workers.items():
            progress_task_id = future_task_id[worker]

            if worker.cancelled():
                logger.debug('Folder("%s"): Future was cancelled', folder["folderPath"])
                progress.update(progress_task_id, status="Cancelled")
                continue

            worker_exception = False

            try:
                export_result = worker.result()
            except concurrent.futures.CancelledError:
                # This will never happen because we skip cancelled futures before calling result()
                worker_exception = True
            else:
                export_results[folder["id"]] = (folder, export_result)

                result_type = [
                    "aborted" if export_result.aborted else None,
                    "stats" if export_result.stats else None,
                    "messages" if export_result.messages else None,
                ]

                logger.debug(
                    'Folder("%s"): Worker result is available (%s)',
                    folder["folderPath"],
                    ",".join(filter(None, result_type)),
                )

                if export_result.aborted:
                    task = progress_tasks[progress_task_id]
                    progress.update(progress_task_id, completed=task.total, status="cancelled")

            if worker_exception and not abort_workers_event.is_set():
                logger.debug('Folder("%s"): Worker resulted in an exception.', folder["folderPath"])
                _signal_abort_workers()

            if abort_workers_event.is_set() and not shutdown_pending:
                # Cancel any workers that have not run yet
                logger.debug('Folder("%s"): Abort signal is set, shutting down worker threads.', folder["folderPath"])
                executor.shutdown(wait=False, cancel_futures=True)
                shutdown_pending = True

    logger.debug("All worker threads finished.")

    _print_export_results(export_results)

    if abort_workers_event.is_set():
        raise typer.Exit(code=1)
