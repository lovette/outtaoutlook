"""outlook2obsidian ls command."""

from pathlib import Path
from typing import Annotated

import requests
import typer
from requests import sessions

from outlook2obsidian.shared import (
    AUTH_ERROR_USECOMMAND,
    MAILBOX_ROOT,
    AuthRequiredError,
    CommandContext,
    GraphApiConfig,
    GraphFolderT,
    console,
    err_console,
    get_graph_config,
    graph_enum_folders,
    graph_get_folder_by_name,
    graph_get_folder_by_path,
)

app = typer.Typer(add_completion=False)


######################################################################
# Helpers


def _print_folder_leaf(folder: GraphFolderT, depth: int) -> GraphFolderT:
    console.print(f"{'  ' * (depth - 1)}|- '{folder['displayName']}' ({folder['totalItemCount']})")
    return folder


def _print_folder_path(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_path: Path,
    depth: int,
    parent_folder: GraphFolderT | None,
) -> GraphFolderT | None:
    path_parts = list(Path(folder_path).parts)

    folder = graph_get_folder_by_name(session, graph_config, path_parts.pop(0), parent_folder)
    if not folder:
        return None  # No such folder

    _print_folder_leaf(folder, depth)

    if not path_parts:
        # Print children of the last part
        for subfolder in graph_enum_folders(session, graph_config, folder).values():
            _print_folder_leaf(subfolder, depth + 1)

        return folder

    return _print_folder_path(session, graph_config, Path(*path_parts), depth + 1, parent_folder=folder)


def _print_folder_tree(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder: GraphFolderT,
    depth: int,
    maxdepth: int,
) -> GraphFolderT | None:
    _print_folder_leaf(folder, depth)

    if depth < maxdepth:
        for subfolder in graph_enum_folders(session, graph_config, folder).values():
            _print_folder_tree(session, graph_config, subfolder, depth + 1, maxdepth)

    return folder


######################################################################
# Commands


@app.command()
def ls(
    ctx: typer.Context,
    folder_path: Annotated[
        str | None,
        typer.Argument(
            hidden=True,
            metavar="FOLDER",
        ),
    ] = None,
    maxdepth: Annotated[
        int,
        typer.Option(
            "--maxdepth",
            min=1,
            help="Maximum depth of folder tree to display.",
        ),
    ] = 10,
) -> None:
    """List mailbox folder tree."""
    # Options common to all commands
    ctx_obj: CommandContext = ctx.obj

    assert ctx_obj is not None  # noqa: S101
    assert isinstance(ctx_obj, CommandContext)  # noqa: S101

    common_opts = ctx_obj.common_opts

    try:
        graph_config = get_graph_config(common_opts)

        console.print(f"Mailbox '{graph_config.token_info.username}':")

        with sessions.Session() as session:
            if folder_path is not None:
                # First, validate the folder exists
                folder = graph_get_folder_by_path(session, graph_config, Path(folder_path), MAILBOX_ROOT)
                if not folder:
                    err_console.print(f"Error: '{folder_path}': No such folder.")
                    raise typer.Exit(code=1)

                _print_folder_path(session, graph_config, Path(folder_path), 1, MAILBOX_ROOT)

            else:
                for folder in graph_enum_folders(session, graph_config, MAILBOX_ROOT).values():
                    _print_folder_tree(session, graph_config, folder, 1, maxdepth)

    except requests.RequestException as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err
    except AuthRequiredError as err:
        err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
        raise typer.Exit(code=1) from err
    except ValueError as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err
