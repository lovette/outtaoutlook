"""outlook2obsidian ls command."""

from pathlib import Path
from typing import Annotated

import requests
import typer
from requests import sessions

from outlook2obsidian.shared import (
    AUTH_ERROR_USECOMMAND,
    AuthRequiredError,
    CommandContext,
    GraphFolderNotFoundError,
    console,
    err_console,
    get_graph_config,
    graph_resolve_folder_paths,
    print_graph_request_exception,
)

app = typer.Typer(add_completion=False)


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

        with sessions.Session() as session:
            folders = graph_resolve_folder_paths(
                session,
                graph_config,
                [folder_path] if folder_path else None,
                maxdepth=maxdepth,
            )

    except requests.RequestException as err:
        print_graph_request_exception(err)
        raise typer.Exit(code=1) from err
    except AuthRequiredError as err:
        err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
        raise typer.Exit(code=1) from err
    except (GraphFolderNotFoundError, ValueError, OSError) as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err

    console.print(f"Mailbox '{graph_config.token_info.username}':")

    for folder in folders:
        depth = len(Path(folder["folderPath"]).parts)
        console.print(f"{'  ' * (depth - 1)}|- '{folder['displayName']}' ({folder['totalItemCount']})")
