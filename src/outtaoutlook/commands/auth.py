"""outtaoutlook auth command."""

from pathlib import Path
from typing import TYPE_CHECKING, Annotated

import requests
import typer

from outtaoutlook.shared import (
    AUTH_ERROR_USECOMMAND,
    AUTH_NO_CACHE_ERROR,
    GRAPH_SCOPES,
    CommandContext,
    GraphApiError,
    GraphAuthRequiredError,
    console,
    err_console,
    get_graph_client,
    get_token_cache_info,
    logger,
)

if TYPE_CHECKING:
    from pathlib import Path

app = typer.Typer(add_completion=False)


######################################################################
# Helpers


def _raise_if_error(errmsg: str, result: dict | None) -> None:
    """Raise a GraphApiError if the result dictionary contains an error or is None.

    Args:
        errmsg (str): The error message to include if an error is raised.
        result (dict | None): The result dictionary to check for errors.

    Raises:
        GraphApiError
    """
    if not result:
        raise GraphApiError(errmsg)

    if "error" in result:
        errmsg = f"{errmsg}: {result.get('error')}; {result.get('error_description')}"
        raise GraphApiError(errmsg)


def _auth_info(auth_file_path: Path) -> None:
    # First, see if access was previously granted and cached.
    cache_info = get_token_cache_info(auth_file_path)
    if not cache_info:
        console.print(AUTH_NO_CACHE_ERROR)
        raise typer.Exit

    # If cache exists, see if cached tokens are valid.
    try:
        get_graph_client(auth_file_path, use_cache=True)
    except GraphAuthRequiredError as err:
        console.print(err)
        raise typer.Exit from err

    console.print("Access to Microsoft Outlook has been granted to:")
    console.print(f"User: {cache_info.username}")
    console.print(f"Tenant: {cache_info.tenant}")
    console.print(f"Client: {cache_info.client}")
    console.print(f"Cached at: {cache_info.cached_at}")


def _auth_init(tenant: str | None, client: str | None, auth_file_path: Path, *, force: bool) -> None:
    token = None
    accounts = None

    app, cache, account, token = get_graph_client(
        auth_file_path,
        tenant=tenant,
        client=client,
        use_cache=not force,
    )

    if not token:
        # 1. Initiate the user authentication browser flow
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        _raise_if_error("Failed to initiate user authentication flow", flow)

        # 2. Print the instructions to the terminal for the user
        console.print(f"You are authorizing client '{client}' to access your account in tenant '{tenant}'.")
        console.print(flow["message"])

        # 3. Block and wait for the user to complete authentication in their browser
        token = app.acquire_token_by_device_flow(flow)
        _raise_if_error("Access was not granted", token)

    if not token or "access_token" not in token:
        _raise_if_error("Failed to acquire access token", None)

    if not account:
        accounts = app.get_accounts()

        if not accounts:
            _raise_if_error("Failed to acquire account information", None)

        account = accounts[0]

    console.print(f"Successfully authenticated as '{account['username']}'.")

    if cache.has_state_changed:
        auth_file_path.write_text(cache.serialize(), encoding="utf-8")
        logger.info("Token cache written to '%s'", auth_file_path)


######################################################################
# Commands


@app.command()
def auth(
    ctx: typer.Context,
    tenant: Annotated[
        str | None,
        typer.Option(
            help="Tenant identifier domain or UUID.",
            metavar="TENANTID",
            envvar="OUTTAOUTLOOK_TENANTID",
            show_envvar=False,
        ),
    ] = None,
    client: Annotated[
        str | None,
        typer.Option(
            help="Application client identifier registered with Microsoft Entra admin center.",
            metavar="CLIENTID",
            envvar="OUTTAOUTLOOK_CLIENTID",
            show_envvar=False,
        ),
    ] = None,
    *,
    force: Annotated[
        bool,
        typer.Option(
            "--force",
            help="Force re-authentication even if a valid token is cached.",
        ),
    ] = False,
    info: Annotated[
        bool,
        typer.Option(
            "--info",
            help="Show cached authentication information and exit.",
        ),
    ] = False,
) -> None:
    """Sign in and grant access to Microsoft Outlook via the Microsoft Graph API."""
    # Options common to all commands
    ctx_obj: CommandContext = ctx.obj

    assert ctx_obj is not None  # noqa: S101
    assert isinstance(ctx_obj, CommandContext)  # noqa: S101

    auth_file_path = ctx_obj.common_opts.auth_file_path

    if not auth_file_path:
        err_console.print("Internal error: Authentication file path is not set.")
        raise typer.Exit(1)

    if info:
        try:
            _auth_info(auth_file_path)
        except requests.RequestException as err:
            err_console.print(f"Error: {err}")
            raise typer.Exit(code=1) from err
        except GraphAuthRequiredError as err:
            err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
            raise typer.Exit(code=1) from err
        except GraphApiError as err:
            err_console.print(f"Error: {err}")
            raise typer.Exit(code=1) from err
    else:
        try:
            _auth_init(tenant, client, auth_file_path, force=force)
        except requests.RequestException as err:
            err_console.print(f"Error: {err}")
            raise typer.Exit(code=1) from err
        except GraphAuthRequiredError as err:
            err_console.print(f"{err}")
            raise typer.Exit(code=1) from err
        except GraphApiError as err:
            err_console.print(f"Error: {err}")
            raise typer.Exit(code=1) from err
