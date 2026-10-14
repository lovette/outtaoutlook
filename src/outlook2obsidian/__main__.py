#!/usr/bin/env python3
"""Create Obsidian notes from Microsoft Outlook messages using the Microsoft Graph API."""

import copy
import importlib.metadata
import json
import logging
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any

import typer
import yaml

from .commands import auth_app, export_app, ls_app
from .shared import APP_NAME, VERBOSE_LOGGING_LEVEL_ERROR, CommandContext, console, path_with_vars

if TYPE_CHECKING:
    from io import TextIOWrapper

TOKEN_CACHE_FILE_NAME = f".{APP_NAME}_msal_auth.json"
CONFIG_FILE_NAME = f"{APP_NAME}_config.yaml"

######################################################################
# Globals

# Reference version number in pyproject.toml
# (For unexplained reasons, will change "-dev" suffix to "-dev0".)
__version__ = importlib.metadata.version(APP_NAME)


######################################################################
# Logging

VERBOSE_LOGGING_LEVELS = (
    logging.ERROR,  # default
    logging.WARNING,  # -v
    logging.INFO,  # -vv
    logging.DEBUG,  # -vvv
)

LOGGING_FORMAT = "[%(levelname)s] %(message)s"

logging.basicConfig(level=VERBOSE_LOGGING_LEVELS[0], format=LOGGING_FORMAT)

# Use a named logger instead of root logger
logger = logging.getLogger(APP_NAME)


######################################################################
# YAML UniqueKeyLoader
#
# Thanks to the conversation at:
# https://gist.github.com/pypt/94d747fe5180851196eb


class UniqueKeyLoader(yaml.SafeLoader):
    """Load YAML and fail if key is duplicated."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict:  # noqa: FBT001, FBT002
        """Internal mapping function.

        Args:
            node (yaml.MappingNode): ??
            deep (bool): ??

        Raises:
            ValueError: Duplicate key found.

        Returns:
            dict
        """
        mapping = set()

        for key_node, _value_node in node.value:
            if ":merge" not in key_node.tag:
                key = self.construct_object(key_node, deep=deep)
                if key in mapping:
                    msg = f"{key!r}: Duplicate key found in YAML."
                    raise ValueError(msg)
                mapping.add(key)

        return super().construct_mapping(node, deep)


######################################################################
# Helpers


def _read_config_file(config_path: Path) -> dict[str, Any]:
    """Open and parse a JSON or YAML configuration file.

    Args:
        config_path (Path): Config file path.

    Raises:
        typer.BadParameter: Failed to open or parse file.

    Returns:
        dict; empty if file does not exist or is empty.
    """

    def _read_json_config(fp: TextIOWrapper) -> dict[str, Any]:
        try:
            d = json.load(fp)
            logger.debug("Reading JSON file '%s'", config_path)
        except json.JSONDecodeError as err:
            msg = f"{err}"
            raise typer.BadParameter(msg) from err

        if not isinstance(d, dict):
            msg = "Configuration file must be a JSON set of key/value pairs."
            raise typer.BadParameter(msg)

        return d

    def _read_yaml_config(fp: TextIOWrapper) -> dict[str, Any]:
        try:
            logger.debug("Reading YAML file '%s'", config_path)
            d = yaml.load(fp.read(), Loader=UniqueKeyLoader)  # noqa: S506
        except (yaml.YAMLError, ValueError) as err:
            msg = f"{err}"
            raise typer.BadParameter(msg) from err

        if not isinstance(d, dict):
            msg = "Configuration file must be a YAML set of key/value pairs."
            raise typer.BadParameter(msg)

        return d

    try:
        if config_path.is_file() and config_path.stat().st_size:
            with config_path.open() as fp:
                return _read_json_config(fp) if config_path.suffix == ".json" else _read_yaml_config(fp)

    except OSError as err:
        msg = f"{err}"
        raise typer.BadParameter(msg) from err

    return {}


def _opt_version_callback(value: bool) -> None:  # noqa: FBT001
    """Prints the application version and exits if --version is provided."""
    if value:
        console.print("mnycsvqif", __version__)
        raise typer.Exit(0)


def _opt_config_file_callback(  # noqa: C901, PLR0915
    ctx: typer.Context,
    param: typer.CallbackParam,  # noqa: ARG001
    config_file_or_dir: Path | None,
) -> Path | None:
    """Handle `--config-file` command option.

    Note:
        The values of `default_map` are used as the default values for command options.
        This function is called even if `--config-file` is not specified.

    Args:
        ctx (typer.Context): Typer context object.
        param (typer.CallbackParam): Typer callback parameter object.
        config_file_or_dir (Path): Option value or None if option is not specified.

    Returns:
        None
    """

    def _validate_main_opts(default_map: dict[str, Any]) -> None:
        for opt in ("auth_file", "cache_dir"):
            opt_value = default_map.get(opt)
            if opt_value:
                default_map[opt] = path_with_vars(opt_value)

    def _validate_auth_opts(default_map: dict[str, Any]) -> None:
        auth_cmd_config = default_map.get("auth")

        if auth_cmd_config:
            for opt in ("force", "info"):
                if opt in auth_cmd_config:
                    errmsg = f"'auth' command line option '{opt}' cannot be set in the configuration file."
                    raise typer.BadParameter(errmsg)

    def _validate_export_opts(default_map: dict[str, Any], config_file_opts: dict[str, Any]) -> None:
        export_cmd_config = default_map.get("export")

        if export_cmd_config:
            config_file_export_opts = {}

            for opt in ("filter_test",):
                if opt in export_cmd_config:
                    errmsg = f"'export' command line option '{opt}' cannot be set in the configuration file."
                    raise typer.BadParameter(errmsg)

            # Store some options separately so they are not overwritten by command line options.
            for opt_var, expected_type in (
                ("tag", list),  # --tag
                ("body_tag", list),  # --body-tag
                ("metaprop", dict),  # --metaprop
                ("filter", list),  # --filter
                ("search", list),  # --search
            ):
                if opt_var in export_cmd_config:
                    config_value = export_cmd_config.get(opt_var)
                    if config_value is not None:
                        if not isinstance(config_value, expected_type):
                            errmsg = (
                                f"Configuration file error: Expected type '{expected_type.__name__}' "
                                f"for '{opt_var}', got '{type(config_value).__name__}'"
                            )
                            raise typer.BadParameter(errmsg)

                        config_file_export_opts[opt_var] = config_value
                        del export_cmd_config[opt_var]

            if config_file_export_opts:
                config_file_opts["export"] = config_file_export_opts

    if not config_file_or_dir:
        config_file_path = Path.cwd() / CONFIG_FILE_NAME
    elif config_file_or_dir.is_dir():
        config_file_path = config_file_or_dir / CONFIG_FILE_NAME
    elif config_file_or_dir.is_file():
        config_file_path = config_file_or_dir.absolute()
    else:
        msg = f"{config_file_or_dir}: No such file."
        raise typer.BadParameter(msg)

    if config_file_path.is_file():
        default_map = None

        try:
            default_map = _read_config_file(config_file_path)
        except typer.BadParameter as err:
            msg = f"Failed to read config file {config_file_path}: {err}"
            raise typer.BadParameter(msg) from err

        if default_map:
            ctx.ensure_object(CommandContext)
            ctx_obj: CommandContext = ctx.obj

            config_file_opts: dict[str, Any] = {}

            _validate_main_opts(default_map)
            _validate_auth_opts(default_map)
            _validate_export_opts(default_map, config_file_opts)

            ctx_obj.config_file_path = config_file_path
            ctx_obj.config_file_opts = config_file_opts or None

            if default_map:
                if ctx.default_map is None:
                    ctx.default_map = {}
                ctx.default_map.update(default_map)

    # Return value is ignored; see typer.Option(expose_value=False)
    return None


######################################################################
# Main

app = typer.Typer(name=APP_NAME, no_args_is_help=True, add_completion=False)


@app.callback()
def main(  # noqa: PLR0913, PLR0917
    ctx: typer.Context,
    config_file_or_dir: Annotated[  # noqa: ARG001
        Path | None,
        typer.Option(
            "--config-file",
            "-c",
            exists=False,
            dir_okay=True,  # Use default file name
            resolve_path=False,
            callback=_opt_config_file_callback,
            is_eager=True,
            expose_value=False,  # do not include as an argument to main()
            show_default=CONFIG_FILE_NAME,
            help="Read options from configuration FILE (can also be a DIRECTORY).",
            metavar="FILE",
        ),
    ] = None,
    auth_file: Annotated[
        Path | None,
        typer.Option(
            exists=True,
            dir_okay=True,
            file_okay=True,
            readable=True,
            writable=True,
            help="Path to the authentication file.",
            show_default=TOKEN_CACHE_FILE_NAME,
            metavar="FILE",
        ),
    ] = None,
    cache_dir: Annotated[
        Path | None,
        typer.Option(
            exists=True,
            dir_okay=True,
            file_okay=False,
            writable=True,
            help="Write cached files to DIRECTORY.",
            show_default="current working directory",
            metavar="DIRECTORY",
        ),
    ] = None,
    cache_graph_api: Annotated[
        bool | None,
        typer.Option(
            help="Cache Graph API responses and use cache when available. Implied if --cache-dir is specified.",
        ),
    ] = None,
    verbose: Annotated[
        int,
        typer.Option(
            "--verbose",
            "-v",
            count=True,
            metavar="",
            help="Increase verbose level.",
            show_default="minimal",
        ),
    ] = VERBOSE_LOGGING_LEVEL_ERROR,
    version: Annotated[  # noqa: ARG001
        bool | None,
        typer.Option(
            "--version",
            callback=_opt_version_callback,
            is_eager=True,
            help="Show the version and exit.",
        ),
    ] = None,
) -> None:
    """Download messages from Microsoft Outlook using the Microsoft Graph API and save them to an Obsidian vault."""
    ctx.ensure_object(CommandContext)
    ctx_obj: CommandContext = ctx.obj

    # Adjust logging level
    loglevel = VERBOSE_LOGGING_LEVELS[min(verbose, len(VERBOSE_LOGGING_LEVELS) - 1)]
    logger.setLevel(loglevel)

    if logger.getEffectiveLevel() == logging.DEBUG:
        # 'requests' relies on urllib3 under the hood.
        # Setting the urllib3 logger to DEBUG will capture connection attempts,
        # pool activity, and basic request URLs.
        urllib3_logger = logging.getLogger("urllib3")
        urllib3_logger.setLevel(logging.DEBUG)

    config_file_path = ctx_obj.config_file_path
    cache_dir_abs = cache_dir.absolute() if cache_dir else Path.cwd()
    graph_cache_enabled = (cache_graph_api is True) or (cache_graph_api is None and bool(cache_dir))

    if not auth_file:
        auth_file_path = Path.cwd() / TOKEN_CACHE_FILE_NAME
        for d in (config_file_path.parent if config_file_path else None, Path.cwd(), cache_dir_abs):
            if d and d.is_dir():
                auth_file_path = d / TOKEN_CACHE_FILE_NAME
                if auth_file_path.is_file():
                    break
    elif auth_file.is_dir():
        auth_file_path = auth_file / TOKEN_CACHE_FILE_NAME
    else:
        auth_file_path = auth_file.absolute()

    logger.debug("Working directory: '%s'", Path.cwd())
    logger.debug("Configuration file: %s", f"'{config_file_path}'" if config_file_path else "None")
    logger.debug("Auth token file: '%s'", auth_file_path)
    logger.debug("Cache directory: '%s'", cache_dir_abs)
    logger.debug("Graph API cache enabled: %s", "Yes" if graph_cache_enabled else "No")

    # Update common command options
    ctx_obj.common_opts = copy.replace(
        ctx_obj.common_opts,
        auth_file_path=auth_file_path,
        cache_dir=cache_dir_abs,
        graph_cache_enabled=graph_cache_enabled,
        verbose=verbose,
    )


app.add_typer(auth_app)
app.add_typer(export_app)
app.add_typer(ls_app)

if __name__ == "__main__":
    # Invoked by VSCode debugger
    app()
