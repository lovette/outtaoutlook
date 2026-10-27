"""outtaoutlook export command."""

import re
from decimal import Decimal
from enum import StrEnum
from pathlib import Path
from types import UnionType
from typing import Annotated, Any, get_args

import requests
import typer
import yaml
from requests import sessions

from outtaoutlook.graphapi import (
    AUTH_ERROR_USECOMMAND,
    GRAPH_API_BATCH_MAX,
    GRAPH_API_BATCH_MIN,
    MAILBOX_ROOT,
    GraphAuthRequiredError,
    get_graph_config,
    graph_resolve_folder_paths,
    print_graph_request_exception,
)
from outtaoutlook.shared import (
    CommandContext,
    ExportCancelledError,
    ExportError,
    FileMetaPropT,
    console,
    err_console,
    logger,
    path_with_vars,
    safe_file_name,
)

from .shared import (
    ATTACHMENTS_DEFAULT_FOLDER,
    ATTACHMENTS_FOLDER_FILE,
    ATTACHMENTS_FOLDER_ROOT,
    MAX_EXPORT_WORKERS,
    ExportOptions,
    OutputFormatChoice,
)
from .worker import export_folders

######################################################################
# Global variables

app = typer.Typer(add_completion=False)


######################################################################
# Constants


# Simple detection if filter/search critera contain 'and' or 'or' conditions.
RE_AND_OR = re.compile(r"\s(and|or)\s", re.IGNORECASE)


######################################################################
# Configure YAML to understand Decimal type


def _yaml_represent_decimal(dumper: yaml.dumper.SafeDumper, data: Decimal) -> yaml.ScalarNode:
    return dumper.represent_scalar("tag:yaml.org,2002:float", str(data))


yaml.SafeDumper.add_representer(Decimal, _yaml_represent_decimal)


######################################################################
# Types and classes


class FilterJoinCond(StrEnum):
    """Choices for joining filter conditions."""

    cond_and = "and"
    cond_or = "or"


######################################################################
# Helpers


def _opt_attachments_folder_callback(attachments_folder: str) -> str:
    """Handle `--attachments-folder` command option.

    Args:
        attachments_folder (str): Command line option value.

    Returns:
        str: Special keyword, valid folder name or absolute path.
    """
    # Special keywords.
    if attachments_folder.upper() == ATTACHMENTS_FOLDER_ROOT or attachments_folder.startswith(ATTACHMENTS_FOLDER_FILE):
        return attachments_folder

    # Name of subfolder such as "assets" or "attachments".
    if len(Path(attachments_folder).parts) == 1:
        if attachments_folder != safe_file_name(attachments_folder, ""):
            errmsg = f"'{attachments_folder}' is not a valid attachments directory name."
            raise typer.BadParameter(errmsg)

        return attachments_folder

    # Absolute or relative path to a directory optionally with ~ or $ENV variables.
    attachments_dir_path = path_with_vars(attachments_folder)
    if attachments_dir_path.is_dir():
        return str(attachments_dir_path)

    errmsg = f"Directory '{attachments_folder}' does not exist."
    raise typer.BadParameter(errmsg)


def _opt_root_out_dir_callback(ctx: typer.Context, root_out_dir: str) -> str:
    """Handle `--outdir` command option.

    Args:
        ctx (typer.Context): Typer context object.
        root_out_dir (str): Command line option value.

    Returns:
        str: Special keyword, valid folder name or absolute path.
    """
    out_format = ctx.params.get("out_format")

    if not root_out_dir:
        errmsg = "A path to an existing directory is required."
        raise typer.BadParameter(errmsg)

    # Absolute or relative path to a directory optionally with ~ or $ENV variables.
    root_out_dir_path = path_with_vars(root_out_dir)

    if out_format == OutputFormatChoice.maildir:
        if root_out_dir_path.exists():
            # Validate Maildir subdirectory structure.
            for subdir in ("cur", "new", "tmp"):
                if not (root_out_dir_path / subdir).is_dir():
                    errmsg = f"Directory '{root_out_dir}' exists but is not a Maildir: '/{subdir}' not found."
                    raise typer.BadParameter(errmsg)

    elif not root_out_dir_path.is_dir():
        errmsg = f"Directory '{root_out_dir}' does not exist."
        raise typer.BadParameter(errmsg)

    return str(root_out_dir_path)


def _opt_metaprop_callback(
    ctx: typer.Context,
    param: typer.CallbackParam,  # noqa: ARG001
    meta: list[str] | None,
) -> list[str] | None:
    """Handle `--metaprop` command option.

    Note:
        The values of `default_map` are used as the default values for command options.
        This function is called even if option is not specified.
        Duplicate keys are merged into a list of values

    Args:
        ctx (typer.Context): Typer context object.
        param (typer.CallbackParam): Typer callback parameter object.
        meta (list[str] | None): Command line option value or None if option is not specified.

    Returns:
        list[str] | None: List of parsed key-value pairs or None if no properties were provided.
    """
    if meta:
        ctx.ensure_object(CommandContext)
        ctx_obj: CommandContext = ctx.obj
        ctx_obj.export_metaprops = {}

        for value in meta:
            try:
                k, v = value.split("=", 1)
            except ValueError as err:
                msg = f"Format must be 'name=value', got: '{value}'"
                raise typer.BadParameter(msg) from err
            else:
                v_k = ctx_obj.export_metaprops.get(k)
                if v_k is not None:
                    # Duplicate keys are merged into a list of values
                    v_list = [v_k] if isinstance(v_k, (str, int)) else [*v_k]
                    v = [*v_list, v]

                ctx_obj.export_metaprops[k] = v

    return meta


def _merge_metaprops(a: FileMetaPropT, b: FileMetaPropT | None) -> FileMetaPropT:
    """Combines two dictionaries, merging duplicate keys into a list of values.

    Args:
        a (FileMetaPropT): Dictionary.
        b (FileMetaPropT | None): Dictionary to merge into `a`. If None, a copy of `a` is returned.

    Returns:
        FileMetaPropT
    """
    c = a.copy()

    if b:
        common_keys = a.keys() & b.keys()
        unique_to_b = b.keys() - a.keys()

        # Add b to a
        for key in unique_to_b:
            c[key] = b[key]

        # Merge duplicate keys into a list of values
        for key in common_keys:
            v_c, v_b = c[key], b[key]
            v_a_list = [v_c] if isinstance(v_c, (str, int)) else [*v_c]
            v_b_list = [v_b] if isinstance(v_b, (str, int)) else [*v_b]
            v_a_list.extend(v_b_list)
            c[key] = v_a_list

    return c


def _build_filter_query(
    message_filters: list[str] | None,
    filter_cond: FilterJoinCond,
    message_search: list[str] | None,
    search_cond: FilterJoinCond,
    *,
    filter_test: bool,
) -> tuple[str | None, str | None]:
    """Build Graph API filter and search queries based on critera.

    Args:
        message_filters (list[str] | None): List of filter criteria for messages.
        filter_cond (FilterJoinCond): Condition to join multiple filter criteria (AND/OR).
        message_search (list[str] | None): List of search criteria for messages.
        search_cond (FilterJoinCond): Condition to join multiple search criteria (AND/OR).
        filter_test (bool): Whether to perform a filter test without exporting.

    Returns:
        tuple[str | None, str | None]: (filter_query, search_query)
    """

    def _quote_search(s: str) -> str:
        return s if (s.startswith('"') and s.endswith('"')) else f'"{s.replace('"', '\\"')}"'

    def _quote_paren(s: str) -> str:
        return f"({s})" if RE_AND_OR.search(s) else s

    filter_query, search_query = None, None

    if message_filters:
        filter_query = (
            message_filters[0]
            if len(message_filters) == 1
            else f" {filter_cond.value.lower()} ".join(_quote_paren(f) for f in message_filters)
        )

    if message_search:
        search_query = _quote_search(
            message_search[0]
            if len(message_search) == 1
            else f" {search_cond.value.upper()} ".join(_quote_paren(s) for s in message_search),
        )

    if filter_query and search_query:
        msg = "'filter' and 'search' criteria cannot be used together."
        raise typer.BadParameter(msg)

    if filter_test and not (filter_query or search_query):
        msg = "'filter' or 'search' criteria must be provided."
        raise typer.BadParameter(msg, param_hint="--filter-test")

    return filter_query, search_query


def _validate_config_file_collection(
    config_opt: str,
    collection: list[Any] | dict[str, Any],
    collection_type: type,
    item_type: type | UnionType,
) -> None:
    """Validate types for a collection and its values.

    Args:
        config_opt (str): Configuration file option.
        collection (list[Any] | dict[str, Any]): Collection to validate.
        collection_type (type): Expected type of the collection (list or dict).
        item_type (type | UnionType): Expected type(s) of the items in the collection.

    Raises:
        typer.BadParameter: Raised if the collection or any values are not of the expected type.
    """
    if collection is not None:
        if not isinstance(collection, collection_type):
            errmsg = (
                f"Configuration file error: Expected '{config_opt}' to be a {collection_type.__name__}, "
                f"got a {type(collection).__name__} instead"
            )
            raise typer.BadParameter(errmsg)

        items = collection if isinstance(collection, list) else collection.values()

        if isinstance(item_type, UnionType):
            item_types = get_args(item_type)  # str | int | bool
            item_types_str = ", ".join(t.__name__ for t in item_types)  # "str, int, bool"

            for item in items:
                if not any(isinstance(item, t) for t in item_types):
                    errmsg = (
                        f"Configuration file error: '{config_opt}' must be a {collection_type.__name__} "
                        f"of types {item_types_str}; found a '{type(item).__name__}' instead"
                    )
                    raise typer.BadParameter(errmsg)
        else:
            for item in items:
                if not isinstance(item, item_type):
                    errmsg = (
                        f"Configuration file error: '{config_opt}' must be a {collection_type.__name__} "
                        f"of type {item_type.__name__}; found a '{type(item).__name__}' instead"
                    )
                    raise typer.BadParameter(errmsg)


######################################################################
# Commands


@app.command()
def export(  # noqa: C901, PLR0912, PLR0913, PLR0915, PLR0917
    ctx: typer.Context,
    out_dir: Annotated[
        Path | None,
        typer.Option(
            "--outdir",
            "-o",
            exists=False,  # _opt_root_out_dir_callback will confirm
            dir_okay=True,
            file_okay=False,
            writable=True,
            metavar="DIRECTORY",
            help="Write output files to this directory. [required]",
            callback=_opt_root_out_dir_callback,
            rich_help_panel="Output",
        ),
    ] = None,
    out_format: Annotated[
        OutputFormatChoice,
        typer.Option(
            "--format",
            "-F",
            help="Output file format.",
            rich_help_panel="Output",
            is_eager=True,  # we need to reference in `_opt_root_out_dir_callback`
        ),
    ] = OutputFormatChoice.md,
    folder: Annotated[
        list[str] | None,
        typer.Option(
            "--folder",
            "-f",
            metavar="FOLDER",
            help="Export messages in this mail folder; option can be used more than once. [required unless using --mailbox]",
            rich_help_panel="Mail folders",
        ),
    ] = None,
    allfolders: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--mailbox",
            "-a",
            help="Export messages in all top-level mail folders (and subfolders up to --maxdepth.)",
            rich_help_panel="Mail folders",
        ),
    ] = False,
    tag: Annotated[
        list[str] | None,
        typer.Option(
            "--tag",
            "-t",
            metavar="TAG",
            help="Tag to apply to files (placed either in frontmatter or body); option can be used more than once.",
            rich_help_panel="Tags and metadata",
        ),
    ] = None,
    body_tag: Annotated[
        list[str] | None,
        typer.Option(
            "--body-tag",
            "-b",
            metavar="TAG",
            help="Tag to place in files body; option can be used more than once.",
            rich_help_panel="Tags and metadata",
        ),
    ] = None,
    metaprop: Annotated[  # noqa: ARG001
        list[str] | None,
        typer.Option(
            "--metaprop",
            "-p",
            callback=_opt_metaprop_callback,
            metavar="NAME=VALUE",
            help="Frontmatter metadata property to apply to files; option can be used more than once.",
            rich_help_panel="Tags and metadata",
        ),
    ] = None,
    message_filter: Annotated[
        list[str] | None,
        typer.Option(
            "--filter",
            metavar="CONDITION",
            help="Export messages matching filter criteria; option can be used more than once.",
            rich_help_panel="Filter and search",
        ),
    ] = None,
    filter_cond: Annotated[
        FilterJoinCond,
        typer.Option(
            "--filter-cond",
            help="How to join multiple filter conditions (and/or).",
            rich_help_panel="Filter and search",
        ),
    ] = FilterJoinCond.cond_and,
    filter_test: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--filter-test",
            help="Test filter and search criteria and exit.",
            rich_help_panel="Filter and search",
        ),
    ] = False,
    message_search: Annotated[
        list[str] | None,
        typer.Option(
            "--search",
            metavar="CONDITION",
            help="Export messages matching search criteria; option can be used more than once.",
            rich_help_panel="Filter and search",
        ),
    ] = None,
    search_cond: Annotated[
        FilterJoinCond,
        typer.Option(
            "--search-cond",
            help="How to join multiple search conditions (and/or).",
            rich_help_panel="Filter and search",
        ),
    ] = FilterJoinCond.cond_and,
    not_folder: Annotated[
        list[str] | None,
        typer.Option(
            "--not-folder",
            "-n",
            help="Exclude mail folder from export; option can be used more than once.",
            metavar="FOLDER",
            rich_help_panel="Mail folders",
        ),
    ] = None,
    md_place_tags_frontmatter: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--place-tags-frontmatter/--no-place-tags-frontmatter",
            help="Place --tag tags in file frontmatter.",
            rich_help_panel="Tags and metadata",
        ),
    ] = False,
    md_place_tags_body_end: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--place-tags-body-end/--no-place-tags-body-end",
            help="Place body tags at the end of the file instead of the top.",
            rich_help_panel="Tags and metadata",
        ),
    ] = False,
    flat: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--flat",
            help="Write all messages to the export directory; do not preserve mail folder hierarchy.",
            rich_help_panel="Mail folders",
        ),
    ] = False,
    shallow: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--shallow",
            rich_help_panel="Mail folders",
            help="Use only the leaf (last) folder when creating mail folder hierarchy.",
        ),
    ] = False,
    with_attachments: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--attachments/--no-attachments",
            help="Download and link attachments to file.",
            rich_help_panel="Attachments",
        ),
    ] = True,
    attachments_folder: Annotated[
        str,
        typer.Option(
            exists=False,
            dir_okay=True,
            file_okay=False,
            writable=True,
            metavar=f"'{ATTACHMENTS_FOLDER_ROOT}'|'{ATTACHMENTS_FOLDER_FILE}'|NAME|DIRECTORY",
            help="Store attachments in the output directory, same as files, a subfolder of files, or another directory.",
            rich_help_panel="Attachments",
            callback=_opt_attachments_folder_callback,
        ),
    ] = ATTACHMENTS_DEFAULT_FOLDER,
    attachments_mkdir: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            rich_help_panel="Attachments",
            help="Create mail folder hierarchy in attachments directory when outside the output directory.",
        ),
    ] = False,
    maxdepth: Annotated[
        int,
        typer.Option(
            "--maxdepth",
            min=1,
            rich_help_panel="Mail folders",
            help="Maximum depth of folder tree to export.",
        ),
    ] = 1,
    use_manifest: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--manifest/--no-manifest",
            rich_help_panel="Other options",
            help="Write a manifest of exported messages and use to skip duplicates.",
        ),
    ] = True,
    limit: Annotated[
        int | None,
        typer.Option(
            min=1,
            help="Number of messages to export.",
            rich_help_panel="Other options",
        ),
    ] = None,
    batch_size: Annotated[
        int,
        typer.Option(
            min=GRAPH_API_BATCH_MIN,
            max=GRAPH_API_BATCH_MAX,
            rich_help_panel="Other options",
            help="Number of messages to to fetch in each Graph API request.",
        ),
    ] = GRAPH_API_BATCH_MAX,
    max_workers: Annotated[
        int | None,
        typer.Option(
            min=1,
            max=MAX_EXPORT_WORKERS,
            rich_help_panel="Other options",
            help="Maximum size of worker pool.",
        ),
    ] = None,
    assume_yes: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--yes",
            rich_help_panel="Other options",
            help="Assume 'yes' as answer to all prompts and run non-interactively.",
        ),
    ] = False,
    prime_cache: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--prime-cache",
            help="Do not write any files, just prime the cache.",
            rich_help_panel="Other options",
        ),
    ] = False,
    dry_run: Annotated[  # noqa: FBT002
        bool,
        typer.Option(
            "--dry-run",
            rich_help_panel="Other options",
            help="Dry run; do not write any files.",
        ),
    ] = False,
) -> None:
    """Export email messages from Microsoft Outlook and save them locally in various file formats.

    For guidance on using filters, see:
    https://learn.microsoft.com/en-us/graph/filter-query-parameter

    For guidance on using search, see:
    https://learn.microsoft.com/en-us/graph/search-query-parameter
    """
    # Options common to all commands
    ctx_obj: CommandContext = ctx.obj

    assert ctx_obj is not None  # noqa: S101
    assert isinstance(ctx_obj, CommandContext)  # noqa: S101

    common_opts = ctx_obj.common_opts
    config_file_export_opts: dict[str, Any] = ctx_obj.config_file_opts.get("export", {}) if ctx_obj.config_file_opts else {}
    config_file_md_tags: list[str] = config_file_export_opts.get("tag", [])
    config_file_md_body_tags: list[str] = config_file_export_opts.get("body_tag", [])
    config_file_md_metaprop: dict[str, Any] = config_file_export_opts.get("metaprop", {})
    config_file_filters: list[str] = config_file_export_opts.get("filter", [])
    config_file_search: list[str] = config_file_export_opts.get("search", [])
    config_file_not_folder: list[str] = config_file_export_opts.get("not_folder", [])
    config_file_select_folder: list[str] = config_file_export_opts.get("folder", [])

    # Simple YAML syntax changes can result in values being interpreted as the wrong type!
    _validate_config_file_collection("tag", config_file_md_tags, list, str)
    _validate_config_file_collection("body_tag", config_file_md_body_tags, list, str)
    _validate_config_file_collection("metaprop", config_file_md_metaprop, dict, str | int | bool)
    _validate_config_file_collection("filter", config_file_filters, list, str)
    _validate_config_file_collection("search", config_file_search, list, str)
    _validate_config_file_collection("not_folder", config_file_not_folder, list, str)
    _validate_config_file_collection("folder", config_file_select_folder, list, str)

    md_tags = config_file_md_tags + (tag or [])
    md_body_tags = config_file_md_body_tags + (body_tag or [])
    md_metaprop = _merge_metaprops(config_file_md_metaprop, ctx_obj.export_metaprops)
    not_folder = config_file_not_folder + (not_folder or [])
    not_folder_set = {nf.lower() for nf in not_folder}
    folder = config_file_select_folder + (folder or [])

    filter_query, search_query = _build_filter_query(
        config_file_filters + (message_filter or []),
        filter_cond,
        config_file_search + (message_search or []),
        search_cond,
        filter_test=filter_test,
    )

    # Mbox is a single file so requires one worker per mbox.
    one_mbox = out_format == OutputFormatChoice.mbox and flat

    if not out_dir:
        errmsg = "Missing option '--outdir' / '-o'."
        raise typer.BadParameter(errmsg)
    if not (folder or allfolders):
        errmsg = "At least one '--folder' or the '--mailbox' option must be specified."
        raise typer.BadParameter(errmsg)
    if folder and allfolders:
        errmsg = "The '--folder' and '--mailbox' options cannot be used together."
        raise typer.BadParameter(errmsg)
    if one_mbox and (max_workers and max_workers > 1):
        errmsg = "When exporting all mail to a single mbox, --max-workers must be set to 1."
        raise typer.BadParameter(errmsg)
    if flat and shallow:
        errmsg = "The '--flat' and '--shallow' options cannot be used together."
        raise typer.BadParameter(errmsg)
    if prime_cache and not common_opts.graph_cache_enabled:
        errmsg = "The '--prime-cache' option requires the Graph API cache to be enabled."
        raise typer.BadParameter(errmsg)
    if prime_cache and common_opts.offline:
        errmsg = "The '--offline' and '--prime-cache' options cannot be used together."
        raise typer.BadParameter(errmsg)

    try:
        graph_config = get_graph_config(common_opts)

        logger.debug("Resolving folder paths")

        with sessions.Session() as session:
            folders = graph_resolve_folder_paths(
                session,
                graph_config,
                MAILBOX_ROOT if allfolders else folder,
                maxdepth=maxdepth,
            )

    except requests.RequestException as err:
        print_graph_request_exception(err_console, err)
        raise typer.Exit(code=1) from err
    except GraphAuthRequiredError as err:
        err_console.print(f"{err} {AUTH_ERROR_USECOMMAND}")
        raise typer.Exit(code=1) from err
    except ExportCancelledError as err:
        # Whichever thread signaled the cancellation already logged the reason...
        raise typer.Exit(code=1) from err
    except (ExportError, ValueError, OSError) as err:
        err_console.print(f"Error: {err}")
        raise typer.Exit(code=1) from err

    # Remove excluded folders
    folders = [f for f in folders if f["folderPath"].lower() not in not_folder_set]
    if not folders:
        err_console.print("No folders selected to export.")
        raise typer.Exit(code=1)

    # We don't need more workers than we have folders.
    if one_mbox:
        max_workers = 1
    elif max_workers:
        max_workers = min(max_workers, len(folders))
    elif len(folders) < MAX_EXPORT_WORKERS:
        max_workers = len(folders)

    export_opts = ExportOptions(
        out_format=out_format,
        attachments_folder=attachments_folder,
        attachments_mkdir=attachments_mkdir,
        batch_size=batch_size,
        dry_run=dry_run or prime_cache,
        filter_query=filter_query,
        filter_test=filter_test,
        manifest_dir=common_opts.cache_dir if use_manifest else None,
        hier_shallow=shallow,
        md_body_tags=md_body_tags or None,
        md_metaprop=md_metaprop or None,
        md_tags=md_tags or None,
        files_dir=out_dir,
        export_limit=limit or 0,
        hier_flat=flat,
        md_place_tags_body_end=md_place_tags_body_end,
        md_place_tags_frontmatter=md_place_tags_frontmatter,
        search_query=search_query,
        root_out_dir=out_dir,
        with_attachments=with_attachments,
        max_workers=max_workers or MAX_EXPORT_WORKERS,
        assume_yes=assume_yes,
        prime_cache=prime_cache,
    )

    if filter_test:
        if filter_query:
            console.print(f"[bold]Testing filter criteria for {len(folders)} folders")
        if search_query:
            console.print(f"[bold]Testing search criteria for {len(folders)} folders")
    else:
        msg = [f"Exporting messages from {len(folders)} folders to '{out_dir}'"]
        if export_opts.dry_run:
            msg.append("(dry run)")
        msg.append("...")

        console.print(" ".join(msg))

    if filter_query:
        console.print(f"Filter criteria: {filter_query}")
    else:
        logger.debug("Filter criteria: %s", filter_query)

    if search_query:
        console.print(f"Search criteria: {search_query}")
    else:
        logger.debug("Search criteria: %s", search_query)

    console.print()

    export_folders(folders, graph_config, export_opts, max_workers)
