"""outtaoutlook shared variables and functions."""

import base64
import binascii
import hashlib
import json
import logging
import os
import re
import textwrap
import threading
from collections import defaultdict
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from enum import StrEnum
from operator import itemgetter
from pathlib import Path
from typing import TYPE_CHECKING, Any

import msal
import requests
from requests import Response, sessions
from rich.console import Console

if TYPE_CHECKING:
    from rich.progress import Progress, TaskID

######################################################################
# Constants

APP_NAME = "outtaoutlook"
AUTHORITY_ENVIRONMENT_DEFAULT = "login.microsoftonline.com"
GRAPH_API_URL = "https://graph.microsoft.com/beta"
GRAPH_SCOPES = ["Mail.Read"]
GRAPH_API_BATCH_MIN = 25  # Arbitrary limit
GRAPH_API_BATCH_MAX = 100  # Arbitrary limit
MAILBOX_ROOT = None
VERBOSE_LOGGING_LEVEL_ERROR = 0  # logging.ERROR
TITLE_UNTITLED = "Untitled"
TITLE_MAXLEN = 255 - 25  # macOS limit minus a few characters for whatever `unique_file_path` tacks on
GRAPH_API_TIMEOUT = 30

# These are banned outright: \/:
# Links do not work if title includes: #^[]|
# macOS is fine with: *"<>?
# Title cannot start with a period.
# \x00-\x1f\x7f = [:cntrl:]
RE_TITLE_IGNORECHARS = re.compile(r"[\[\]\/\\:|^#\x00-\x1f\x7f]")
RE_TITLE_WHITESPACE = re.compile(r"\s+")

AUTH_ERROR_USECOMMAND = "Use the 'auth' command to sign in and grant access."
AUTH_REQUIRED_ERROR = "Access to Microsoft Outlook via the Microsoft Graph API is required."
AUTH_EXPIRED_ERROR = "Previously granted access to Microsoft Outlook has expired."
AUTH_FAILED_ERROR = "Access to Microsoft Outlook was not granted."
AUTH_NO_CACHE_ERROR = "Access to Microsoft Outlook has not been granted."

GRAPH_ACCESS_TOKEN_FIELDS = [
    "access_token",
    "client_info",
    "expires_in",
    "ext_expires_in",
    "id_token_claims",
    "id_token",
    "refresh_token",
    "scope",
    "token_source",
    "token_type",
]

GRAPH_CACHED_TOKEN_FIELDS = [
    "cached_at",
    "client_id",
    "credential_type",
    "environment",
    "expires_on",  # always seems to be the same as `cached_at` :/
    "extended_expires_on",
    "home_account_id",
    "realm",
    "secret",
    "target",
    "token_type",
]

GRAPH_ACCOUNT_FIELDS = [
    "account_source",
    "authority_type",
    "environment",
    "home_account_id",
    "local_account_id",
    "realm",
    "username",
]

GRAPH_APPMETADATA_FIELDS = [
    "client_id",
    "environment",
]

# Fields returned from '/mailFolders'
GRAPH_FOLDER_FIELDS = [
    "id",  # str
    "displayName",  # str
    "parentFolderId",  # str
    "childFolderCount",  # int
    "unreadItemCount",  # int
    "totalItemCount",  # int
    "sizeInBytes",  # int
    "wellKnownName",  # str or 'null'
    "isHidden",  # bool
]

# Fields returned from '/messages' and '/message'.
# Not all fields are guaranteed to be present in every message.
GRAPH_MESSAGE_FIELDS = [
    "@odata.etag",  # str
    "bccRecipients",  # [{'emailAddress': {'name': str, 'address': str}}, ...]
    "body",  #  { "contentType": "html", "content": str, },
    "bodyPreview",  # str
    "categories",  # []
    "ccRecipients",  # [{'emailAddress': {'name': str, 'address': str}}, ...]
    "changeKey",  # str
    "conversationId",  # "UUID"
    "conversationIndex",  # "UUID"
    "createdDateTime",  # "YYYY-MM-DDTHH:MM:SSZ"
    "flag",  # {"flagStatus": str},
    "from",  # {'emailAddress': {'name': str, 'address': str}}
    "hasAttachments",  # bool
    "id",  # str
    "importance",  # str
    "inferenceClassification",  # str
    "internetMessageId",  # str
    "isDeliveryReceiptRequested",  # bool
    "isDraft",  # bool
    "isRead",  # bool
    "isReadReceiptRequested",  # bool
    "lastModifiedDateTime",  # "YYYY-MM-DDTHH:MM:SSZ"
    "mentionsPreview",  # null
    "parentFolderId",  # "UUID"
    "receivedDateTime",  # "YYYY-MM-DDTHH:MM:SSZ"
    "replyTo",  # [{'emailAddress': {'name': str, 'address': str}}, ...]
    "sender",  # {'emailAddress': {'name': str, 'address': str}}
    "sentDateTime",  # "YYYY-MM-DDTHH:MM:SSZ"
    "subject",  # str
    "toRecipients",  # [{'emailAddress': {'name': str, 'address': str}}, ...]
    "unsubscribeData",  # []
    "unsubscribeEnabled",  # bool
    "webLink",  # "URL"
]

GRAPH_ATTACHMENT_FIELDS = [
    "@odata.mediaContentType",  # str; e.g. 'image/jpeg'
    "@odata.type",  # #microsoft.graph.[itemAttachment|fileAttachment]
    "contentBytes",  # str; if isInline is True
    "contentId",  # str
    "contentLocation",  # str|None
    "contentType",  # str|None; e.g. 'image/jpeg'
    "id",  # UUID
    "isInline",  # bool
    "lastModifiedDateTime",  # isodate; unreliable, how does a message dated 2009 have a modified date of 2025?!
    "name",  # str
    "size",  # int
]

GRAPH_ACCESS_TOKEN_FIELDS_SET = set(GRAPH_ACCESS_TOKEN_FIELDS)
GRAPH_CACHED_TOKEN_FIELDS_SET = set(GRAPH_CACHED_TOKEN_FIELDS)
GRAPH_ACCOUNT_FIELDS_SET = set(GRAPH_ACCOUNT_FIELDS)
GRAPH_FOLDER_FIELDS_SET = set(GRAPH_FOLDER_FIELDS)
GRAPH_MESSAGE_FIELDS_SET = set(GRAPH_MESSAGE_FIELDS)
GRAPH_APPMETADATA_FIELDS_SET = set(GRAPH_APPMETADATA_FIELDS)
GRAPH_ATTACHMENT_FIELDS_SET = set(GRAPH_ATTACHMENT_FIELDS)

# wellKnownName of folders that make no sense to export
INTERNAL_FOLDERS = {
    "conversationhistory",
    "deleteditems",
    "outbox",
    "syncissues",
}

######################################################################
# Types and classes


GraphApiReturnT = dict[str, Any]  # The return type of Graph API calls
GraphAccessTokenT = dict[str, Any]  # GRAPH_ACCESS_TOKEN_FIELDS
GraphCachedTokenT = dict[str, Any]  # GRAPH_TOKEN_FIELDS
GraphAccountT = dict[str, Any]  # GRAPH_ACCOUNT_FIELDS
GraphFolderT = dict[str, Any]  # GRAPH_FOLDER_FIELDS
GraphMessageT = dict[str, Any]  # GRAPH_MESSAGE_FIELDS
GraphAppMetadataT = dict[str, Any]  # GRAPH_APPMETADATA_FIELDS
GraphAttachmentT = dict[str, Any]  # GRAPH_ATTACHMENT_FIELDS_SET
FileMetaPropT = dict[str, str | int | Sequence[str | int]]


class GraphApiError(Exception):
    """Base class for Graph API related errors."""


class GraphEmptyRootFolderError(GraphApiError):
    """Raised if the Graph API returns no properties for the root information store."""


class GraphFolderNotFoundError(GraphApiError):
    """Raised when a folder path cannot be found in the Graph API."""


class GraphAuthRequiredError(GraphApiError):
    """Raised when user authentication is required."""


@dataclass
class CachedAccessTokenInfo:
    """Relevant details from SerializableTokenCache."""

    environment: str
    tenant: str
    client: str
    username: str
    cached_at: datetime


@dataclass
class GraphApiConfig:
    """Configuration for HTTP caching."""

    access_token: str
    cache_dir: Path | None
    cache_enabled: bool
    token_info: CachedAccessTokenInfo


@dataclass
class CommonCommandOptions:
    """Common command options."""

    auth_file_path: Path | None = None
    cache_dir: Path | None = None
    graph_cache_enabled: bool = False
    verbose: int = VERBOSE_LOGGING_LEVEL_ERROR


@dataclass
class CommandContext:
    """Additional context for command execution accessed through `ctx.obj`."""

    common_opts: CommonCommandOptions = field(init=False)

    config_file_path: Path | None = None  # set if --config-file is provided and is read
    config_file_opts: dict[str, Any] | None = None  # set if configuration file contains non-`default_map` options

    export_metaprops: FileMetaPropT | None = None  # set if export command has `--metaprop` options

    def __post_init__(self) -> None:  # noqa: D105
        self.common_opts = CommonCommandOptions()


class OutputFormatChoice(StrEnum):
    """Choices for output file format."""

    eml = "eml"
    html = "html"
    json = "json"
    md = "md"


class ExportError(GraphApiError):
    """Export failure."""


class ExportStat(StrEnum):
    """Export statistic messages."""

    STAT_FILE_CREATED = "Files created"
    STAT_ATTACHMENT_SAVED = "Attachments saved"
    STAT_SKIPPED_DUPLICATE = "Skipped; Message already exported"
    STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED = "Skipped; Markdown conversion failed"
    STAT_SKIPPED_NO_MESSAGE_BODY = "Skipped; No message body"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS = "Skipped; No file content, just attachments (not being saved)"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS = "Skipped; No file content"
    STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN = "Skipped; No file content, unknown reason"
    STAT_SKIPPED_EVENT_MESSAGE = "Skipped; Calendar event"


@dataclass
class ExportOptions:
    """Export command options."""

    out_format: OutputFormatChoice  # Output file format.
    attachments_folder: str | Path  # `ATTACHMENTS_FOLDER_*`, folder name or absolute path.
    attachments_mkdir: bool  # Create mail folder hierarchy in attachments directory.
    batch_size: int  # Number of messages to process in a batch.
    dry_run: bool  # Whether to perform a dry run without making any changes.
    filter_query: str | None  # --filter criteria from command line and config file.
    filter_test: bool  # Whether to perform a filter test without exporting.
    manifest_dir: Path | None  # Directory to write the manifest file.
    mkdir_maxdepth: int | None  # Maximum depth (from the leaf) of folder tree to create when creating mail folder hierarchy.
    md_body_tags: list[str] | None  # --body-tag tags from command line and config file.
    md_metaprop: FileMetaPropT | None  # --metaprop properties from command line and config file.
    md_tags: list[str] | None  # --tag tags from command line and config file.
    files_dir: Path  # Directory to save files (set internally.)
    files_limit: int  # Maximum number of files to export (per folder).
    files_mkdir: bool  # Create mail folder hierarchy in files directory.
    md_place_tags_body_end: bool  # True to place body tags at the end of the file instead the top.
    md_place_tags_frontmatter: bool  # True to place non-body tags in frontmatter.
    search_query: str | None  # --search criteria from command line and config file.
    root_out_dir: Path  # Root output directory (set by user.)
    with_attachments: bool  # Whether to download and link attachments to files.
    max_workers: int  # Maximum number of workers being used for exporting.


@dataclass
class ExportFolderContext:
    """Context for exporting a folder."""

    folder: GraphFolderT
    session: sessions.Session
    graph_config: GraphApiConfig
    progress: Progress
    progress_task_id: TaskID
    export_opts: ExportOptions
    export_stats: ExportStatsT


ExportStatsT = dict[ExportStat, int]


######################################################################
# Globals

console = Console(soft_wrap=True)  # Enable soft wrap so newlines are not inserted to wrap lines wider than the terminal
err_console = Console(stderr=True, soft_wrap=True, style="red", highlight=False)
logger = logging.getLogger(APP_NAME)  # initialized in __main__


# Graph API limits concurrent requests.
# Requests will fail with 'MailboxConcurrency' error if limit is exceeded.
# https://learn.microsoft.com/en-us/graph/throttling-limits
GRAPH_API_CONCURRENT_LIMIT = 4
graph_get_resource_lock = threading.Semaphore(GRAPH_API_CONCURRENT_LIMIT)

# There is a gap of time between a unique path name being generated
# and the file is actually created on disk.
# Therefore we need a way to reserve unique paths that is coordinated between all threads.
# We do this by keeping a set of unique paths and a lock to coordinate access.
# dict[path.parent]=(set(path.name), lock)
unique_file_path_sets: dict[Path, tuple[set[str], threading.Lock]] = defaultdict(lambda: (set(), threading.Lock()))


######################################################################
# Helpers


def datetime_now() -> datetime:
    """Return current datetime in UTC.

    Returns:
        datetime
    """
    return datetime.now(tz=timezone.utc)  # noqa: UP017


def get_token_cache_info(auth_file_path: Path) -> CachedAccessTokenInfo | None:
    """Retrieve information from the token cache if it exists.

    Returns:
        TokenCacheInfo | None: Information about the cached token, or None if no valid cache is found.
    """
    if not auth_file_path.is_file():
        logger.debug("Access token is not cached (looking for '%s')", auth_file_path)
        return None

    cache_text = auth_file_path.read_text()
    cache_json = json.loads(cache_text)

    # We're interested in the first key/value pair
    token_uuid = next(iter(cache_json.get("AccessToken", {})), None)
    if not token_uuid:
        logger.warning("Access token cache exists but is missing expected field 'AccessToken'.")
        return None

    # We're interested in the first key/value pair
    account_uuid = next(iter(cache_json.get("Account", {})), None)
    if not account_uuid:
        logger.warning("Access token cache exists but is missing expected field 'Account'.")
        return None

    # We're interested in the first key/value pair
    metadata_uuid = next(iter(cache_json.get("AppMetadata", {})), None)
    if not metadata_uuid:
        logger.warning("Access token cache exists but is missing expected field 'AppMetadata'.")
        return None

    access_token: GraphCachedTokenT = cache_json["AccessToken"][token_uuid]
    account_info: GraphAccountT = cache_json["Account"][account_uuid]
    client_info: GraphAppMetadataT = cache_json["AppMetadata"][metadata_uuid]

    # Sanity check GraphCachedTokenT fields
    missing_fields = GRAPH_CACHED_TOKEN_FIELDS_SET - set(access_token.keys())
    if missing_fields:
        logger.debug(
            "Graph API cached access token is missing fields: '%s'",
            "', '".join(sorted(missing_fields)),
        )

    # Sanity check GraphAccountT fields
    missing_fields = GRAPH_ACCOUNT_FIELDS_SET - set(account_info.keys())
    if missing_fields:
        logger.debug(
            "Graph API cached account is missing fields: '%s'",
            "', '".join(sorted(missing_fields)),
        )

    # Sanity check GraphAppMetadataT fields
    missing_fields = GRAPH_APPMETADATA_FIELDS_SET - set(client_info.keys())
    if missing_fields:
        logger.debug(
            "Graph API cached app metadata is missing fields: '%s'",
            "', '".join(sorted(missing_fields)),
        )

    return CachedAccessTokenInfo(
        environment=account_info["environment"],
        client=client_info["client_id"],
        tenant=account_info["home_account_id"].split(".")[1],
        username=account_info["username"],
        cached_at=datetime.fromtimestamp(int(access_token["cached_at"]), tz=timezone.utc),  # noqa: UP017
    )


def get_graph_config(common_opts: CommonCommandOptions) -> GraphApiConfig:
    """Get cached Graph API configuration details including access token from the MSAL cache.

    Returns:
        GraphApiConfig: Graph API configuration.

    Raises:
        GraphAuthRequiredError: Raised if access token is missing or invalid.
        GraphApiError: MSAL-related error, such as an invalid MSAL environment or invalid token cache.
    """
    if not common_opts.auth_file_path:
        errmsg = "Internal error: Authentication file path is not set."
        raise GraphApiError(errmsg)

    # First see if there is cached authentication information.
    token_info = get_token_cache_info(common_opts.auth_file_path)
    if not token_info:
        raise GraphAuthRequiredError(AUTH_REQUIRED_ERROR)

    # Try to acquire a valid access token using cached credentials.
    _, _, _, token = get_graph_client(common_opts.auth_file_path, use_cache=True)
    if not token or "access_token" not in token:
        raise GraphAuthRequiredError(AUTH_EXPIRED_ERROR)

    return GraphApiConfig(
        access_token=token["access_token"],
        cache_enabled=common_opts.graph_cache_enabled,
        cache_dir=common_opts.cache_dir,
        token_info=token_info,
    )


def get_graph_client(  # noqa: C901, PLR0912
    auth_file_path: Path,
    environment: str | None = None,
    tenant: str | None = None,
    client: str | None = None,
    *,
    use_cache: bool = True,
) -> tuple[
    msal.PublicClientApplication,
    msal.SerializableTokenCache,
    GraphAccountT | None,
    GraphAccessTokenT | None,
]:
    """Create client application and token cache along with authenticated user and token if possible.

    Account and token are returned only if token cache exists and is being used.
    Token cache is not updated.

    Args:
        auth_file_path (Path): Path to the token cache file.
        environment (str, optional): Typically the domain of the Azure AD instance.
            If not provided, it will be inferred from the cached token if available, otherwise 'login.microsoftonline.com'.
        tenant (str, optional): Tenant ID. If not provided, it will be inferred from the cached token if available.
        client (str, optional): Client ID. If not provided, it will be inferred from the cached token if available.
        use_cache (bool, optional): Whether to use cached tokens. Defaults to True.

    Returns:
        tuple[msal.PublicClientApplication, msal.SerializableTokenCache, GraphAccountT | None, GraphAccessTokenT | None]

    Raises:
        GraphAuthRequiredError: Cache is being used and access token is unavailable or invalid.
        GraphApiError: MSAL-related error, such as an invalid MSAL environment or invalid token cache.
    """
    account = None
    token = None
    cache = msal.SerializableTokenCache()
    cache_info = get_token_cache_info(auth_file_path)

    if use_cache and cache_info is not None:
        cache_text = auth_file_path.read_text()
        cache.deserialize(cache_text)

        if not environment:
            environment = cache_info.environment
        if not tenant:
            tenant = cache_info.tenant
        if not client:
            client = cache_info.client

    elif auth_file_path.is_file():
        # Reset cache
        auth_file_path.unlink()
        cache_info = None

    if not environment:
        environment = AUTHORITY_ENVIRONMENT_DEFAULT

    if not (client and tenant):
        if cache_info:
            raise GraphAuthRequiredError(AUTH_EXPIRED_ERROR)
        raise GraphAuthRequiredError(AUTH_REQUIRED_ERROR)

    app = msal.PublicClientApplication(
        client_id=client,
        authority=f"https://{environment}/{tenant}",
        token_cache=cache,
    )

    if cache_info:
        # List previously authenticated user accounts.
        accounts = app.get_accounts()

        if accounts:
            # Look for a valid token or use a refresh token automatically.
            account = accounts[0]
            token = app.acquire_token_silent(scopes=GRAPH_SCOPES, account=account)

            # Sanity check GraphAccountT fields
            missing_fields = GRAPH_ACCOUNT_FIELDS_SET - set(account.keys())
            if missing_fields:
                logger.debug(
                    "Graph API account object is missing fields: '%s'",
                    "', '".join(sorted(missing_fields)),
                )

        if not token:
            raise GraphAuthRequiredError(AUTH_EXPIRED_ERROR)

        if "access_token" not in token:
            errmsg = "Graph API returned an access token that does not have an 'access_token' property."
            raise GraphApiError(errmsg)

        # Sanity check GraphAccessTokenT fields
        missing_fields = GRAPH_ACCESS_TOKEN_FIELDS_SET - set(token.keys())
        if missing_fields:
            logger.debug(
                "Graph API acquired access token is missing fields: '%s'",
                "', '".join(sorted(missing_fields)),
            )

    return app, cache, account, token


def graph_get_resource(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    url: str,
    params: dict[str, Any] | None = None,
    headers: dict[str, Any] | None = None,
) -> GraphApiReturnT:
    """Execute a GET request to Microsoft Graph API.

    Args:
        session (sessions.Session): Requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        url (str): Entity URL.
        params (dict, optional): Query parameters for the GET request. Defaults to None.
        headers (dict, optional): HTTP headers for the GET request. Defaults to None.

    Returns:
        GraphApiReturnT: JSON response from the Microsoft Graph API.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    req_headers = {
        "Authorization": f"Bearer {graph_config.access_token}",
        "Content-Type": "application/json",
    }

    if headers:
        req_headers.update(headers)

    req = requests.Request("GET", url, headers=req_headers, params=params).prepare()
    assert req.url is not None  # noqa: S101

    cache_file_path, cache_json_header = _get_cache_file_path_and_header(graph_config, "GET", url, params=params)

    if cache_file_path:
        resp_json = _get_cached_response(cache_file_path, req.url, cache_json_header)
        if resp_json is not None:
            return resp_json

    with graph_get_resource_lock:
        resp = session.send(req, timeout=GRAPH_API_TIMEOUT)

    resp.raise_for_status()

    resp_json = _get_content_as_json(resp)

    if cache_file_path:
        _write_cached_response(cache_file_path, cache_json_header, resp, resp_json)

    return resp_json


def graph_get_mailbox_root(
    session: sessions.Session,
    graph_config: GraphApiConfig,
) -> GraphFolderT:
    """Get the mailbox "Top of Information Store" folder.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.

    Returns:
        GraphFolderT

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
        GraphEmptyRootFolderError: Root information store cannot be retrieved.
    """
    url = f"{GRAPH_API_URL}/me/mailFolders/msgfolderroot"

    folder = graph_get_resource(session, graph_config, url)
    if not folder:
        errmsg = "Failed to get the root information store of the mailbox."
        raise GraphEmptyRootFolderError(errmsg)

    return folder


def graph_enum_folders(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    parent_folder: GraphFolderT | None,
    *,
    include_internal: bool = False,
) -> dict[str, GraphFolderT]:
    """Enumerate mailbox root folders or subfolders of a specific folder.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        parent_folder (GraphFolderT | None): The parent folder or None to list mailbox root folders.
        include_internal (bool, optional): Whether to include internal folders. Defaults to False.

    Returns:
        dict[str, GraphFolderT]: A dictionary mapping folder names (lower cased) to folder properties.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
        GraphEmptyRootFolderError: If the root information store of the mailbox cannot be retrieved.
    """
    folders: dict[str, GraphFolderT] = {}

    if not parent_folder:
        # If no parent folder is specified, use the root information store of the mailbox.
        # An unqualified '/mailFolders' query returns a flat list of all folders :/
        parent_folder = graph_get_mailbox_root(session, graph_config)
        parent_folder["folderParent"] = None
        parent_folder["folderPath"] = ""

    if parent_folder["childFolderCount"] == 0:
        return folders

    url = f"{GRAPH_API_URL}/me/mailFolders/{parent_folder['id']}/childFolders"
    params = {"$top": GRAPH_API_BATCH_MAX}

    while url:
        data = graph_get_resource(session, graph_config, url, params=params)

        for folder in data.get("value", []):
            is_internal_folder = (
                not parent_folder["folderParent"]
                and folder["wellKnownName"]
                and folder["wellKnownName"].lower() in INTERNAL_FOLDERS
            )

            # Exclude toplevel wellknown folders from export.
            if is_internal_folder and not include_internal:
                continue

            # Add some custom properties.
            folder["internalUse"] = is_internal_folder
            folder["folderParent"] = parent_folder
            folder["folderPath"] = (
                folder["displayName"]
                if not parent_folder
                else str(Path(*[parent_folder["folderPath"], folder["displayName"]]))
            )

            folders[folder["displayName"].lower()] = folder

        url = data.get("@odata.nextLink")
        params = None  # these are included in nextLink automatically

    return folders


def graph_get_folder_by_id(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_id: str,
) -> GraphFolderT | None:
    """Get a mail folder by its ID.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder_id (str): The ID of the folder to retrieve.

    Returns:
        GraphFolderT | None: The folder object if found, otherwise None.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    folder = None
    url = f"{GRAPH_API_URL}/me/mailFolders/{folder_id}"

    try:
        folder = graph_get_resource(session, graph_config, url)
    except requests.HTTPError as err:
        # If not found, API returns 400 Client Error: Bad Request for url: ...
        if err.response is None or err.response.status_code != 400:  # noqa: PLR2004
            raise

    return folder


def graph_get_folder_by_path(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_path: str | Path,
    parent_folder: GraphFolderT | None,
) -> GraphFolderT:
    """Get a mailbox root folder by name or a subfolder of a folder by its qualified path.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder_path (str | Path): The full path of the folder to retrieve, case-insensitive.
        parent_folder (GraphFolderT | None): The parent folder or None to search the mailbox root.

    Returns:
        GraphFolderT: The folder object.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        GraphEmptyRootFolderError: If the root information store of the mailbox cannot be retrieved.
        GraphFolderNotFoundError: Raised if any of the specified folder paths cannot be resolved.
    """
    folder_path = folder_path if isinstance(folder_path, Path) else Path(folder_path)
    path_parts = list(folder_path.parts)
    folder_name = path_parts.pop(0)
    folder = None

    folders_by_name = graph_enum_folders(
        session,
        graph_config,
        parent_folder,
        include_internal=True,
    )

    folder = folders_by_name.get(folder_name.lower())
    if not folder:
        fullpath = Path(parent_folder["folderPath"]) / folder_name if parent_folder else folder_name
        errmsg = f"'{fullpath}': No such mail folder."
        raise GraphFolderNotFoundError(errmsg)

    if folder["internalUse"]:
        errmsg = f"'{folder_name}': Folder is for internal use only."
        raise GraphFolderNotFoundError(errmsg)

    if not path_parts:
        return folder

    # Recursively navigate children by the remaining path parts.
    return graph_get_folder_by_path(session, graph_config, Path(*path_parts), folder)


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


def graph_walk_folders(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folders: Iterable[GraphFolderT],
    maxdepth: int,
) -> list[GraphFolderT]:
    """Build a dictionary containing a folder and all it's subfolders.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folders (Iterable[GraphFolderT]): Iterable of folder objects to walk.
        maxdepth (int): Maximum depth to traverse the folder hierarchy.

    Returns:
        list[GraphFolderT]: List of folder objects.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """

    def _recurse_subfolder(folder: GraphFolderT, folders_by_id: dict[str, GraphFolderT], depth: int) -> None:
        folders_by_id[folder["id"]] = folder
        if depth > 1:
            for subfolder in graph_enum_folders(session, graph_config, folder).values():
                _recurse_subfolder(subfolder, folders_by_id, depth - 1)

    foldertree: dict[str, GraphFolderT] = {}

    # Sort paths shallow to deep.
    folders = sorted(folders, key=itemgetter("folderPath"))

    for folder in folders:
        foldertree[folder["id"]] = folder
        if maxdepth > 1:
            for subfolder in graph_enum_folders(session, graph_config, folder).values():
                _recurse_subfolder(subfolder, foldertree, maxdepth - 1)

    return list(foldertree.values())


def graph_walk_mailbox(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    maxdepth: int,
) -> list[GraphFolderT]:
    """Build a list containing all mailbox folders and subfolders.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        maxdepth (int): Maximum depth to traverse the folder hierarchy.

    Returns:
        list[GraphFolderT]: List of folder objects.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
        GraphEmptyRootFolderError: If the root information store of the mailbox cannot be retrieved.
    """
    root_folders = graph_enum_folders(session, graph_config, MAILBOX_ROOT)
    return graph_walk_folders(session, graph_config, root_folders.values(), maxdepth=maxdepth)


def graph_resolve_folder_paths(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_paths: list[str] | None,
    *,
    maxdepth: int,
) -> list[GraphFolderT]:
    """Resolve folder paths to folder objects, optionally including subfolders.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder_paths (list[str] | None): List of folder paths to resolve. If None, all root folders are considered.
        maxdepth (int): Maximum depth to traverse the folder hierarchy.

    Returns:
        list[GraphFolderT]: List of resolved folder objects, including subfolders if requested.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
        GraphFolderNotFoundError: Raised if any of the specified folder paths cannot be resolved.
        GraphEmptyRootFolderError: If the root information store of the mailbox cannot be retrieved.
    """
    folders: list[GraphFolderT] = []

    if folder_paths is not None:
        # Ignore duplicate names while preserving original casing
        folder_paths = list({n.lower(): n for n in folder_paths}.values())

        for folder_path in folder_paths:
            folder = graph_get_folder_by_path(session, graph_config, folder_path, MAILBOX_ROOT)
            folders.append(folder)
    else:
        folders = list(graph_enum_folders(session, graph_config, MAILBOX_ROOT).values())

    return graph_walk_folders(session, graph_config, folders, maxdepth)


def print_graph_request_exception(console: Console, err: requests.RequestException) -> None:
    """Print error message resulting from a Graph API request exception.

    Args:
        console (Console): The console to print messages to.
        err (requests.RequestException): The exception.
    """
    resp_error_code, resp_error_message = None, None

    if err.response is not None:
        # See if Graph API returned an error message in the response body
        resp_json = err.response.json()
        resp_error = resp_json.get("error", {})
        resp_error_code, resp_error_message = resp_error.get("code"), resp_error.get("message")

    if resp_error_code and resp_error_message:
        console.print(f"[red]{resp_error_message} [{resp_error_code}]; {err}")
    else:
        console.print(f"[red]{err}")


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
    full_hash = hashlib.sha256(content.encode("utf-8")).hexdigest()
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
    path_names_set, path_names_lock = unique_file_path_sets[p.parent]

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


def exportstat_incr(export_stats: ExportStatsT, stat: ExportStat, incr: int = 1) -> None:
    """Increment specified export stat.

    Args:
        export_stats (ExportStatsT): Export stats dictionary.
        stat (ExportStat): Stat to increment.
        incr (int, optional): Increment amount. Defaults to 1.
    """
    export_stats[stat] = export_stats.get(stat, 0) + incr


def log_message_notice(message: GraphMessageT, errmsg: str, level: int = logging.WARNING) -> None:
    """Log a warning related to a message.

    Args:
        message (GraphMessageT): Message object.
        errmsg (str): Warning description.
        level (int): Logging level code (e.g., logging.WARNING).
    """
    logger.log(level, 'message("%s"): %s', message["subject"], errmsg)


def write_output_file(
    file_path: Path,
    file_text: str,
    file_date: datetime | None,
    export_opts: ExportOptions,
) -> None:
    """Write output file and set its modification date.

    Args:
        file_path (Path): Output file path.
        file_text (str): File content.
        file_date (datetime | None): Modification date. If None, the file's modification date will not be set.
        export_opts (ExportOptions): Export options.
    """
    if not export_opts.dry_run:
        file_path.write_text(file_text, encoding="utf-8")

        if file_date:
            os.utime(file_path, (file_date.timestamp(), file_date.timestamp()))

        logger.debug("Wrote file to '%s'", file_path.name)
    else:
        logger.debug("Wrote file to '%s' (dry run)", file_path.name)


def _get_cache_file_path_and_header(
    graph_config: GraphApiConfig,
    method: str,
    url: str,
    *,
    params: dict[str, Any] | None,
    cache_key: dict[str, Any] | None = None,
) -> tuple[Path | None, dict[str, Any]]:
    """Get the cache file path and header for a Graph API request.

    Args:
        graph_config (GraphApiConfig): Graph API configuration.
        method (str): HTTP method (e.g., "GET", "POST").
        url (str): The request URL.
        params (dict[str, Any] | None): Query parameters for the request.
        cache_key (dict, optional): Cache key to add specificity. Stored in the cached file as a sha256 hash.

    Returns:
        tuple[Path | None, dict[str, Any]]: (File path to write to, JSON header used for the cache key)
    """
    if not (graph_config.cache_enabled and graph_config.cache_dir):
        return None, {}

    cache_json_header: dict[str, Any] = {
        "@etag": "",
        "request": {
            "method": method,
            "url": url,
            "params": params,
        },
        "auth": {
            "environment": graph_config.token_info.environment,
            "tenant": graph_config.token_info.tenant,
            "client": graph_config.token_info.client,
            "username": graph_config.token_info.username,
        },
    }

    if cache_key:
        cache_json_header["@etag"] = hashlib.sha256(json.dumps(cache_key, sort_keys=True).encode("utf-8")).hexdigest()

    cache_file_key = json.dumps(cache_json_header, sort_keys=True)  # sort to keep deterministic
    cache_file_name = get_hash_filename(cache_file_key, "graph_api", ".json")
    cache_file_path = graph_config.cache_dir / cache_file_name

    return cache_file_path, cache_json_header


def _get_cached_response(cache_file_path: Path, req_url: str, cache_json_header: dict) -> GraphApiReturnT | None:
    """Read cached response from cache.

    Args:
        cache_file_path (Path): Path to the cache file.
        req_url (str): The request URL.
        cache_json_header (dict): JSON header used for the cache key.

    Returns:
        GraphApiReturnT | None: JSON object or None if request was not found in cache.
    """
    if not (cache_file_path and cache_file_path.is_file()):
        return None

    cached_resp_json = json.loads(cache_file_path.read_text(encoding="utf-8"))

    response = cached_resp_json.get("response")
    if not response:
        return None

    # The cached request must match the current request and authentication details.
    for key in ("request", "auth", "@etag"):
        if cache_json_header[key] != cached_resp_json[key]:
            logger.warning(
                "Graph API cache mismatch! %s(%s) != cached(%s)",
                key,
                cache_json_header[key],
                cached_resp_json[key],
            )
            return None

    logger.debug(
        'Using cached Graph API response for "%s %s" - %d',
        cache_json_header["request"]["method"],
        req_url,
        response.get("contentLength"),
    )

    return response["content"]


def _get_content_type_field(json_data: dict[str, Any], json_field: str) -> str | None:
    """Extract the Content-Type field from a JSON object.

    Args:
        json_data (dict[str, Any]): The JSON object.
        json_field (str): The name of the field.

    Returns:
        str | None: The content type if present, otherwise None.
    """
    # e.g. "application/json; odata.metadata=minimal; odata.streaming=true; IEEE754Compatible=false; charset=utf-8"
    content_type = json_data.get(json_field)
    if content_type is None:
        return None
    return content_type.split(";", 1)[0].strip().lower()


def _get_content_as_json(resp: requests.Response) -> dict[str, Any]:
    """Return response content as a JSON object, even if server replies with another content type.

    Args:
        resp (requests.Response): The response.

    Returns:
        dict[str, Any]: JSON object.
    """
    # e.g. "application/json; odata.metadata=minimal; odata.streaming=true; IEEE754Compatible=false; charset=utf-8"
    content_type = resp.headers.get("Content-Type", "").split(";", 1)[0].strip().lower()

    if content_type == "application/json":
        return resp.json()

    # Most Graph API responses are JSON but some endpoints such as '/attachments/id/$value'
    # return plain text or other content types.
    # Property names are consistent with those returned by '/attachments/id/$value'.

    if content_type == "text/plain":
        return {
            "contentType": content_type,
            "contentBytes": resp.text,
            "size": len(resp.text),
        }

    return {
        "contentType": content_type,
        "contentBytes": resp.content,
        "size": len(resp.content),
    }


def _write_cached_response(
    cache_file_path: Path,
    cache_json_header: dict[str, Any],
    resp: Response,
    resp_json: dict[str, Any],
) -> None:
    """Write a JSON response to cache directory.

    Args:
        cache_file_path (Path): Cache file path to write the response to.
        cache_json_header (dict[str, Any]): JSON header containing request and authentication details.
        resp (Response): The HTTP response object.
        resp_json (dict[str, Any]): The JSON response to be cached.
    """
    if cache_file_path:
        cache_file_path.write_text(
            json.dumps(
                cache_json_header
                | {
                    "response": {
                        "date": datetime_now().isoformat(),
                        "contentType": "application/json",
                        "contentLength": resp_json.get("size") or len(resp.text),
                        "content": resp_json,
                    },
                },
            ),
            encoding="utf-8",
        )


def _batch_raise_for_status(  # noqa: C901
    url: str,
    http_resp: requests.Response,
    response_json: dict[str, Any],
    payload_json: dict[str, Any],
) -> None:
    """Raise an HTTPError exception if any batch individual responses returned an error code.

    Args:
        url (str): The URL.
        http_resp (requests.Response): The POST response.
        response_json (dict[str, Any]): The JSON response from the batch request.
        payload_json (dict[str, Any]): The POST request JSON payload.

    Raises:
        requests.HTTPError
    """

    def _decode_blob(body: str) -> tuple[str, str]:
        code = "--"

        try:
            body = base64.b64decode(body).decode("utf-8")
        except binascii.Error as err:
            message = f"Unexpected batch request error response body: {body}; {err}"
        else:
            json_body = json.loads(body)

            if isinstance(json_body, dict):
                error = json_body.get("error", {})
                code = error.get("code")  # BadRequest
                message = error.get("message")  # "Resource not found for the segment 'messages'."
            else:
                message = f"Unexpected batch request error response body: {body}"

        return code, message

    # The batch response contains an array of individual responses.
    for response in response_json.get("responses", []):
        status_code = int(response.get("status", 0))
        if status_code == 200:  # noqa: PLR2004
            continue

        request_id = response.get("id")
        headers = response.get("headers", {})
        content_type = _get_content_type_field(headers, "Content-Type")
        body = response.get("body")
        code = "--"

        # See if we can get URL for individual request
        request_url_by_id = {req.get("id"): req.get("url") for req in payload_json.get("requests", [])}
        url = request_url_by_id.get(request_id, url) if request_id else url

        if content_type != "application/json":
            message = f"Unexpected error response content type '{content_type}'"
        elif isinstance(body, dict):
            error = body.get("error", {})
            code = error.get("code")  # BadRequest
            message = error.get("message")  # "Resource not found for the segment 'messages'."
        else:
            # The headers don't indicate Base64 encoded JSON, but it probably is.
            code, message = _decode_blob(body)

        if 400 <= status_code < 500:  # noqa: PLR2004
            http_error_msg = f"{message} {status_code} Client Error: {code} for url in batch: {url}"
        elif 500 <= status_code < 600:  # noqa: PLR2004
            http_error_msg = f"{message} {status_code} Server Error: {code} for url in batch: {url}"
        else:
            http_error_msg = f"{message} {status_code} Unexpected Error: {code} for url in batch: {url}"

        raise requests.HTTPError(http_error_msg, response=http_resp)


def graph_post_resource(  # noqa: PLR0913, PLR0917
    session: sessions.Session,
    graph_config: GraphApiConfig,
    url: str,
    json_data: dict[str, Any],
    params: dict[str, Any] | None = None,
    headers: dict[str, Any] | None = None,
    cache_key: dict[str, Any] | None = None,
) -> GraphApiReturnT:
    """Execute a POST request to Microsoft Graph API.

    Args:
        session (sessions.Session): Requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        url (str): Entity URL.
        json_data (dict[str, Any]): JSON body for the POST request.
        params (dict[str, Any], optional): Query parameters for the POST request. Defaults to None.
        headers (dict[str, Any], optional): HTTP headers for the POST request. Defaults to None.
        cache_key (dict, optional): Cache key for the batch request. Defaults to `json_data` if not provided.

    Returns:
        GraphApiReturnT: JSON response from the Microsoft Graph API.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    req_headers = {
        "Authorization": f"Bearer {graph_config.access_token}",
        "Content-Type": "application/json",
    }

    if headers:
        req_headers.update(headers)

    req = requests.Request("POST", url, headers=req_headers, params=params, json=json_data).prepare()
    assert req.url is not None  # noqa: S101

    cache_file_path, cache_json_header = _get_cache_file_path_and_header(
        graph_config,
        "POST",
        url,
        params=params,
        cache_key=cache_key or json_data,
    )

    if cache_file_path:
        resp_json = _get_cached_response(cache_file_path, req.url, cache_json_header)
        if resp_json is not None:
            return resp_json

    with graph_get_resource_lock:
        resp = session.send(req, timeout=GRAPH_API_TIMEOUT)

    resp.raise_for_status()

    resp_json = _get_content_as_json(resp)

    if url.endswith("/$batch"):
        _batch_raise_for_status(url, resp, resp_json, json_data)

    if cache_file_path:
        _write_cached_response(cache_file_path, cache_json_header, resp, resp_json)

    return resp_json


def graph_get_resource_batch(  # noqa: PLR0913
    session: sessions.Session,
    graph_config: GraphApiConfig,
    urls: dict[str, str],
    *,
    params: dict[str, Any] | None = None,
    headers: dict[str, Any] | None = None,
    cache_key: dict[str, Any] | None = None,
    sequence: bool = True,
    short_key_factory: Callable[[int], str] | None = None,
) -> dict[str, GraphApiReturnT]:
    """Execute a batch of GET request to Microsoft Graph API.

    https://learn.microsoft.com/en-us/graph/json-batching

    The Outlook Graph API limits requests to 4 concurrent requests and will fail if the limit is exceeded.
    When sending a batch request, Microsoft Graph sends the Outlook service up to 4 individual requests
    from the batch at a time. If another request is made during this time, the limit will have been exceeded
    and the batch will fail. The solution when using multiple threads is to sequence the batch individual requests.

    Args:
        session (sessions.Session): Requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        urls (dict[str, str]): Entity URLs mapped by request ID.
        params (dict, optional): Query parameters for the GET request. Defaults to None.
        headers (dict, optional): HTTP headers for the GET request. Defaults to None.
        cache_key (dict, optional): Cache key for the batch request. Defaults to requested URLs if not provided.
        sequence (bool, optional): True if individual batched requests should be executed in sequence. Defaults to True.
        short_key_factory (Callable[[int], str], optional): Factory function to generate short request IDs. Defaults to None.

    Returns:
        dict[str, GraphApiReturnT]: Dictionary mapping request IDs to their JSON responses.

    Raises:
        requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
        OSError derived error (e.g. file system errors, permission errors, etc.)
    """
    if not urls:
        msg = "No URLs provided for the batch request."
        raise ValueError(msg)

    logger.debug(
        "Graph API POST request for batch of %d URLs (%s)",
        len(urls),
        "sequenced" if sequence else "parallel",
    )

    depends_id: str | None = None
    requests: list[dict[str, Any]] = []
    request_id_map: dict[str, str] = {}

    for request_id, url in urls.items():
        short_id: str = request_id

        # I can't find an official statement, but experience says Graph API truncates long IDs.
        # As a convenience, we can translate long IDs into shorter batch specific IDs.
        if short_key_factory is not None:
            short_id = short_key_factory(len(requests) + 1)
            request_id_map[short_id] = request_id

        request: dict[str, str | list[str]] = {
            "id": short_id,
            "method": "GET",
            "url": url.removeprefix(GRAPH_API_URL),
        }

        if sequence and depends_id:
            request["dependsOn"] = [depends_id]
        depends_id = short_id

        requests.append(request)

    data = graph_post_resource(
        session,
        graph_config,
        f"{GRAPH_API_URL}/$batch",
        json_data={"requests": requests},
        params=params,
        headers=headers,
        cache_key=cache_key,
    )

    responses = data.get("responses", [])

    if request_id_map:
        return {request_id_map[resp["id"]]: resp for resp in responses}
    return {resp["id"]: resp for resp in responses}


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
