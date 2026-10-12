"""outlook2obsidian shared variables and functions."""

import hashlib
import json
import logging
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import msal
import requests
from requests import sessions
from rich.console import Console

######################################################################
# Constants

APP_NAME = "outlook2obsidian"
AUTHORITY_ENVIRONMENT_DEFAULT = "login.microsoftonline.com"
GRAPH_API_URL = "https://graph.microsoft.com/beta"
GRAPH_SCOPES = ["Mail.Read"]
GRAPH_API_BATCH_MIN = 25  # Arbitrary limit
GRAPH_API_BATCH_MAX = 100  # Arbitrary limit
MAILBOX_ROOT = None
VERBOSE_LOGGING_LEVEL_ERROR = 0  # logging.ERROR

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
    "expires_on",
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

GRAPH_ACCESS_TOKEN_FIELDS_SET = set(GRAPH_ACCESS_TOKEN_FIELDS)
GRAPH_CACHED_TOKEN_FIELDS_SET = set(GRAPH_CACHED_TOKEN_FIELDS)
GRAPH_ACCOUNT_FIELDS_SET = set(GRAPH_ACCOUNT_FIELDS)
GRAPH_FOLDER_FIELDS_SET = set(GRAPH_FOLDER_FIELDS)
GRAPH_MESSAGE_FIELDS_SET = set(GRAPH_MESSAGE_FIELDS)
GRAPH_APPMETADATA_FIELDS_SET = set(GRAPH_APPMETADATA_FIELDS)

######################################################################
# Types and classes


GraphApiReturnT = dict[str, Any]  # The return type of Graph API calls
GraphAccessTokenT = dict[str, Any]  # GRAPH_ACCESS_TOKEN_FIELDS
GraphCachedTokenT = dict[str, Any]  # GRAPH_TOKEN_FIELDS
GraphAccountT = dict[str, Any]  # GRAPH_ACCOUNT_FIELDS
GraphFolderT = dict[str, Any]  # GRAPH_FOLDER_FIELDS
GraphMessageT = dict[str, Any]  # GRAPH_MESSAGE_FIELDS
GraphAppMetadataT = dict[str, Any]  # GRAPH_APPMETADATA_FIELDS

NoteMetaPropT = dict[str, str | int | Sequence[str | int]]


@dataclass
class CachedAccessTokenInfo:
    """Relevant details from SerializableTokenCache."""

    environment: str
    tenant: str
    client: str
    username: str
    cached_at: datetime
    expires_on: datetime


@dataclass
class GraphApiConfig:
    """Configuration for HTTP caching."""

    access_token: str
    cache_dir: Path | None
    cache_enabled: bool
    token_info: CachedAccessTokenInfo


class AuthRequiredError(ValueError):
    """Raised when user authentication is required."""


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

    export_metaprops: NoteMetaPropT | None = None  # set if export command has `--metaprop` options

    def __post_init__(self) -> None:  # noqa: D105
        self.common_opts = CommonCommandOptions()


######################################################################
# Globals

console = Console(soft_wrap=True)  # Enable soft wrap so newlines are not inserted to wrap lines wider than the terminal
err_console = Console(stderr=True, soft_wrap=True, style="red", highlight=False)
logger = logging.getLogger(APP_NAME)  # initialized in __main__


######################################################################
# Helpers


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
        expires_on=datetime.fromtimestamp(int(access_token["expires_on"]), tz=timezone.utc),  # noqa: UP017
    )


def get_graph_config(common_opts: CommonCommandOptions) -> GraphApiConfig:
    """Get cached Graph API configuration details including access token from the MSAL cache.

    Returns:
        GraphApiConfig: Graph API configuration.

    Raises:
        AuthRequiredError: Raised if access token is missing or invalid.
        ValueError: MSAL-related error, such as an invalid MSAL environment or invalid token cache.
    """
    if not common_opts.auth_file_path:
        errmsg = "Internal error: Authentication file path is not set."
        raise ValueError(errmsg)

    # First see if there is cached authentication information.
    token_info = get_token_cache_info(common_opts.auth_file_path)
    if not token_info:
        raise AuthRequiredError(AUTH_REQUIRED_ERROR)

    # Try to acquire a valid access token using cached credentials.
    _, _, _, token = get_graph_client(common_opts.auth_file_path, use_cache=True)
    if not token or "access_token" not in token:
        raise AuthRequiredError(AUTH_EXPIRED_ERROR)

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
        AuthRequiredError: Cache is being used and access token is unavailable or invalid.
        ValueError: MSAL-related error, such as an invalid MSAL environment or invalid token cache.
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
            raise AuthRequiredError(AUTH_EXPIRED_ERROR)
        raise AuthRequiredError(AUTH_REQUIRED_ERROR)

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
            raise AuthRequiredError(AUTH_EXPIRED_ERROR)

        if "access_token" not in token:
            errmsg = "Graph API returned an access token that does not have an 'access_token' property."
            raise ValueError(errmsg)

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
) -> GraphApiReturnT:
    """Execute a GET request to Microsoft Graph API.

    Args:
        session (sessions.Session): Requests session to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        url (str): Entity URL.
        params (dict, optional): Query parameters for the GET request. Defaults to None.

    Returns:
        GraphApiReturnT: JSON response from the Microsoft Graph API.

    Raises:
        requests.ConnectionError: If there is a network-related error during the HTTP request.
        requests.Timeout: If the HTTP request times out.
        requests.HTTPError: If the HTTP request to the Microsoft Graph API fails.
        requests.RequestException: For other errors during the HTTP request (base class.)
        requests.exceptions.JSONDecodeError: If the response body is not valid JSON.
    """

    def _get_cache_file_path_and_header() -> tuple[Path | None, dict[str, Any]]:
        if not (graph_config.cache_enabled and graph_config.cache_dir):
            return None, {}

        cache_json_header = {
            "request": {
                "method": "GET",
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

        cache_file_key = json.dumps(cache_json_header, sort_keys=True)  # sort to keep deterministic
        cache_file_name = Path("graph_api_" + hashlib.sha256(cache_file_key.encode("utf-8")).hexdigest() + ".json")
        cache_file_path = graph_config.cache_dir / cache_file_name

        return cache_file_path, cache_json_header

    def _get_cached_response(cache_file_path: Path, req_url: str) -> GraphApiReturnT | None:
        if cache_file_path.is_file():
            cached_resp_json = json.loads(cache_file_path.read_text(encoding="utf-8"))

            response = cached_resp_json.get("response")
            if response:
                logger.debug(
                    'Using cached Graph API response for "GET %s" - %d',
                    req_url,
                    response.get("content-length"),
                )

                return response["content-json"]

        return None

    headers = {
        "Authorization": f"Bearer {graph_config.access_token}",
        "Content-Type": "application/json",
    }

    req = requests.Request("GET", url, headers=headers, params=params).prepare()
    assert req.url is not None  # noqa: S101

    cache_file_path, cache_json_header = _get_cache_file_path_and_header()

    if cache_file_path:
        resp_json = _get_cached_response(cache_file_path, req.url)
        if resp_json is not None:
            return resp_json

    resp = session.send(req, timeout=30)

    resp.raise_for_status()

    resp_json = resp.json()

    if cache_file_path:
        cache_file_path.write_text(
            json.dumps(
                cache_json_header
                | {
                    "response": {
                        "content-date": datetime.now(tz=timezone.utc).isoformat(),  # noqa: UP017
                        "content-json": resp_json,
                        "content-length": len(resp.text),
                    },
                },
            ),
            encoding="utf-8",
        )

    return resp_json


def graph_get_mailbox_root(
    session: sessions.Session,
    graph_config: GraphApiConfig,
) -> GraphFolderT | None:
    """Get the mailbox "Top of Information Store" folder.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.

    Returns:
        GraphFolderT

    Raises:
        requests.RequestException derived error.
    """
    url = f"{GRAPH_API_URL}/me/mailFolders/msgfolderroot"

    return graph_get_resource(session, graph_config, url)


def graph_enum_folders(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    parent_folder: GraphFolderT | None,
) -> dict[str, GraphFolderT]:
    """Enumerate mailbox root folders or subfolders of a specific folder.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        parent_folder (GraphFolderT | None): The parent folder or None to list mailbox root folders.

    Returns:
        dict[str, GraphFolderT]: A dictionary mapping folder names (lower cased) to folder properties.

    Raises:
        requests.RequestException derived error.
        ValueError: If the root information store of the mailbox cannot be retrieved.
    """
    folders: dict[str, GraphFolderT] = {}

    if not parent_folder:
        # If no parent folder is specified, use the root information store of the mailbox.
        # An unqualified '/mailFolders' query returns a flat list of all folders :/
        parent_folder = graph_get_mailbox_root(session, graph_config)
        if not parent_folder:
            errmsg = "Failed to get the root information store of the mailbox."
            raise ValueError(errmsg)

    if parent_folder["childFolderCount"] == 0:
        return folders

    url = f"{GRAPH_API_URL}/me/mailFolders/{parent_folder['id']}/childFolders"
    params = {"$top": GRAPH_API_BATCH_MAX}

    while url:
        data = graph_get_resource(session, graph_config, url, params=params)
        folders.update({folder["displayName"].lower(): folder for folder in data.get("value", [])})
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
        requests.RequestException derived error.
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


def graph_get_folder_by_name(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_name: str,
    parent_folder: GraphFolderT | None,
) -> GraphFolderT | None:
    """Get a mailbox root folder or a subfolder of a folder by its name.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder_name (str): The name of the folder to retrieve.
        parent_folder (GraphFolderT | None): The parent folder or None to search mailbox root.

    Returns:
        GraphFolderT | None: The folder object if found, otherwise None.

    Raises:
        requests.RequestException derived error.
    """
    folders = graph_enum_folders(session, graph_config, parent_folder)
    return folders.get(folder_name.lower(), None) if folders else None


def graph_get_folder_by_path(
    session: sessions.Session,
    graph_config: GraphApiConfig,
    folder_path: Path,
    parent_folder: GraphFolderT | None,
) -> GraphFolderT | None:
    """Get a mail folder by its path.

    Args:
        session (sessions.Session): Session to use to pool requests.
        graph_config (GraphApiConfig): Graph API configuration.
        folder_path (Path): The full path of the folder to retrieve.
        parent_folder (GraphFolderT | None): The parent folder or None to search the mailbox root.

    Returns:
        GraphFolderT | None: The folder object if found, otherwise None.

    Raises:
        requests.RequestException derived error.
    """
    path_parts = list(Path(folder_path).parts)

    folder = graph_get_folder_by_name(session, graph_config, path_parts.pop(0), parent_folder)
    if not folder:
        return None

    # Add a custom property with the full path to the folder.
    folder["folderPath"] = (
        folder["displayName"] if not parent_folder else str(Path(*[parent_folder["folderPath"], folder["displayName"]]))
    )

    if path_parts:
        # Recursively get the child folder by the remaining path parts.
        folder = graph_get_folder_by_path(session, graph_config, Path(*path_parts), folder)

    return folder
