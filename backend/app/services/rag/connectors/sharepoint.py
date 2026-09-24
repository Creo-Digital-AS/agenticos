"""SharePoint sync connector for RAG ingestion.

Reads a site's document library through Microsoft Graph, authenticating as an
Entra ID app registration: an `entra_app` secret in the organization's vault,
named by the source's `secret_id` and unsealed by whoever runs the sync. There
is no deployment-wide fallback, for the reason the Drive and S3 connectors give
- a fallback means one tenant's `site_path` chooses what is read under the
*operator's* registration.

**App-only, not delegated**, because a scheduled sync runs at 3am with nobody
present to consent and no refresh token anybody is keeping alive. The cost of
that is stated rather than hidden: an application permission is not narrowed by
who configured the source, so `Sites.Read.All` grants this deployment every site
in the directory, and the setup a person should follow is `Sites.Selected` plus
a read grant on the one site. `docs/howto/configure-sync-sources.md` is where
that argument is made to the person doing it.

Setup:
1. Register an application in Entra ID
2. Give it the Microsoft Graph *application* permission `Sites.Selected`, and
   grant admin consent
3. Grant that application `read` on the specific site (Graph's
   `/sites/{id}/permissions`, or the SharePoint admin centre)
4. Create a client secret, and add its value to the Vault once as a Microsoft
   Entra app credential - each site is then a source pointing at that one secret

Everything here goes to `graph.microsoft.com` and `login.microsoftonline.com`,
both constants below. The configured hostname and paths are *path components* of
a Graph URL rather than addresses, which is why they are checked by
`remote_names` before they are interpolated: a value carrying `:` or `?` does
not reach a different server, it addresses a different Graph resource.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Any, ClassVar
from urllib.parse import quote

import anyio
import httpx
from pydantic import BaseModel, Field

from app.core.exceptions import BadRequestError
from app.core.secret_kinds import EntraAppSecret, SecretKind, StorableSecret
from app.services.rag.connectors import (
    BaseSyncConnector,
    ConfigRefusal,
    ConnectorConfig,
    RemoteFile,
)
from app.services.rag.remote_names import checked_sharepoint_host, checked_sharepoint_path

logger = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com"

# The only scope a client-credentials grant can ask for: Entra resolves
# `.default` to whatever application permissions the registration has been
# granted and consented for, so what this connector can reach is decided in the
# portal rather than here.
SCOPE = "https://graph.microsoft.com/.default"

_TIMEOUT = httpx.Timeout(30.0, read=120.0)

# Graph's own ceiling for `$top` on a children collection. Asking for it makes a
# library of two thousand files four round trips rather than twenty.
_PAGE_SIZE = 999


def graph_client() -> httpx.AsyncClient:
    """The client every request in this module is sent through.

    One factory rather than the same constructor in two methods, so the timeout
    and the redirect policy are decided once - `_fetch` depends on redirects
    *not* being followed, and that is a property of the client it is handed.
    """
    return httpx.AsyncClient(timeout=_TIMEOUT, follow_redirects=False)


class SharePointConfig(BaseModel):
    """Where a SharePoint source's documents live.

    The credential is not here - it is an `EntraAppSecret` the source names in
    `secret_id`, and the tenant and application ids travel with it rather than
    with this. That split is the one #937 made: this model says how to *find*
    the documents, and a reader of it learns nothing that would let them fetch
    any.

    `folder_path` is optional and empty means the library root. A site has one
    default document library and this connector reads that one - a site with
    several is several sources, which is also how its permissions are granted.
    """

    hostname: str = Field(
        title="SharePoint hostname",
        description="The tenant's SharePoint host, e.g. contoso.sharepoint.com",
    )
    site_path: str = Field(
        title="Site path",
        description="The site's server-relative path, e.g. /sites/Engineering",
    )
    folder_path: str | None = Field(
        default=None,
        title="Folder path",
        description="A folder inside the document library, e.g. Legal/Contracts - empty for the whole library",
    )
    include_subfolders: bool = Field(default=True, title="Include subfolders")


class SharePointConnector(BaseSyncConnector):
    """SharePoint document libraries, read through Microsoft Graph.

    One instance serves one sync - `sync_source_flow` builds it, lists once and
    downloads each file through it - which is what makes the two caches below
    safe. They are per-instance and never per-class: a token cached on the class
    would be one tenant's token answering another tenant's sync.

    **No change signal yet**, so a nightly sync costs the whole library in
    transfer even where nothing moved: the flow falls back to comparing a
    `content_hash`, which saves the embedding and not the download. Graph does
    offer one - `/drives/{id}/root/delta` answers with what changed since a
    token - and wiring it is the obvious next thing this connector wants
    (`docs/file-processing.md`, "What a new connector owes").
    """

    CONNECTOR_TYPE: ClassVar[str] = "sharepoint"
    DISPLAY_NAME: ClassVar[str] = "SharePoint"
    SECRET_KIND: ClassVar[SecretKind] = SecretKind.ENTRA_APP
    CONFIG_MODEL: ClassVar[type[BaseModel]] = SharePointConfig

    def __init__(self) -> None:
        # An access token lives an hour and a sync of five hundred files makes
        # five hundred `_fetch` calls, so minting one per call would be five
        # hundred round trips to Entra for a credential that has not changed.
        # No expiry tracked: this object does not outlive the sync that made it,
        # and a sync that runs for an hour has a token expiring mid-flight
        # whatever is cached - which arrives as a 401 the caller already reports.
        self._token: str | None = None
        # The drive `list_files` resolved. `_fetch` is handed a `RemoteFile` and
        # a destination, never the drive the file came from, so without this
        # every download would re-resolve the site and its library: two more
        # requests per file.
        self._drive_id: str | None = None

    def _app(self, credential: StorableSecret | None) -> EntraAppSecret:
        """The app registration this source authenticates as.

        Raises:
            BadRequestError: the source names no credential, its secret has been
                deleted, or the secret is not an Entra app registration.
        """
        if credential is None:
            raise BadRequestError(
                message=(
                    "This SharePoint source has no credential. Pick a Microsoft "
                    "Entra app in the Vault and point the source at it."
                )
            )
        if not isinstance(credential, EntraAppSecret):
            raise BadRequestError(
                message=(
                    "A SharePoint source needs a Microsoft Entra app credential, "
                    "and the one it names is not one."
                )
            )
        return credential

    async def _access_token(self, client: httpx.AsyncClient, app: EntraAppSecret) -> str:
        """Mint an app-only access token, or answer the one already minted.

        Raises:
            BadRequestError: Entra could not be reached, or refused the
                registration. Its own text is logged and not returned - a token
                endpoint's error body echoes the request it is about, and this
                request carries the client secret (the same rule
                `portals/google_oauth.py` states).
        """
        if self._token is not None:
            return self._token
        try:
            response = await client.post(
                f"{LOGIN}/{quote(app.tenant_id, safe='')}/oauth2/v2.0/token",
                data={
                    "client_id": app.client_id,
                    "client_secret": app.client_secret.get_secret_value(),
                    "scope": SCOPE,
                    "grant_type": "client_credentials",
                },
            )
        except httpx.HTTPError as exc:
            logger.warning("entra_token_unreachable", extra={"error": exc.__class__.__name__})
            raise BadRequestError(
                message="Microsoft Entra could not be reached to authenticate this source."
            ) from exc
        if response.status_code >= 400:
            # The status and nothing else. Entra's `error_description` restates
            # the request it is about, and this request is the one carrying the
            # client secret - the same reason `portals/google_oauth.py` logs a
            # status where `_graph` below logs a body.
            logger.warning("entra_token_refused", extra={"status": response.status_code})
            raise BadRequestError(
                message=(
                    "Microsoft Entra refused this source's credential "
                    f"(HTTP {response.status_code}). Check the tenant, the "
                    "application id and whether the client secret has expired."
                )
            )
        token = response.json().get("access_token")
        if not isinstance(token, str) or not token:
            raise BadRequestError(message="Microsoft Entra answered without an access token.")
        self._token = token
        return token

    async def _graph(
        self, client: httpx.AsyncClient, token: str, url: str, *, about: str
    ) -> dict[str, Any]:
        """One Graph GET, answered as JSON.

        Raises:
            BadRequestError: Graph could not be reached or refused the request.
                The status is repeated because it is the diagnosis - 403 is a
                permission never granted, 404 a site path that does not resolve -
                and `about` names which lookup it was, since a sync makes several.
                Graph's own body is logged rather than stored: it is rendered on
                the source's sync history to everyone who can see the collection
                (`services/rag/failures.py`). Logging it at all is safe in a way
                the token endpoint's body is not - the credential travels in a
                header here, so there is nothing of it for an error to echo.
        """
        try:
            response = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        except httpx.HTTPError as exc:
            logger.warning(
                "graph_unreachable", extra={"about": about, "error": exc.__class__.__name__}
            )
            raise BadRequestError(
                message=f"Microsoft Graph could not be reached while reading the {about}."
            ) from exc
        if response.status_code >= 400:
            logger.warning(
                "graph_refused",
                extra={"about": about, "status": response.status_code, "body": response.text},
            )
            raise BadRequestError(
                message=(
                    f"Microsoft Graph refused the {about} (HTTP {response.status_code}). "
                    "Check the hostname and site path, and that the app registration "
                    "has been granted read access to this site."
                )
            )
        payload: dict[str, Any] = response.json()
        return payload

    async def _resolve_drive(
        self, client: httpx.AsyncClient, token: str, config: ConnectorConfig
    ) -> str:
        """The id of the site's default document library, resolved once.

        Two lookups rather than one: Graph addresses a site by host and path and
        answers with an opaque id, and the drive hangs off that id. Both values
        are checked before they are interpolated - see `remote_names`.
        """
        if self._drive_id is not None:
            return self._drive_id
        hostname = checked_sharepoint_host(config.get("hostname"))
        site_path = checked_sharepoint_path(config.get("site_path"), what="site path")
        site = await self._graph(
            client, token, f"{GRAPH}/sites/{hostname}:/{site_path}", about="site"
        )
        drive = await self._graph(
            client, token, f"{GRAPH}/sites/{quote(str(site['id']), safe='')}/drive", about="library"
        )
        self._drive_id = str(drive["id"])
        return self._drive_id

    async def _children(
        self, client: httpx.AsyncClient, token: str, url: str
    ) -> list[dict[str, Any]]:
        """Every child of one folder, following Graph's paging to the end.

        `@odata.nextLink` is an absolute URL Graph builds, carrying the skip
        token and the `$top` already asked for - so it is followed as given
        rather than rebuilt, which is the only way to page a collection whose
        ordering Graph owns.
        """
        items: list[dict[str, Any]] = []
        next_url: str | None = f"{url}?$top={_PAGE_SIZE}"
        while next_url:
            page = await self._graph(client, token, next_url, about="folder listing")
            items.extend(page.get("value", []))
            link = page.get("@odata.nextLink")
            next_url = link if isinstance(link, str) else None
        return items

    async def _walk(
        self,
        client: httpx.AsyncClient,
        token: str,
        drive_id: str,
        url: str,
        *,
        include_subfolders: bool,
    ) -> list[RemoteFile]:
        """Every file under one folder, recursing when the source asked for it."""
        files: list[RemoteFile] = []
        for item in await self._children(client, token, url):
            if "folder" in item:
                if include_subfolders:
                    item_id = quote(str(item["id"]), safe="")
                    files.extend(
                        await self._walk(
                            client,
                            token,
                            drive_id,
                            f"{GRAPH}/drives/{drive_id}/items/{item_id}/children",
                            include_subfolders=include_subfolders,
                        )
                    )
                continue
            if "file" not in item:
                # A OneNote section, a bundle, a link to something elsewhere -
                # an item with neither facet has no bytes to download, and the
                # ingestion pipeline has nothing to parse.
                continue
            modified_at = None
            raw_modified = item.get("lastModifiedDateTime")
            if isinstance(raw_modified, str):
                modified_at = datetime.fromisoformat(raw_modified.replace("Z", "+00:00"))
            files.append(
                RemoteFile(
                    id=str(item["id"]),
                    name=str(item["name"]),
                    mime_type=item.get("file", {}).get("mimeType"),
                    size=item.get("size"),
                    modified_at=modified_at,
                    # The drive as well as the item: an item id is unique within
                    # its drive and not across them, so two libraries holding a
                    # file each would otherwise be one dedup key (#988's shape,
                    # where `s3://bucket/key` carries the bucket for the same
                    # reason).
                    source_path=f"sharepoint://{drive_id}/{item['id']}",
                )
            )
        return files

    async def list_files(
        self, config: ConnectorConfig, credential: StorableSecret | None
    ) -> list[RemoteFile]:
        """Every file in the configured library, or in one folder of it."""
        app = self._app(credential)
        # Every configured value is checked before the first request rather than
        # where each is interpolated: a stored row edited underneath the route's
        # own check should be refused without a token having been minted for it.
        folder_path = config.get("folder_path") or ""
        folder = checked_sharepoint_path(folder_path, what="folder path") if folder_path else ""
        include_subfolders = bool(config.get("include_subfolders", True))
        async with graph_client() as client:
            token = await self._access_token(client, app)
            drive_id = quote(await self._resolve_drive(client, token, config), safe="")
            if folder:
                url = f"{GRAPH}/drives/{drive_id}/root:/{folder}:/children"
            else:
                url = f"{GRAPH}/drives/{drive_id}/root/children"
            return await self._walk(
                client, token, drive_id, url, include_subfolders=include_subfolders
            )

    async def _fetch(
        self,
        file: RemoteFile,
        dest_path: Path,
        config: ConnectorConfig,
        credential: StorableSecret | None,
    ) -> None:
        """Stream one item's bytes to the path the base class chose.

        The redirect is walked by hand rather than by `follow_redirects=True`.
        `/content` answers 302 to a short-lived URL on a Microsoft CDN, and the
        bearer token has no business being sent there - it is scoped to Graph and
        the download URL authenticates itself. Whether a client strips an
        `Authorization` header across a redirect is a property of the client's
        version; not sending it is a property of this function.

        Streamed rather than fetched: a document library holds the files nobody
        wanted to email, and reading a 400MB one into memory to write it back
        out is a worker the next sync finds already dead.
        """
        app = self._app(credential)
        async with graph_client() as client:
            token = await self._access_token(client, app)
            drive_id = quote(await self._resolve_drive(client, token, config), safe="")
            item_id = quote(file.id, safe="")
            url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"
            try:
                async with client.stream(
                    "GET", url, headers={"Authorization": f"Bearer {token}"}
                ) as response:
                    if not response.is_redirect:
                        await self._write_stream(response, dest_path, file)
                        return
                    # Read before the context closes: the body of a 302 is a
                    # few bytes of XML nobody wants, but leaving it unread on a
                    # streamed response is what makes the connection unusable.
                    await response.aread()
                    # Joined against the request rather than taken bare: Graph
                    # answers an absolute URL, and a relative `Location` would
                    # otherwise raise out of `httpx.InvalidURL`, which is not an
                    # `HTTPError` and so not one of the failures below.
                    location = response.url.join(response.headers["location"])
                async with client.stream("GET", location) as download:
                    await self._write_stream(download, dest_path, file)
            except httpx.HTTPError as exc:
                logger.warning(
                    "graph_download_unreachable",
                    extra={"item": file.id, "error": exc.__class__.__name__},
                )
                raise BadRequestError(
                    message=f"Microsoft Graph could not be reached to download '{file.name}'."
                ) from exc

    async def _write_stream(
        self, response: httpx.Response, dest_path: Path, file: RemoteFile
    ) -> None:
        """Write a streamed response to `dest_path`, or refuse what it answered.

        Raises:
            BadRequestError: the response is not a 2xx. The file's name is in
                the message because it is the tenant's own and the reader is
                looking at a sync log full of other files; the body is not,
                for the reason `_graph` gives.
        """
        if response.status_code >= 400:
            await response.aread()
            logger.warning(
                "graph_download_refused",
                extra={"status": response.status_code, "item": file.id},
            )
            raise BadRequestError(
                message=(
                    f"Microsoft Graph refused the download of '{file.name}' "
                    f"(HTTP {response.status_code})."
                )
            )
        # `anyio.open_file`, not `open`: the write is inside the same event loop
        # the stream is being read on, so a blocking one stalls every other task
        # in the worker for the length of a document library's largest file.
        async with await anyio.open_file(dest_path, "wb") as handle:
            async for chunk in response.aiter_bytes():
                await handle.write(chunk)
        logger.info("Downloaded %s from SharePoint (%d bytes)", file.id, dest_path.stat().st_size)

    async def validate_config(self, config: ConnectorConfig) -> ConfigRefusal | None:
        """Refuse a config the wizard can still fix.

        Connectivity is not answered here - `validate_config` sees the config and
        not the credential, so "can this registration reach that site" is a
        question for the first sync. What is answered is the shape of what was
        typed, and it is answered here as well as where the URL is built so a
        hostile value is refused by the route that accepted it rather than by a
        sync log an hour later. The two cannot disagree: both ask `remote_names`.

        The field is named here because this is the one caller that was sent a
        form to mark; `remote_names` names none, for the reason its own module
        docstring gives.
        """
        refusal = await super().validate_config(config)
        if refusal is not None:
            return refusal
        try:
            checked_sharepoint_host(config["hostname"])
        except BadRequestError as exc:
            return ConfigRefusal(message=exc.message, field="hostname")
        try:
            checked_sharepoint_path(config["site_path"], what="site path")
        except BadRequestError as exc:
            return ConfigRefusal(message=exc.message, field="site_path")
        folder_path = config.get("folder_path")
        if folder_path:
            try:
                checked_sharepoint_path(folder_path, what="folder path")
            except BadRequestError as exc:
                return ConfigRefusal(message=exc.message, field="folder_path")
        return None
