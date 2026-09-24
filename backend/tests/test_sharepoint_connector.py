"""The SharePoint connector: which Graph resource it addresses, and with what.

Written against a real `httpx.AsyncClient` over a `MockTransport` rather than a
mocked client, because almost everything worth pinning here is a *URL*. A mock
that records `client.get(...)` calls proves the connector called something; the
transport proves what Graph would have received - which is the only level at
which "a folder called `x:/root/children` addresses a different resource" is a
statement about this code rather than about the assertion.

Three things this file exists to hold shut:

- **The bearer token reaches Graph and nothing else.** `/content` redirects to a
  pre-authenticated CDN URL, and a client that follows the redirect with the
  header still attached hands a Graph token to a host that never needed one.
- **Config is a path component, never a path.** Every value the wizard accepts
  is interpolated into a Graph URL, and `remote_names` is where that promotion
  is refused - here is where it is confirmed to actually be asked.
- **The connector's own type is a `Source`.** `sync_source_flow` stamps
  `source.connector_type` on every document it ingests, so a value missing from
  that vocabulary is a collection nobody can filter.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pytest

from app.core import secret_purposes
from app.core.exceptions import BadRequestError
from app.core.secret_kinds import ApiKeySecret, EntraAppSecret, SecretKind
from app.services.rag.connectors import CONNECTOR_REGISTRY
from app.services.rag.connectors.sharepoint import GRAPH, LOGIN, SharePointConnector
from app.services.rag.filters import SOURCE_VOCABULARY

pytestmark = pytest.mark.anyio

CONFIG: dict[str, Any] = {
    "hostname": "contoso.sharepoint.com",
    "site_path": "/sites/Engineering",
}

TOKEN_URL = f"{LOGIN}/a-tenant/oauth2/v2.0/token"


def _app() -> EntraAppSecret:
    return EntraAppSecret(
        tenant_id="a-tenant", client_id="an-application", client_secret="not-a-real-secret"
    )


def _file(name: str, item_id: str = "item-1", **extra: Any) -> dict[str, Any]:
    return {
        "id": item_id,
        "name": name,
        "file": {"mimeType": "application/pdf"},
        "size": 12,
        "lastModifiedDateTime": "2026-01-02T03:04:05Z",
        **extra,
    }


class Graph:
    """A tenant's Graph, as far as this connector can tell.

    `routes` maps a full URL to the JSON answered for it. Every request is
    recorded with its headers, which is what the token assertions read.
    """

    def __init__(self, routes: dict[str, Any], *, token_status: int = 200) -> None:
        self.routes = routes
        self.token_status = token_status
        self.seen: list[httpx.Request] = []

    def handler(self, request: httpx.Request) -> httpx.Response:
        self.seen.append(request)
        url = str(request.url)
        if url == TOKEN_URL:
            if self.token_status >= 400:
                return httpx.Response(
                    self.token_status,
                    json={"error_description": "AADSTS7000215: bad secret not-a-real-secret"},
                )
            return httpx.Response(200, json={"access_token": "a-graph-token", "expires_in": 3599})
        if url in self.routes:
            answer = self.routes[url]
            if isinstance(answer, httpx.Response):
                return answer
            return httpx.Response(200, json=answer)
        return httpx.Response(404, json={"error": {"code": "itemNotFound"}})

    def client(self) -> httpx.AsyncClient:
        return httpx.AsyncClient(
            transport=httpx.MockTransport(self.handler), follow_redirects=False
        )

    @property
    def urls(self) -> list[str]:
        return [str(request.url) for request in self.seen]

    def header_for(self, url: str) -> str | None:
        for request in self.seen:
            if str(request.url) == url:
                return request.headers.get("authorization")
        raise AssertionError(f"nothing requested {url}")


def _site(monkeypatch: pytest.MonkeyPatch, graph: Graph) -> SharePointConnector:
    monkeypatch.setattr(
        "app.services.rag.connectors.sharepoint.graph_client", graph.client, raising=True
    )
    return SharePointConnector()


SITE_URL = f"{GRAPH}/sites/contoso.sharepoint.com:/sites/Engineering"
DRIVE_URL = f"{GRAPH}/sites/site-1/drive"
ROOT_URL = f"{GRAPH}/drives/drive-1/root/children?$top=999"

BASE_ROUTES: dict[str, Any] = {
    SITE_URL: {"id": "site-1"},
    DRIVE_URL: {"id": "drive-1"},
}


class TestTheCredential:
    async def test_a_source_with_no_credential_is_refused_by_name(self) -> None:
        """Not a fallback to a deployment-wide registration: there is not one,
        and inventing one would read under the operator's identity (#937)."""
        with pytest.raises(BadRequestError, match="no credential"):
            await SharePointConnector().list_files(CONFIG, None)

    async def test_a_credential_of_another_kind_is_refused(self) -> None:
        with pytest.raises(BadRequestError, match="Entra app credential"):
            await SharePointConnector().list_files(CONFIG, ApiKeySecret(api_key="not-an-app"))

    async def test_the_client_secret_never_reaches_the_refusal(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Entra's own `error_description` echoes the request it is about, and
        the request carries the secret - so the refusal is written here."""
        graph = Graph(BASE_ROUTES, token_status=401)
        connector = _site(monkeypatch, graph)

        with pytest.raises(BadRequestError) as refusal:
            await connector.list_files(CONFIG, _app())

        assert "not-a-real-secret" not in str(refusal.value)
        assert "AADSTS" not in str(refusal.value)
        assert "HTTP 401" in refusal.value.message

    async def test_the_token_is_minted_once_for_a_whole_sync(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """One instance serves one sync. A token per file would be one round
        trip to Entra per document for a credential that has not changed."""
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("a.pdf"), _file("b.pdf", "item-2")]},
                f"{GRAPH}/drives/drive-1/items/item-1/content": httpx.Response(200, content=b"pdf"),
                f"{GRAPH}/drives/drive-1/items/item-2/content": httpx.Response(200, content=b"pdf"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        for remote in files:
            await connector.download_file(remote, tmp_path, CONFIG, _app())

        assert graph.urls.count(TOKEN_URL) == 1

    async def test_the_site_is_resolved_once_for_a_whole_sync(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """`_fetch` is handed a file and a destination, never the drive - so
        without the memo each download would re-resolve site and library."""
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("a.pdf")]},
                f"{GRAPH}/drives/drive-1/items/item-1/content": httpx.Response(200, content=b"pdf"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        await connector.download_file(files[0], tmp_path, CONFIG, _app())

        assert graph.urls.count(SITE_URL) == 1
        assert graph.urls.count(DRIVE_URL) == 1


class TestWhatItAddresses:
    async def test_the_site_and_its_default_library_are_resolved_in_order(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        graph = Graph({**BASE_ROUTES, ROOT_URL: {"value": []}})
        connector = _site(monkeypatch, graph)

        await connector.list_files(CONFIG, _app())

        assert graph.urls == [TOKEN_URL, SITE_URL, DRIVE_URL, ROOT_URL]

    async def test_a_folder_path_addresses_the_folder_and_is_encoded(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """ "Shared Documents" is what a document library is actually called, so
        the legal values need encoding - the allowlist is what makes that safe
        rather than a way of smuggling a delimiter back in."""
        folder_url = f"{GRAPH}/drives/drive-1/root:/Shared%20Documents/Legal:/children?$top=999"
        graph = Graph({**BASE_ROUTES, folder_url: {"value": []}})
        connector = _site(monkeypatch, graph)

        await connector.list_files({**CONFIG, "folder_path": "Shared Documents/Legal"}, _app())

        assert folder_url in graph.urls

    async def test_subfolders_are_walked_when_the_source_asked_for_them(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        child_url = f"{GRAPH}/drives/drive-1/items/folder-1/children?$top=999"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [{"id": "folder-1", "name": "Legal", "folder": {}}]},
                child_url: {"value": [_file("contract.pdf")]},
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())

        assert [f.name for f in files] == ["contract.pdf"]

    async def test_subfolders_are_left_alone_when_it_did_not(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {
                    "value": [{"id": "folder-1", "name": "Legal", "folder": {}}, _file("a.pdf")]
                },
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files({**CONFIG, "include_subfolders": False}, _app())

        assert [f.name for f in files] == ["a.pdf"]
        assert f"{GRAPH}/drives/drive-1/items/folder-1/children?$top=999" not in graph.urls

    async def test_an_item_that_is_neither_a_file_nor_a_folder_is_skipped(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """A OneNote section or a bundle has no bytes behind `/content`, so
        listing one produces a document the pipeline fails to parse."""
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {
                    "value": [{"id": "note-1", "name": "Notebook", "package": {}}, _file("a.pdf")]
                },
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())

        assert [f.name for f in files] == ["a.pdf"]

    async def test_paging_follows_the_link_graph_built(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """The skip token and the ordering are Graph's, so the next page is
        followed as given rather than rebuilt from a page number."""
        second = f"{GRAPH}/drives/drive-1/root/children?$top=999&$skiptoken=xyz"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("a.pdf")], "@odata.nextLink": second},
                second: {"value": [_file("b.pdf", "item-2")]},
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())

        assert [f.name for f in files] == ["a.pdf", "b.pdf"]

    async def test_the_dedup_key_carries_the_drive(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """An item id is unique within its drive, not across them - two
        libraries each holding a file would otherwise share one key."""
        graph = Graph({**BASE_ROUTES, ROOT_URL: {"value": [_file("a.pdf")]}})
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())

        assert files[0].source_path == "sharepoint://drive-1/item-1"

    async def test_a_listing_refusal_names_the_status_and_not_the_body(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """403 is "a permission never granted" and 404 is "that site does not
        resolve"; the status is the diagnosis and Graph's body is not ours."""
        graph = Graph({SITE_URL: httpx.Response(403, json={"error": {"message": "/secret/path"}})})
        connector = _site(monkeypatch, graph)

        with pytest.raises(BadRequestError) as refusal:
            await connector.list_files(CONFIG, _app())

        assert "HTTP 403" in refusal.value.message
        assert "/secret/path" not in refusal.value.message


class TestTheDownload:
    async def test_the_graph_token_is_not_handed_to_the_cdn(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """`/content` answers 302 to a pre-authenticated URL that needs no
        bearer. Whether a client strips the header across a redirect is the
        client's business; not sending it is this function's."""
        content = f"{GRAPH}/drives/drive-1/items/item-1/content"
        cdn = "https://contoso.sharepoint.com/_layouts/download.aspx?share=abc"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("report.pdf")]},
                content: httpx.Response(302, headers={"location": cdn}),
                cdn: httpx.Response(200, content=b"%PDF-1.7 payload"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        landed = await connector.download_file(files[0], tmp_path, CONFIG, _app())

        assert landed.read_bytes() == b"%PDF-1.7 payload"
        assert graph.header_for(content) == "Bearer a-graph-token"
        assert graph.header_for(cdn) is None

    async def test_a_body_served_without_a_redirect_is_still_written(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        content = f"{GRAPH}/drives/drive-1/items/item-1/content"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("report.pdf")]},
                content: httpx.Response(200, content=b"bytes"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        landed = await connector.download_file(files[0], tmp_path, CONFIG, _app())

        assert landed.read_bytes() == b"bytes"
        assert graph.urls.count(content) == 1

    async def test_a_traversing_remote_name_lands_inside_the_sync_directory(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """A file name is chosen by whoever can drop a file in the library,
        which on a shared site is not only the tenant. `download_file` is not
        overridden, so this is inherited - and that is the thing being pinned."""
        sync_dir = tmp_path / "sync"
        sync_dir.mkdir()
        content = f"{GRAPH}/drives/drive-1/items/item-1/content"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("../../evil.txt")]},
                content: httpx.Response(200, content=b"payload"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        landed = await connector.download_file(files[0], sync_dir, CONFIG, _app())

        assert landed == sync_dir / "evil.txt"
        assert not (tmp_path / "evil.txt").exists()

    async def test_a_relative_redirect_is_resolved_rather_than_raised(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """`httpx.InvalidURL` is not an `HTTPError`, so a bare relative
        `Location` would leave the sync log naming a class instead of a
        cause."""
        content = f"{GRAPH}/drives/drive-1/items/item-1/content"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("report.pdf")]},
                content: httpx.Response(302, headers={"location": "/download/abc"}),
                "https://graph.microsoft.com/download/abc": httpx.Response(200, content=b"bytes"),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        landed = await connector.download_file(files[0], tmp_path, CONFIG, _app())

        assert landed.read_bytes() == b"bytes"

    async def test_a_refused_download_names_the_file_and_the_status(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        content = f"{GRAPH}/drives/drive-1/items/item-1/content"
        graph = Graph(
            {
                **BASE_ROUTES,
                ROOT_URL: {"value": [_file("report.pdf")]},
                content: httpx.Response(423, json={"error": {"message": "locked"}}),
            }
        )
        connector = _site(monkeypatch, graph)

        files = await connector.list_files(CONFIG, _app())
        with pytest.raises(BadRequestError) as refusal:
            await connector.download_file(files[0], tmp_path, CONFIG, _app())

        assert "report.pdf" in refusal.value.message
        assert "HTTP 423" in refusal.value.message


class TestTheConfigTheWizardPosted:
    @pytest.mark.parametrize("missing", ["hostname", "site_path"])
    async def test_a_required_field_is_refused_by_the_name_it_was_drawn_under(
        self, missing: str
    ) -> None:
        config = {key: value for key, value in CONFIG.items() if key != missing}

        refusal = await SharePointConnector().validate_config(config)

        assert refusal is not None
        assert refusal.field == missing

    @pytest.mark.parametrize(
        ("field", "value"),
        [
            ("hostname", "contoso.sharepoint.com/sites/Other"),
            ("hostname", "https://contoso.sharepoint.com"),
            ("site_path", "/sites/Engineering:/drive/root:"),
            ("site_path", "/sites/../../users"),
            ("folder_path", "Legal?$expand=children"),
            ("folder_path", "Legal/../../other"),
        ],
    )
    async def test_a_value_that_would_address_something_else_is_refused_by_field(
        self, field: str, value: str
    ) -> None:
        """Answered by the route that accepted it rather than by a sync log an
        hour later - and by the same helper the URL builder asks, so the two
        cannot come to different conclusions."""
        refusal = await SharePointConnector().validate_config({**CONFIG, field: value})

        assert refusal is not None
        assert refusal.field == field

    async def test_a_config_that_names_a_real_site_is_accepted(self) -> None:
        assert await SharePointConnector().validate_config(CONFIG) is None
        assert (
            await SharePointConnector().validate_config(
                {**CONFIG, "folder_path": "Shared Documents/Legal", "include_subfolders": False}
            )
            is None
        )

    async def test_a_hostile_value_is_refused_at_sync_time_too(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """A row stored before the check existed, or edited underneath it,
        reaches `list_files` without passing `validate_config` again."""
        graph = Graph(BASE_ROUTES)
        connector = _site(monkeypatch, graph)

        with pytest.raises(BadRequestError):
            await connector.list_files({**CONFIG, "hostname": "contoso.sharepoint.com/x"}, _app())

        assert SITE_URL not in graph.urls

    async def test_a_hostile_folder_is_refused_before_a_token_is_minted(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """The folder is the last value read and the first one that could be
        checked late - which would have spent a token on a config the route
        would have refused."""
        graph = Graph(BASE_ROUTES)
        connector = _site(monkeypatch, graph)

        with pytest.raises(BadRequestError):
            await connector.list_files({**CONFIG, "folder_path": "Legal/../../other"}, _app())

        assert graph.urls == []


class TestTheRegistration:
    def test_the_connector_is_registered_under_its_own_type(self) -> None:
        assert CONNECTOR_REGISTRY["sharepoint"] is SharePointConnector
        assert SharePointConnector.CONNECTOR_TYPE == "sharepoint"

    def test_its_type_is_a_source_documents_can_be_filtered_by(self) -> None:
        """`sync_source_flow` stamps `source.connector_type` verbatim, and
        `RetrievalFilters` refuses a value outside the vocabulary - so a
        connector missing here ingests documents nobody can filter for."""
        assert SharePointConnector.CONNECTOR_TYPE in SOURCE_VOCABULARY

    def test_the_credential_it_needs_can_be_stored_for_it(self) -> None:
        """The vault picker narrows on a purpose, not on a kind alone. Without
        an entry of the right kind there is nothing to point a source at."""
        entry = secret_purposes.get("sharepoint")
        assert entry is not None
        assert entry.kind is SharePointConnector.SECRET_KIND
        assert entry.category is secret_purposes.PurposeCategory.CONNECTOR

    def test_its_config_model_holds_no_credential(self) -> None:
        """A field for a token here is a credential in a JSONB column, which
        is what `0042_sync_source_secret_id` removed (#937)."""
        fields = set(SharePointConnector.CONFIG_MODEL.model_fields)
        assert fields == {"hostname", "site_path", "folder_path", "include_subfolders"}
        # The tenant and the application id belong with the secret that needs
        # them, not beside the site: they are what the token endpoint is built
        # from, and splitting a credential across two stores is how half of one
        # gets rotated.
        assert not fields & {"tenant_id", "client_id", "client_secret", "token", "password"}

    async def test_a_client_secret_pasted_into_the_config_is_refused(self) -> None:
        """Every Entra walkthrough outside this repository puts the tenant, the
        application id and the secret in one place, so this is the shape a
        person will try - and `config` is a dict that accepts a key nobody
        declared, so what it would otherwise do is store a plaintext client
        secret in a JSONB column (#937)."""
        from app.services.sync_source import _refuse_a_credential_in_the_config

        with pytest.raises(BadRequestError, match="does not go in a source"):
            await _refuse_a_credential_in_the_config(
                {**CONFIG, "client_secret": "pasted-from-the-portal"}, "sharepoint"
            )

    def test_every_registered_connector_declares_a_storable_kind(self) -> None:
        for connector in CONNECTOR_REGISTRY.values():
            assert connector.SECRET_KIND is not SecretKind.NONE
            assert secret_purposes.get(connector.CONNECTOR_TYPE) is not None
