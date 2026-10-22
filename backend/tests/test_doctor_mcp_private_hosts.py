"""`doctor`'s answer to "which private MCP servers may a connection reach?".

`MCP_ALLOWED_PRIVATE_HOSTS` matches a connection's host exactly, so an entry
written as a wildcard, a range, a URL or with a port admits nothing, and the
first sign of it would otherwise be an SSRF refusal in somebody's dialog.
"""

import pytest

from app.commands.doctor import _mcp_private_hosts
from app.core.config import settings


def test_an_empty_list_is_reported_rather_than_failed(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(settings, "MCP_ALLOWED_PRIVATE_HOSTS", [" "])

    status, detail = _mcp_private_hosts()

    assert status == "unconfigured"
    assert "refused" in detail


def test_exact_hosts_and_addresses_are_listed(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        settings,
        "MCP_ALLOWED_PRIVATE_HOSTS",
        [" SP.taild84bfc.ts.net ", "100.72.29.28", "fd7a:115c:a1e0::1"],
    )

    status, detail = _mcp_private_hosts()

    assert status == "healthy"
    assert detail == "sp.taild84bfc.ts.net, 100.72.29.28, fd7a:115c:a1e0::1"


@pytest.mark.parametrize(
    "entry",
    [
        "*.taild84bfc.ts.net",
        "100.64.0.0/10",
        "https://sp.taild84bfc.ts.net",
        "sp.taild84bfc.ts.net:8443",
    ],
)
def test_an_entry_that_can_never_match_is_named_without_failing(
    monkeypatch: pytest.MonkeyPatch, entry: str
) -> None:
    monkeypatch.setattr(settings, "MCP_ALLOWED_PRIVATE_HOSTS", ["sp.taild84bfc.ts.net", entry])

    status, detail = _mcp_private_hosts()

    assert status == "not_checked"
    assert entry in detail
    assert "will never match" in detail
