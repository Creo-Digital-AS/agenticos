"""What a sync source is not allowed to decide: where a file lands, and what is queried.

A remote file's name and a folder's identifier both arrive from outside the
deployment, and the name is not even the tenant's: sharing a Drive folder with
people outside the organization is what folder sharing is *for*, so whoever can
drop a file in it chooses what the next sync writes to disk. Both are labels
until something promotes one to a path component or to a fragment of a query,
and this module is the one place that promotion is refused.

Kept outside `connectors/` because `sources/` needs the same two answers and
must not import from its sibling.

**Neither refusal names a field**, and that is the point of the paragraph above:
what they refuse is not something a caller sent. A remote file's name is chosen
by whoever can drop a file in the folder, and both checks run inside a sync,
where the reader is a log rather than a form. The folder id *is* configured, but
`SyncSourceConnector.validate_config` answers `(False, message)` - the details
never reach the wire, and the route re-raises about the connector (#897). So
these carry no `details["fields"]`, which the one shape is for (#891).
"""

import re
from pathlib import Path
from urllib.parse import quote

from app.core.exceptions import BadRequestError

# Drive issues base64url identifiers. The upper bound is generous - the longest
# id Google has issued is well under 64 characters - and exists so a refusal
# reads as a refusal rather than as a regex walking a megabyte of config.
_DRIVE_ID = re.compile(r"[A-Za-z0-9_-]{1,256}")


def checked_drive_folder_id(folder_id: object) -> str:
    """Answer `folder_id` if Google could have issued it, and refuse it otherwise.

    The Drive query language wraps a parent id in single quotes, so an id
    carrying one closes the literal and everything after it is read as query:
    `x' in parents or name contains 'salary` is well-formed and lists whatever
    the credential can reach. An allowlist rather than an escape, because a real
    identifier needs nothing the allowlist withholds, and an escape leaves every
    future sink to remember what this one remembered.

    Takes an `object` because a source's config is `dict[str, object]` and JSON
    carries numbers and nested structures - the field arrives as whatever was
    posted, and a value that is not a string is refused here rather than
    stringified into one somewhere on the way.

    Raises:
        BadRequestError: the value is not a Drive identifier.
    """
    if not isinstance(folder_id, str) or not _DRIVE_ID.fullmatch(folder_id):
        raise BadRequestError(
            message="A Google Drive folder ID may contain only letters, digits, '-' and '_'."
        )
    return folder_id


def destination_within(directory: Path, remote_name: str) -> Path:
    """Answer where a file called `remote_name` may be written inside `directory`.

    `../../../../home/app/.ssh/authorized_keys` is a legal Drive file name, so
    `directory / remote_name` writes wherever the worker's uid can reach - and
    the sync then ingests from there. The name is reduced to its final
    component, and the result is *resolved and confirmed* to be a child of
    `directory` rather than cleaned of the spellings we happened to think of:
    `..`, its percent-encodings, its unicode lookalikes and a symlink already
    sitting in the directory are one question after `resolve()`, and an
    enumeration of separators is a list that is always one entry short.

    A name that is no component at all - `..`, `.`, `/`, the empty string -
    resolves onto the directory itself and is refused rather than silently
    renamed, because a file with no name is nothing this pipeline can ingest.

    Raises:
        BadRequestError: the name does not name a file inside `directory`.
    """
    if "\x00" in remote_name:
        raise BadRequestError(message="A remote file name may not contain a NULL byte.")

    base = directory.resolve()
    destination = (base / Path(remote_name).name).resolve()
    if destination.parent != base:
        raise BadRequestError(
            message="A remote file name must name one file inside the sync directory."
        )
    return destination


# A DNS name, and nothing that could end the URL path it is interpolated into.
# The bound is DNS's own; a name longer than this is not one Microsoft issued.
_SHAREPOINT_HOST = re.compile(r"[A-Za-z0-9][A-Za-z0-9.-]{0,252}[A-Za-z0-9]")

# What may appear in one component of a SharePoint path. Deliberately an
# allowlist of "printable, not a delimiter": Graph's path addressing reads `:`
# as the end of a path and the start of a resource, `?` as the start of
# `$select` and friends, and `#` as a fragment - so any of the three inside a
# component silently changes which Graph call is made rather than which folder
# is read. `%` is excluded too, because the component is percent-encoded on the
# way out and a literal one would arrive double-encoded.
_SHAREPOINT_COMPONENT = re.compile(r"[^\x00-\x1f\x7f:?#%\\]+")


def checked_sharepoint_host(hostname: object) -> str:
    """Answer `hostname` if it is a host name, and refuse it otherwise.

    It is interpolated into a Graph URL - `/v1.0/sites/{host}:{site}` - where it
    is a *path component* rather than an address: every request this connector
    makes goes to `graph.microsoft.com` whatever this says, so the risk is not
    somewhere else being dialled but somewhere else being *addressed*. A value
    carrying `/` or `:` ends the component early and the remainder is read by
    Graph as more of the resource path.

    Takes an `object` for the reason `checked_drive_folder_id` does: a source's
    config is JSON, and a value that is not a string is refused here rather than
    stringified into one on the way.

    Raises:
        BadRequestError: the value is not a host name.
    """
    if not isinstance(hostname, str) or not _SHAREPOINT_HOST.fullmatch(hostname):
        raise BadRequestError(
            message=(
                "A SharePoint hostname is a host name alone, such as "
                "contoso.sharepoint.com - no scheme, port or path."
            )
        )
    if ".." in hostname:
        raise BadRequestError(message="A SharePoint hostname may not contain '..'.")
    return hostname


def checked_sharepoint_path(path: object, *, what: str) -> str:
    """Answer `path` percent-encoded for a Graph URL, or refuse it.

    A site path (`/sites/Engineering`) and a folder path (`Shared Documents/Legal`)
    are the same question twice, so they are the same function: split on `/`,
    refuse a component that is empty, `.`, `..` or carries a Graph delimiter, and
    encode each survivor with nothing left safe. The answer has no leading or
    trailing slash - the caller owns the delimiters around it, which is what
    keeps `/sites//x` and `/sites/x/` from being two ways to spell one site.

    Encoding is part of the answer rather than the caller's next step because
    the legal values need it: a document library is called "Shared Documents",
    with a space, and a caller who validated and then interpolated raw would
    build a URL Graph answers 400 to. The allowlist above is the control; the
    encoding is what makes the allowlist affordable.

    Args:
        path: The value as it was configured.
        what: What to call it in a refusal - "site path", "folder path".

    Raises:
        BadRequestError: the value is not a path, or names a component that
            would change which Graph resource is addressed.
    """
    if not isinstance(path, str):
        raise BadRequestError(message=f"A SharePoint {what} must be text.")
    components = [component for component in path.split("/") if component]
    if not components:
        raise BadRequestError(message=f"A SharePoint {what} must name at least one folder.")
    for component in components:
        if component in {".", ".."} or not _SHAREPOINT_COMPONENT.fullmatch(component):
            raise BadRequestError(
                message=(
                    f"A SharePoint {what} may not contain '..', ':', '?', '#', '%' or a backslash."
                )
            )
    return "/".join(quote(component, safe="") for component in components)
