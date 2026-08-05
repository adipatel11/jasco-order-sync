"""Microsoft Graph access to the staging workbook in the owner's OneDrive.

The staging workbook used to live on whichever machine ran the sync — which stopped
working when the daily job moved to a VPS the owner can't reach. Keeping it in OneDrive
puts one copy somewhere every machine can write and the owner can open from a bookmark.

It stays a real .xlsx and is edited by download → openpyxl → upload, NOT through the
Excel REST API. That's deliberate: the Excel API doesn't support app-only auth, and its
v1.0 reference states it doesn't cover consumer OneDrive at all. Round-tripping the file
through the DriveItem content endpoints works identically on personal and business
accounts, and leaves xlsx_writer's formatting logic untouched.

Why delegated auth rather than a client secret: Graph has no app-only route for
per-user OneDrive files here, so the owner signs in ONCE via device code and the
refresh token is cached on disk. MSAL renews it on every silent acquire, so a job that
runs daily keeps it alive indefinitely; it dies only if the password changes or a
session is revoked, which surfaces as GraphAuthRequiredError. That mirrors the
trusted-device pattern the TAP scraper already uses: one interactive bootstrap, then
unattended forever, with a loud, actionable error if it ever lapses.
"""

from __future__ import annotations

import base64
import logging
import os
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

# Files.ReadWrite covers reading and replacing the file's content. offline_access is
# what earns us a refresh token; MSAL appends it itself and rejects it as a reserved
# scope if passed explicitly.
SCOPES = ["Files.ReadWrite"]

# Graph's simple upload endpoint is documented for payloads up to 4 MB; anything
# larger has to go through a chunked upload session.
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
# Chunks must be a multiple of 320 KiB. 10 × 320 KiB keeps the request count low
# without holding much in memory.
CHUNK_SIZE = 10 * 320 * 1024

RETRY_STATUS = {429, 500, 502, 503, 504}
MAX_ATTEMPTS = 5

log = logging.getLogger("onedrive")


class GraphAuthRequiredError(RuntimeError):
    """No usable cached token — a human must complete the device-code sign-in."""


class WorkbookChangedError(RuntimeError):
    """The cloud copy changed while we were editing it; our upload was refused."""


REAUTH_MESSAGE = (
    "OneDrive sign-in has expired.\n\n"
    "Run `python onedrive.py login` in a terminal on this machine and complete the "
    "sign-in it shows. It only needs doing once — after that the token renews itself."
)


# --- auth -------------------------------------------------------------------

def _cache_path() -> Path:
    return Path(
        os.environ.get("GRAPH_TOKEN_CACHE", Path(__file__).parent / ".graph_token_cache.json")
    ).expanduser()


def _build_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    client_id = os.environ.get("GRAPH_CLIENT_ID")
    if not client_id:
        raise RuntimeError("GRAPH_CLIENT_ID is not set in .env")

    # 'consumers' for a personal Microsoft account, a tenant GUID to pin one org, or
    # 'common' (the default) to accept either — one app registration serves all three.
    authority = os.environ.get("GRAPH_AUTHORITY", "https://login.microsoftonline.com/common")

    cache = msal.SerializableTokenCache()
    path = _cache_path()
    if path.exists():
        cache.deserialize(path.read_text())
    return msal.PublicClientApplication(client_id, authority=authority, token_cache=cache), cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if not cache.has_state_changed:
        return
    path = _cache_path()
    path.write_text(cache.serialize())
    # The refresh token is a bearer credential for the owner's OneDrive — keep it
    # unreadable to other accounts on a shared VPS.
    path.chmod(0o600)


def get_token() -> str:
    """Return an access token from the cache, refreshing silently if needed.

    Never prompts: an unattended run has no terminal to prompt at, and a Tk picker
    has no console to show a device code in. Raises GraphAuthRequiredError instead so
    callers can surface REAUTH_MESSAGE in whatever way suits them.
    """
    app, cache = _build_app()
    accounts = app.get_accounts()
    if not accounts:
        raise GraphAuthRequiredError("No cached OneDrive sign-in on this machine")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)
    if not result or "access_token" not in result:
        raise GraphAuthRequiredError("The cached OneDrive sign-in could not be refreshed")
    return result["access_token"]


def device_code_login() -> None:
    """One-time interactive sign-in, printing a short code to enter on another device.

    Device code rather than a localhost redirect because this bootstrap usually happens
    over SSH on a headless VPS, where no browser opens and no redirect can be reached.
    """
    app, cache = _build_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start sign-in: {flow.get('error_description')}")

    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)  # blocks until signed in or timeout
    _save_cache(cache)
    if "access_token" not in result:
        raise RuntimeError(f"Sign-in failed: {result.get('error_description')}")


# --- transport --------------------------------------------------------------

def request(method: str, url: str, **kwargs) -> requests.Response:
    """Authenticated Graph call that retries throttling and transient backend errors.

    Graph throttles per-drive and occasionally 503s mid-upload; honouring Retry-After
    is the difference between a reliable append and a half-written workbook.
    """
    if not url.startswith("http"):
        url = GRAPH + url

    resp = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        headers = dict(kwargs.pop("headers", {}) or {})
        headers.setdefault("Authorization", f"Bearer {get_token()}")
        resp = requests.request(method, url, headers=headers, timeout=120, **kwargs)

        if resp.status_code not in RETRY_STATUS or attempt == MAX_ATTEMPTS:
            return resp

        delay = float(resp.headers.get("Retry-After", 2 ** attempt))
        log.warning(
            "Graph %s → %d; retrying in %.0fs (attempt %d/%d)",
            url, resp.status_code, delay, attempt, MAX_ATTEMPTS,
        )
        time.sleep(delay)
    return resp


def _fail(resp: requests.Response, context: str) -> None:
    try:
        detail = resp.json()["error"]["message"]
    except Exception:  # noqa: BLE001 — Graph error bodies aren't guaranteed to be JSON
        detail = resp.text[:400]
    raise RuntimeError(f"{context} failed ({resp.status_code}): {detail}")


# --- the workbook -----------------------------------------------------------

@dataclass
class WorkbookRef:
    """Everything needed to address the staging workbook, plus the owner's link."""

    drive_id: str
    item_id: str
    name: str
    web_url: str  # what the owner bookmarks — opens in Excel Online

    @property
    def base(self) -> str:
        return f"/drives/{self.drive_id}/items/{self.item_id}"


def _encode_share_url(url: str) -> str:
    """Turn a OneDrive share link into the /shares/{id} token Graph expects."""
    b64 = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    return "u!" + b64


def resolve() -> WorkbookRef:
    """Locate the staging workbook from .env.

    Two ways to point at it, because they suit different people:
      ONEDRIVE_SHARE_LINK — the owner clicks Share → Copy link and pastes it, with no
                            need to know or spell out a folder path.
      ONEDRIVE_FILE_PATH  — a path relative to the drive root, for scripted setup.
    """
    share_link = os.environ.get("ONEDRIVE_SHARE_LINK")
    file_path = os.environ.get("ONEDRIVE_FILE_PATH")
    select = "?$select=id,name,webUrl,eTag,size,parentReference"

    if share_link:
        resp = request("GET", f"/shares/{_encode_share_url(share_link)}/driveItem{select}")
        context = "Resolving ONEDRIVE_SHARE_LINK"
    elif file_path:
        resp = request("GET", f"/me/drive/root:/{file_path.lstrip('/')}:{select}")
        context = f"Resolving ONEDRIVE_FILE_PATH ({file_path})"
    else:
        raise RuntimeError("Set ONEDRIVE_SHARE_LINK or ONEDRIVE_FILE_PATH in .env")

    if not resp.ok:
        _fail(resp, context)
    item = resp.json()
    drive_id = item.get("parentReference", {}).get("driveId")
    if not drive_id:
        raise RuntimeError(f"Graph returned no driveId for {item.get('name')!r}")
    return WorkbookRef(drive_id, item["id"], item["name"], item.get("webUrl", ""))


def download(ref: WorkbookRef, dest: Path) -> str:
    """Fetch the current cloud copy to dest. Returns its eTag for the later upload.

    The eTag is the concurrency guard: run.py on the VPS and pick.py on a desktop are
    on different machines, so the repo's flock can't keep them apart. Passing this
    value back to upload() turns a lost update into a detectable failure.
    """
    meta = request("GET", f"{ref.base}?$select=eTag,size")
    if not meta.ok:
        _fail(meta, f"Reading {ref.name}")
    etag = meta.json()["eTag"]

    resp = request("GET", f"{ref.base}/content", stream=True)
    if not resp.ok:
        _fail(resp, f"Downloading {ref.name}")

    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as fh:
        for chunk in resp.iter_content(chunk_size=1024 * 256):
            fh.write(chunk)
    log.info("Downloaded %s (%d bytes) from OneDrive", ref.name, dest.stat().st_size)
    return etag


def upload(ref: WorkbookRef, src: Path, if_match: str | None = None) -> None:
    """Replace the cloud copy with src, refusing if it changed since if_match."""
    size = src.stat().st_size
    if size < SIMPLE_UPLOAD_LIMIT:
        _upload_simple(ref, src, if_match)
    else:
        _upload_chunked(ref, src, if_match)
    log.info("Uploaded %s (%d bytes) to OneDrive", ref.name, size)


def _upload_simple(ref: WorkbookRef, src: Path, if_match: str | None) -> None:
    headers = {"Content-Type": "application/octet-stream"}
    if if_match:
        headers["if-match"] = if_match
    resp = request("PUT", f"{ref.base}/content", headers=headers, data=src.read_bytes())
    if resp.status_code == 412:
        raise WorkbookChangedError(f"{ref.name} changed in OneDrive while we were editing it")
    if not resp.ok:
        _fail(resp, f"Uploading {ref.name}")


def _upload_chunked(ref: WorkbookRef, src: Path, if_match: str | None) -> None:
    """Chunked upload for workbooks past the simple-upload limit.

    createUploadSession takes no if-match, so the best available guard is to re-check
    the eTag immediately beforehand. That leaves a small window a simple PUT doesn't
    have — acceptable because only this program writes the staging workbook, and the
    caller retries on WorkbookChangedError anyway.
    """
    if if_match:
        meta = request("GET", f"{ref.base}?$select=eTag")
        if meta.ok and meta.json().get("eTag") != if_match:
            raise WorkbookChangedError(f"{ref.name} changed in OneDrive before the upload began")

    resp = request(
        "POST",
        f"{ref.base}/createUploadSession",
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
    )
    if not resp.ok:
        _fail(resp, f"Starting an upload session for {ref.name}")
    session_url = resp.json()["uploadUrl"]

    size = src.stat().st_size
    with open(src, "rb") as fh:
        start = 0
        while start < size:
            chunk = fh.read(CHUNK_SIZE)
            end = start + len(chunk) - 1
            # The session URL carries its own credential, so no Authorization header.
            put = requests.put(
                session_url,
                headers={"Content-Range": f"bytes {start}-{end}/{size}"},
                data=chunk,
                timeout=120,
            )
            if put.status_code not in (200, 201, 202):
                requests.delete(session_url, timeout=30)  # don't leave a stale session
                _fail(put, f"Uploading a chunk of {ref.name}")
            start = end + 1


# --- CLI --------------------------------------------------------------------

def main() -> int:
    from dotenv import load_dotenv

    load_dotenv(Path(__file__).parent / ".env")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    cmd = sys.argv[1] if len(sys.argv) > 1 else "info"
    try:
        if cmd == "login":
            device_code_login()
            print("\nSigned in. Unattended runs on this machine need no further sign-in.")
            return 0
        if cmd == "info":
            ref = resolve()
            print(f"Workbook : {ref.name}")
            print(f"Size     : {request('GET', ref.base + '?$select=size').json().get('size')} bytes")
            print(f"\nBookmark this link for the owner:\n{ref.web_url}")
            return 0
    except GraphAuthRequiredError:
        print(REAUTH_MESSAGE, file=sys.stderr)
        return 2

    print(f"usage: python onedrive.py [login|info]  (got {cmd!r})", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
