#!/usr/bin/env python3
"""Teams file handling: outbound uploads and inbound attachment downloads.

Three paths, mirroring the OpenClaw msteams model:
  * Personal (1:1) chats   -> FileConsentCard handshake (no Graph needed).
  * Channel / group chats   -> Microsoft Graph upload to SharePoint + share link.
  * Inbound attachments     -> download to a temp file for Claude to read.

VERIFY-AGAINST-SAMPLE: the exact attachment content-type constants and the
file-consent invoke payload shape are stable Teams protocol values, but the
Agents SDK class names for Attachment/Activity should be confirmed against
github.com/microsoft/Agents samples/python on first run.
"""

from __future__ import annotations

import tempfile
import uuid
from pathlib import Path

import httpx

from config import SHAREPOINT_SITE_ID, logger

try:
    from microsoft_agents.activity import Activity, Attachment
except Exception:  # pragma: no cover - import shape verified at runtime
    Activity = None
    Attachment = None

FILE_CONSENT_CARD = "application/vnd.microsoft.teams.card.file.consent"
FILE_INFO_CARD = "application/vnd.microsoft.teams.card.file.info"

# Maps a per-upload token -> local file path, set when we send a consent card and
# read back when the user accepts. In-process only; fine for a single instance.
_pending_uploads: dict[str, str] = {}

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# Graph app-only token (client credentials) for channel/group uploads
# ---------------------------------------------------------------------------


def _graph_token() -> str | None:
    """Acquire an app-only Graph token via MSAL client credentials.

    Requires the bot's app registration to have Sites.ReadWrite.All (application)
    granted with admin consent. Returns None if creds are missing or acquisition
    fails (caller falls back to posting the local path)."""
    from config import CLIENT_ID, CLIENT_SECRET, TENANT_ID

    if not (CLIENT_ID and CLIENT_SECRET and TENANT_ID):
        return None
    try:
        import msal

        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        return result.get("access_token")
    except Exception as e:
        logger.error(f"Graph token acquisition failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Outbound — personal chat: FileConsentCard flow
# ---------------------------------------------------------------------------


def build_file_consent_attachment(file_path: str) -> "Attachment | None":
    """Build a FileConsentCard attachment for a local file and register a pending
    upload token. Returns None if the file is missing."""
    path = Path(file_path)
    if not path.exists() or not path.is_file():
        logger.error(f"File not found for consent card: {file_path}")
        return None
    token = uuid.uuid4().hex
    _pending_uploads[token] = str(path)
    return Attachment(
        content_type=FILE_CONSENT_CARD,
        name=path.name,
        content={
            "description": f"Claudette wants to share {path.name}",
            "sizeInBytes": path.stat().st_size,
            "acceptContext": {"token": token},
            "declineContext": {"token": token},
        },
    )


async def handle_file_consent_accept(value: dict) -> "Attachment | None":
    """Handle a fileConsent/invoke accept. Uploads the pending file to the
    upload URL Teams provided and returns a FileInfoCard attachment to post,
    or None on failure."""
    context = value.get("acceptContext") or value.get("context") or {}
    token = context.get("token")
    upload_info = value.get("uploadInfo", {})
    upload_url = upload_info.get("uploadUrl")
    path_str = _pending_uploads.pop(token, None) if token else None

    if not (path_str and upload_url):
        logger.error("file consent accept missing token/uploadUrl")
        return None

    path = Path(path_str)
    data = path.read_bytes()
    size = len(data)
    try:
        async with httpx.AsyncClient(timeout=120) as client:
            resp = await client.put(
                upload_url,
                content=data,
                headers={
                    "Content-Length": str(size),
                    "Content-Range": f"bytes 0-{size - 1}/{size}",
                },
            )
            resp.raise_for_status()
    except Exception as e:
        logger.error(f"FileConsent upload PUT failed for {path.name}: {e}")
        return None

    logger.info(f"Uploaded {path.name} ({size} bytes) via FileConsent")
    return Attachment(
        content_type=FILE_INFO_CARD,
        name=path.name,
        content_url=upload_info.get("contentUrl"),
        content={
            "uniqueId": upload_info.get("uniqueId"),
            "fileType": path.suffix.lstrip("."),
        },
    )


def handle_file_consent_decline(value: dict) -> None:
    context = value.get("declineContext") or value.get("context") or {}
    token = context.get("token")
    if token:
        _pending_uploads.pop(token, None)


# ---------------------------------------------------------------------------
# Outbound — channel/group: Graph upload to SharePoint, return a share link
# ---------------------------------------------------------------------------


async def upload_to_sharepoint(file_path: str) -> str | None:
    """Upload a local file to the configured SharePoint site's drive and return a
    shareable link. Returns None if Graph isn't configured or the upload fails."""
    if not SHAREPOINT_SITE_ID:
        return None
    token = _graph_token()
    if not token:
        return None

    path = Path(file_path)
    if not path.exists():
        return None
    headers = {"Authorization": f"Bearer {token}"}
    try:
        async with httpx.AsyncClient(timeout=120) as client:
            # Simple upload (<4 MB). Larger files would need an upload session.
            put = await client.put(
                f"{_GRAPH_BASE}/sites/{SHAREPOINT_SITE_ID}/drive/root:/"
                f"Claudette/{path.name}:/content",
                content=path.read_bytes(),
                headers={**headers, "Content-Type": "application/octet-stream"},
            )
            put.raise_for_status()
            item_id = put.json()["id"]

            link = await client.post(
                f"{_GRAPH_BASE}/sites/{SHAREPOINT_SITE_ID}/drive/items/{item_id}/createLink",
                json={"type": "view", "scope": "organization"},
                headers=headers,
            )
            link.raise_for_status()
            return link.json()["link"]["webUrl"]
    except Exception as e:
        logger.error(f"SharePoint upload failed for {path.name}: {e}")
        return None


# ---------------------------------------------------------------------------
# Inbound — download attachments a user sent to a temp file
# ---------------------------------------------------------------------------


async def download_attachments(activity) -> list[Path]:
    """Download inbound attachments to temp files. Handles Teams file-download
    info cards (DMs) and direct content URLs. Channel attachments without Graph
    access arrive as HTML stubs and are skipped."""
    attachments = getattr(activity, "attachments", None) or []
    out: list[Path] = []
    graph_token = None

    async with httpx.AsyncClient(timeout=120, follow_redirects=True) as client:
        for att in attachments:
            content_type = getattr(att, "content_type", "") or ""
            name = getattr(att, "name", None) or "attachment"
            content = getattr(att, "content", None) or {}

            # Teams "file download info" card carries a direct downloadUrl.
            download_url = None
            if isinstance(content, dict):
                download_url = content.get("downloadUrl")
            if not download_url:
                # Inline images / generic attachments expose contentUrl.
                download_url = getattr(att, "content_url", None)
            if not download_url:
                continue

            headers = {}
            # Graph/SharePoint-hosted URLs need a bearer token.
            if "sharepoint.com" in download_url or "graph.microsoft.com" in download_url:
                graph_token = graph_token or _graph_token()
                if graph_token:
                    headers["Authorization"] = f"Bearer {graph_token}"

            try:
                resp = await client.get(download_url, headers=headers)
                resp.raise_for_status()
                suffix = Path(name).suffix or ".bin"
                tmp = tempfile.NamedTemporaryFile(suffix=suffix, prefix="teams-", delete=False)
                tmp.write(resp.content)
                tmp.close()
                out.append(Path(tmp.name))
                logger.info(f"Downloaded Teams attachment: {name} -> {tmp.name}")
            except Exception as e:
                logger.error(f"Failed to download attachment {name}: {e}")

    return out
