#!/usr/bin/env python3
"""Shared configuration, logging, auth, and persistence for Claudette-Teams.

Platform-agnostic helpers lifted from the Slack bot live here so that bot.py,
claude_session.py, and teams_files.py can share them without circular imports.
"""

from __future__ import annotations

import json
import logging
import logging.handlers
import os
import re
import threading
from pathlib import Path

from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_DIR = Path(__file__).parent

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
for _noisy in ("aiohttp", "httpx", "httpcore", "msal", "urllib3"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)

# The Agents SDK logs under this namespace; surface it at INFO.
logging.getLogger("microsoft_agents").setLevel(logging.INFO)

logger = logging.getLogger("claudette")

_rotating_handler = logging.handlers.RotatingFileHandler(
    LOG_DIR / "bot.log", maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_rotating_handler.setFormatter(
    logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
)
logger.addHandler(_rotating_handler)

audit_handler = logging.FileHandler(LOG_DIR / "audit.log", encoding="utf-8")
audit_handler.setFormatter(
    logging.Formatter("%(asctime)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
)
audit_logger = logging.getLogger("claudette.audit")
audit_logger.addHandler(audit_handler)
audit_logger.setLevel(logging.INFO)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

load_dotenv(override=True)

# Azure Bot / app registration credentials (Agents SDK reads these directly from
# the environment too; we read CLIENTID here for proactive messaging).
CLIENT_ID = os.environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID", "")
CLIENT_SECRET = os.environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET", "")
TENANT_ID = os.environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID", "")

PROJECT_DIR = os.environ.get("PROJECT_DIR", "")
if not PROJECT_DIR:
    logger.error("PROJECT_DIR not set. Add it to .env")
    raise SystemExit(1)

PORT = int(os.environ.get("PORT", "3978"))

CLAUDE_TIMEOUT = int(os.environ.get("CLAUDE_TIMEOUT", "7200"))
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-4-6[1m]")
CLAUDE_EFFORT = os.environ.get("CLAUDE_EFFORT", "medium")

AUTHORIZED_USERS = set(
    u.strip() for u in os.environ.get("AUTHORIZED_USERS", "").split(",") if u.strip()
)
SUPERVISOR_USERS = set(
    u.strip() for u in os.environ.get("SUPERVISOR_USERS", "").split(",") if u.strip()
) or AUTHORIZED_USERS
RESTRICTED_USERS = set(
    u.strip() for u in os.environ.get("RESTRICTED_USERS", "").split(",") if u.strip()
)
RESTRICTED_ALLOWED_TOOLS = [
    t.strip() for t in os.environ.get(
        "RESTRICTED_ALLOWED_TOOLS",
        "Read,Edit,Write,Grep,Glob,Bash,WebSearch,WebFetch,Agent",
    ).split(",") if t.strip()
]
RESTRICTED_DISALLOWED_TOOLS = [
    t.strip() for t in os.environ.get("RESTRICTED_DISALLOWED_TOOLS", "").split(",") if t.strip()
]

ALLOWED_CHANNEL_PREFIXES = tuple(
    p.strip() for p in os.environ.get("ALLOWED_CHANNELS", "").split(",") if p.strip()
)

TRUST_BATTERY_DIR = os.environ.get("TRUST_BATTERY_DIR", "")
BOT_DISPLAY_NAME = os.environ.get("BOT_DISPLAY_NAME", "Claudette")
SHAREPOINT_SITE_ID = os.environ.get("SHAREPOINT_SITE_ID", "")

# Teams accepts ~28 KB per message; stay well under to allow markdown overhead.
MAX_TEAMS_MSG_LEN = 17000

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


def is_authorized(user_id: str) -> bool:
    return not AUTHORIZED_USERS or user_id in AUTHORIZED_USERS


def log_unauthorized(user_id: str, conversation_id: str, text: str) -> None:
    audit_logger.warning(
        f'UNAUTHORIZED | USER:{user_id} | CONV:{conversation_id} | MSG:"{text[:100]}"'
    )


def audit_interaction(
    user_id: str, conversation_id: str, response_text: str, duration: float, session_id: str | None
) -> None:
    audit_logger.info(
        f"USER:{user_id} | CONV:{conversation_id} | SESSION:{session_id or 'new'} "
        f"| DURATION:{duration:.1f}s | RESP_LEN:{len(response_text)}"
    )


# ---------------------------------------------------------------------------
# Session store: conversation_id -> Claude session_id (file-backed)
# ---------------------------------------------------------------------------

SESSION_FILE = LOG_DIR / ".sessions.json"
MAX_SESSIONS = 200
_session_file_lock = threading.Lock()


def _load_sessions() -> dict:
    try:
        return json.loads(SESSION_FILE.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_session(conversation_id: str, session_id: str) -> None:
    with _session_file_lock:
        sessions = _load_sessions()
        sessions[conversation_id] = session_id
        if len(sessions) > MAX_SESSIONS:
            for key in sorted(sessions.keys())[:-MAX_SESSIONS]:
                del sessions[key]
        SESSION_FILE.write_text(json.dumps(sessions))


def get_session(conversation_id: str) -> str | None:
    return _load_sessions().get(conversation_id)


# ---------------------------------------------------------------------------
# Conversation-reference store: conversation_id -> serialized ConversationReference
#
# Teams cannot cold-message a user; you can only reach a conversation you've
# already received a message from. We persist the reference from every inbound
# activity so background tasks and CLI proactive modes can push messages later.
# ---------------------------------------------------------------------------

REFERENCE_FILE = LOG_DIR / ".references.json"
MAX_REFERENCES = 500
_reference_file_lock = threading.Lock()


def _load_references() -> dict:
    try:
        return json.loads(REFERENCE_FILE.read_text())
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_reference(conversation_id: str, reference: dict) -> None:
    with _reference_file_lock:
        refs = _load_references()
        refs[conversation_id] = reference
        if len(refs) > MAX_REFERENCES:
            for key in list(refs.keys())[:-MAX_REFERENCES]:
                del refs[key]
        REFERENCE_FILE.write_text(json.dumps(refs))


def get_reference(conversation_id: str) -> dict | None:
    return _load_references().get(conversation_id)


# ---------------------------------------------------------------------------
# Trust battery (unchanged from Slack bot)
# ---------------------------------------------------------------------------


def get_trust_battery_context() -> str:
    if not TRUST_BATTERY_DIR:
        return ""
    battery_dir = Path(TRUST_BATTERY_DIR)
    if not battery_dir.exists():
        return ""
    tiers = [
        (0, 25, "Propose and Wait"),
        (25, 50, "Routine Execution"),
        (50, 75, "Judgment Calls"),
        (75, 100, "Full Autonomy"),
    ]
    lines = ["## Trust Battery — Current State"]
    for fpath in sorted(battery_dir.glob("*.json")):
        try:
            data = json.loads(fpath.read_text())
            name = data.get("team_member", fpath.stem)
            charge = data.get("current_charge", 0)
            last_updated = data.get("last_updated", "unknown")
            last_delta = 0.0
            if data.get("history"):
                last_delta = data["history"][-1].get("delta", 0.0)
            tier = next((t for lo, hi, t in tiers if lo <= charge < hi), "Full Autonomy")
            sign = "+" if last_delta >= 0 else ""
            lines.append(f"- {name}: {charge:.1f}% ({tier}) | Last: {sign}{last_delta:.1f} on {last_updated}")
        except Exception:
            continue
    if len(lines) == 1:
        return ""
    lines += [
        "",
        "Your autonomy level is determined by the battery charge for the",
        "team member you're interacting with:",
        "  0-25%  = Propose and Wait",
        "  25-50% = Routine Execution",
        "  50-75% = Judgment Calls",
        "  75-100% = Full Autonomy",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Markdown + chunking for Teams
# ---------------------------------------------------------------------------


def teams_markdown(text: str) -> str:
    """Teams renders standard markdown (bold, links, lists, code) but NOT ATX
    headings. Convert `# heading` lines to bold; leave the rest untouched."""
    return re.sub(r"^#{1,6}\s+(.+)$", r"**\1**", text, flags=re.MULTILINE)


def chunk_message(text: str) -> list[str]:
    """Split a message into Teams-safe chunks, preferring newline boundaries."""
    if len(text) <= MAX_TEAMS_MSG_LEN:
        return [text]
    chunks = []
    while text:
        if len(text) <= MAX_TEAMS_MSG_LEN:
            chunks.append(text)
            break
        split_at = text.rfind("\n", 0, MAX_TEAMS_MSG_LEN)
        if split_at == -1:
            split_at = text.rfind(" ", 0, MAX_TEAMS_MSG_LEN)
        if split_at == -1:
            split_at = MAX_TEAMS_MSG_LEN
        chunks.append(text[:split_at])
        text = text[split_at:].lstrip("\n")
    return chunks


# ---------------------------------------------------------------------------
# attach:<path> detection for outbound file uploads (unchanged regex)
# ---------------------------------------------------------------------------

_ATTACH_PATTERN = re.compile(
    r"attach:\s*("
    r"[A-Za-z]:[\\/][^\s`'\"<>|*?,]+\.\w+"
    r"|~/[^\s`'\"<>|*?,]+\.\w+"
    r"|/(?:Users|tmp|var|home)/[^\s`'\"<>|*?,]+\.\w+"
    r")",
    re.MULTILINE,
)


def extract_attach_paths(text: str) -> list[Path]:
    """Return existing local files referenced by `attach:<path>` markers."""
    seen: set[str] = set()
    out: list[Path] = []
    for match in _ATTACH_PATTERN.findall(text):
        fp_str = match.rstrip(".,;:!?)]`\"'")
        if fp_str.startswith("~"):
            fp_str = str(Path.home() / fp_str[2:])
        if fp_str in seen:
            continue
        seen.add(fp_str)
        fp = Path(fp_str)
        if fp.exists() and fp.is_file():
            out.append(fp)
    return out
