#!/usr/bin/env python3
"""Async live-session pool bridging Teams conversations to long-lived `claude`
CLI processes over stream-json I/O.

This is the async rewrite of the Slack bot's threaded LiveSession pool. One
`claude` process is kept alive per Teams conversation; messages are piped to its
stdin and the CLI queues them, matching terminal behavior. A per-session
asyncio task reads stdout and dispatches assistant text blocks to a callback.

Note: asyncio subprocesses require a loop with subprocess support. On Linux
(Azure App Service) the default loop is fine; on Windows the default Proactor
loop also supports subprocess pipes.
"""

from __future__ import annotations

import asyncio
import json
import tempfile
import time
from dataclasses import dataclass, field
from typing import Awaitable, Callable

from config import (
    CLAUDE_EFFORT,
    CLAUDE_MODEL,
    PROJECT_DIR,
    RESTRICTED_ALLOWED_TOOLS,
    RESTRICTED_DISALLOWED_TOOLS,
    RESTRICTED_USERS,
    SUPERVISOR_USERS,
    get_session,
    get_trust_battery_context,
    logger,
    save_session,
)

IDLE_TIMEOUT = 10800  # 3 hours
MAX_LIVE_SESSIONS = 5

OnText = Callable[[str], Awaitable[None]]


@dataclass
class LiveSession:
    proc: asyncio.subprocess.Process
    conversation_id: str
    session_id: str | None = None
    last_activity: float = field(default_factory=time.time)
    turn_lock: asyncio.Lock = field(default_factory=asyncio.Lock)
    stdin_lock: asyncio.Lock = field(default_factory=asyncio.Lock)
    on_text: OnText | None = field(default=None, repr=False)
    turn_done: asyncio.Event = field(default_factory=asyncio.Event)
    reader_task: asyncio.Task | None = field(default=None, repr=False)
    produced_text: bool = False


_live_sessions: dict[str, LiveSession] = {}
_live_sessions_lock = asyncio.Lock()


def _build_command(session_id: str | None, user_id: str) -> list[str]:
    cmd = [
        "claude",
        "-p", get_trust_battery_context(),
        "--input-format", "stream-json",
        "--output-format", "stream-json",
        "--verbose",
        "--model", CLAUDE_MODEL,
        "--effort", CLAUDE_EFFORT,
    ]
    if user_id in SUPERVISOR_USERS:
        cmd += ["--permission-mode", "bypassPermissions"]
    elif user_id in RESTRICTED_USERS:
        cmd += ["--permission-mode", "dontAsk"]
        if RESTRICTED_ALLOWED_TOOLS:
            cmd += ["--allowedTools", " ".join(RESTRICTED_ALLOWED_TOOLS)]
        if RESTRICTED_DISALLOWED_TOOLS:
            cmd += ["--disallowedTools", " ".join(RESTRICTED_DISALLOWED_TOOLS)]
    else:
        cmd += ["--permission-mode", "dontAsk"]
    if session_id:
        cmd += ["--resume", session_id]
    return cmd


async def _spawn(session_id: str | None, user_id: str) -> asyncio.subprocess.Process:
    cmd = _build_command(session_id, user_id)
    stderr_tmp = tempfile.NamedTemporaryFile(mode="w+", suffix=".stderr", delete=False)
    proc = await asyncio.create_subprocess_exec(
        *cmd,
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=stderr_tmp,
        cwd=PROJECT_DIR,
    )
    perm = "bypassPermissions" if user_id in SUPERVISOR_USERS else "dontAsk"
    logger.info(
        f"Spawned Claude pid={proc.pid} (resume={session_id or 'none'}, user={user_id}, perm={perm})"
    )
    return proc


async def _reader_loop(session: LiveSession) -> None:
    """Read stdout line-by-line; dispatch assistant text, capture session id."""
    try:
        while True:
            raw = await session.proc.stdout.readline()
            if not raw:
                break
            line = raw.decode("utf-8", "replace").strip()
            if not line:
                continue
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue

            msg_type = data.get("type")
            if msg_type == "system":
                sid = data.get("session_id")
                if sid:
                    session.session_id = sid
            elif msg_type == "assistant":
                for block in data.get("message", {}).get("content", []):
                    if isinstance(block, dict) and block.get("type") == "text":
                        text = block.get("text", "").strip()
                        if text and session.on_text:
                            session.produced_text = True
                            await session.on_text(text)
            elif msg_type == "result":
                sid = data.get("session_id")
                if sid:
                    session.session_id = sid
                    save_session(session.conversation_id, sid)
                session.turn_done.set()
    except Exception as e:
        logger.error(f"Reader loop error for conv {session.conversation_id}: {e}")
    finally:
        session.turn_done.set()
        logger.info(f"Reader loop ended for conv {session.conversation_id} (pid={session.proc.pid})")
        async with _live_sessions_lock:
            _live_sessions.pop(session.conversation_id, None)


async def _evict_oldest_locked() -> None:
    """Caller must hold _live_sessions_lock."""
    oldest_id = min(_live_sessions, key=lambda k: _live_sessions[k].last_activity)
    old = _live_sessions.pop(oldest_id)
    logger.info(f"Evicting idle session for conv {oldest_id} (pid={old.proc.pid})")
    try:
        old.proc.stdin.close()
        await asyncio.wait_for(old.proc.wait(), timeout=10)
    except Exception:
        try:
            old.proc.kill()
        except Exception:
            pass


async def _get_or_create(conversation_id: str, user_id: str) -> LiveSession:
    async with _live_sessions_lock:
        session = _live_sessions.get(conversation_id)
        if session and session.proc.returncode is None:
            session.last_activity = time.time()
            return session
        if len(_live_sessions) >= MAX_LIVE_SESSIONS:
            await _evict_oldest_locked()
        saved = get_session(conversation_id)
        proc = await _spawn(saved, user_id)
        session = LiveSession(proc=proc, conversation_id=conversation_id, session_id=saved)
        _live_sessions[conversation_id] = session
        session.reader_task = asyncio.create_task(_reader_loop(session))
        return session


async def _send(session: LiveSession, text: str) -> None:
    msg = json.dumps({
        "type": "user",
        "session_id": "",
        "message": {"role": "user", "content": text},
        "parent_tool_use_id": None,
    })
    async with session.stdin_lock:
        session.proc.stdin.write((msg + "\n").encode("utf-8"))
        await session.proc.stdin.drain()
    session.last_activity = time.time()


def has_live_process(conversation_id: str) -> bool:
    s = _live_sessions.get(conversation_id)
    return bool(s and s.proc.returncode is None)


def get_session_id(conversation_id: str) -> str | None:
    """Saved Claude session id for a conversation, if any (file-backed)."""
    return get_session(conversation_id)


async def run_turn(
    conversation_id: str, user_id: str, text: str, on_text: OnText, timeout: int
) -> dict:
    """Feed one user message to the conversation's Claude process and wait for the
    turn to complete. Returns {status: ok|timeout|died, session_id, produced_text}.
    `on_text` is awaited for each assistant text block as it streams."""
    session = await _get_or_create(conversation_id, user_id)
    async with session.turn_lock:
        session.on_text = on_text
        session.produced_text = False
        session.turn_done.clear()
        await _send(session, text)
        try:
            await asyncio.wait_for(session.turn_done.wait(), timeout=timeout)
        except asyncio.TimeoutError:
            return {"status": "timeout", "session_id": session.session_id,
                    "produced_text": session.produced_text}
        if not session.produced_text and session.proc.returncode is not None:
            return {"status": "died", "session_id": session.session_id,
                    "produced_text": False}
    return {"status": "ok", "session_id": session.session_id,
            "produced_text": session.produced_text}


async def cleanup_idle_sessions() -> None:
    """Background task: kill processes idle past IDLE_TIMEOUT."""
    while True:
        await asyncio.sleep(300)
        now = time.time()
        async with _live_sessions_lock:
            stale = [(cid, s) for cid, s in _live_sessions.items()
                     if now - s.last_activity > IDLE_TIMEOUT]
        for cid, session in stale:
            logger.info(f"Cleaning up idle session for conv {cid} (pid={session.proc.pid})")
            try:
                session.proc.stdin.close()
                await asyncio.wait_for(session.proc.wait(), timeout=15)
            except Exception:
                try:
                    session.proc.kill()
                except Exception:
                    pass
            if session.session_id:
                save_session(cid, session.session_id)
            async with _live_sessions_lock:
                _live_sessions.pop(cid, None)
