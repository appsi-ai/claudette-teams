#!/usr/bin/env python3
"""Claudette-Teams — Microsoft Teams bot powered by Claude Code.

Port of the Slack `claudette` bot to the Microsoft 365 Agents SDK. Teams posts
activities to /api/messages; we ack fast (return from the handler immediately
after a typing indicator) and stream Claude's reply back PROACTIVELY via a stored
conversation reference, because Teams enforces a ~15s webhook deadline.

CLI proactive modes (require a previously stored conversation reference — Teams
cannot cold-message a user):
    python bot.py --send <conversation_id> "message"
    python bot.py --channel <conversation_id> "message"
    echo '{"result":"..."}' | python bot.py --send-result <conversation_id>

VERIFY-AGAINST-SAMPLE (github.com/microsoft/Agents samples/python): the exact
names `start_agent_process`, `jwt_authorization_middleware`, and the
`CloudAdapter.continue_conversation(reference, callback, bot_app_id)` signature.
The handler/config patterns below follow the official bf-migration-python guide.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import sys
import time
from os import environ
from pathlib import Path

from aiohttp.web import Application, Request, Response, json_response, run_app

from microsoft_agents.activity import Activity, ConversationReference, load_configuration_from_env
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import (
    CloudAdapter,
    jwt_authorization_middleware,
    start_agent_process,
)
from microsoft_agents.hosting.core import (
    AgentApplication,
    Authorization,
    MemoryStorage,
    TurnContext,
    TurnState,
)

import claude_session
import teams_files
from config import (
    ALLOWED_CHANNEL_PREFIXES,
    BOT_DISPLAY_NAME,
    CLAUDE_TIMEOUT,
    CLIENT_ID,
    PORT,
    PROJECT_DIR,
    audit_interaction,
    chunk_message,
    extract_attach_paths,
    get_reference,
    is_authorized,
    log_unauthorized,
    logger,
    save_reference,
    teams_markdown,
)

# ---------------------------------------------------------------------------
# Agents SDK app initialization (per bf-migration-python)
# ---------------------------------------------------------------------------

agents_sdk_config = load_configuration_from_env(environ)

STORAGE = MemoryStorage()
CONNECTION_MANAGER = MsalConnectionManager(**agents_sdk_config)
ADAPTER = CloudAdapter(connection_manager=CONNECTION_MANAGER)
AUTHORIZATION = Authorization(STORAGE, CONNECTION_MANAGER, **agents_sdk_config)

AGENT_APP = AgentApplication[TurnState](
    storage=STORAGE,
    adapter=ADAPTER,
    authorization=AUTHORIZATION,
    **agents_sdk_config,
)

# Keep strong refs to background tasks so they aren't garbage-collected.
_bg_tasks: set[asyncio.Task] = set()


# ---------------------------------------------------------------------------
# Proactive delivery — all Claude replies flow through here
# ---------------------------------------------------------------------------


def _reference_obj(reference_dict: dict) -> ConversationReference:
    return ConversationReference.model_validate(reference_dict)


async def _send_file(turn_context: TurnContext, file_path: str) -> None:
    """Share a local file: FileConsentCard in personal chats, SharePoint link in
    channels/groups."""
    conv = turn_context.activity.conversation
    conv_type = getattr(conv, "conversation_type", None)
    name = Path(file_path).name

    if conv_type == "personal":
        att = teams_files.build_file_consent_attachment(file_path)
        if att:
            await turn_context.send_activity(Activity(type="message", attachments=[att]))
        return

    link = await teams_files.upload_to_sharepoint(file_path)
    if link:
        await turn_context.send_activity(
            Activity(type="message", text=f"\U0001F4CE [{name}]({link})", text_format="markdown")
        )
    else:
        await turn_context.send_activity(
            Activity(type="message", text=f"(Generated `{name}` but file sharing isn't configured for this chat.)")
        )


async def deliver(reference_dict: dict, text: str) -> None:
    """Push a text block (and any attach: files) to a conversation proactively."""
    rendered = teams_markdown(text)
    files = extract_attach_paths(text)

    async def _callback(turn_context: TurnContext) -> None:
        for chunk in chunk_message(rendered):
            await turn_context.send_activity(
                Activity(type="message", text=chunk, text_format="markdown")
            )
        for fp in files:
            await _send_file(turn_context, str(fp))

    # bot_app_id is required so the SDK can mint a token for the outbound call.
    await ADAPTER.continue_conversation(_reference_obj(reference_dict), _callback, CLIENT_ID)


# ---------------------------------------------------------------------------
# Background turn runner
# ---------------------------------------------------------------------------


async def _run_and_reply(
    reference_dict: dict,
    conversation_id: str,
    user_id: str,
    prompt: str,
    is_public_channel: bool,
) -> None:
    state = {"first_sent": False, "skip": False, "texts": []}
    start = time.time()

    async def on_text(block: str) -> None:
        # SKIP on the very first block = channel relevance filter: stay silent.
        if not state["first_sent"] and is_public_channel and block.strip() == "SKIP":
            state["skip"] = True
            return
        state["texts"].append(block)
        await deliver(reference_dict, block)
        state["first_sent"] = True

    try:
        result = await claude_session.run_turn(
            conversation_id, user_id, prompt, on_text, CLAUDE_TIMEOUT
        )
    except Exception as e:
        logger.error(f"Turn failed for conv {conversation_id}: {e}")
        await deliver(reference_dict, f"Something went wrong: {e}")
        return

    if result["status"] == "timeout":
        minutes = CLAUDE_TIMEOUT // 60
        await deliver(
            reference_dict,
            f"Sorry, that timed out after {minutes} minutes. Try a simpler question?",
        )
        return
    if result["status"] == "died" and not state["texts"] and not state["skip"]:
        await deliver(
            reference_dict,
            "Sorry, I lost my train of thought. Could you try sending that again?",
        )
        return

    if state["skip"]:
        logger.info(f"Skipped message in conv {conversation_id} (not relevant)")
        return

    audit_interaction(
        user_id, conversation_id, "\n\n".join(state["texts"]),
        time.time() - start, result.get("session_id"),
    )


# ---------------------------------------------------------------------------
# Message handler
# ---------------------------------------------------------------------------


def _channel_label(activity) -> str:
    conv = activity.conversation
    return getattr(conv, "name", None) or getattr(conv, "id", "") or ""


@AGENT_APP.activity("message")
async def on_message(context: TurnContext, _state: TurnState):
    activity = context.activity
    sender = activity.from_property
    user_id = getattr(sender, "aad_object_id", None) or getattr(sender, "id", "")
    sender_name = getattr(sender, "name", None) or user_id
    conversation_id = activity.conversation.id
    conv_type = getattr(activity.conversation, "conversation_type", None)
    is_public_channel = conv_type == "channel"

    # Authorization. In DMs/groups, reject; in channels, let through (Claude
    # respects information boundaries), but still audit.
    if not is_authorized(user_id):
        log_unauthorized(user_id, conversation_id, activity.text or "")
        if conv_type in ("personal", "groupChat"):
            await context.send_activity(
                Activity(type="message", text="I only respond to authorized users.")
            )
            return True

    # Channel allowlist (best-effort: matches against conversation name when Teams
    # provides one).
    if is_public_channel and ALLOWED_CHANNEL_PREFIXES:
        label = _channel_label(activity)
        if label and not label.startswith(ALLOWED_CHANNEL_PREFIXES):
            logger.info(f"Ignoring message in non-allowed channel {label}")
            return True

    # Text, with the bot's @mention stripped.
    try:
        text = (activity.remove_recipient_mention() or "").strip()
    except Exception:
        text = (activity.text or "").strip()

    # Inbound attachments -> local files for Claude.
    attached = await teams_files.download_attachments(activity)
    if attached:
        paths = ", ".join(str(p) for p in attached)
        label = "Files attached" if len(attached) > 1 else "File attached"
        text = f"{label}: {paths}" + (f"\n\n{text}" if text else "")

    if not text:
        return True

    # Build the prompt, mirroring the Slack bot's attribution + channel SKIP gate.
    has_session = claude_session.get_session_id(conversation_id) is not None
    live = claude_session.has_live_process(conversation_id)
    if is_public_channel and not has_session and not live:
        label = _channel_label(activity)
        prompt = (
            f"A new message in channel {label}. Only respond if you're directly "
            "addressed by name or tagged or if you're already part of the "
            'conversation thread. Respond with exactly "SKIP" in ALL other cases.'
            f"\n\n[{sender_name}]({user_id}):\n{text}"
        )
    else:
        prompt = f"[{sender_name}]({user_id}):\n{text}"

    # Persist the conversation reference (proactive replies + CLI need it).
    reference = activity.get_conversation_reference()
    reference_dict = reference.model_dump(mode="json", by_alias=True)
    save_reference(conversation_id, reference_dict)

    # Ack fast: show a typing indicator, then run Claude in the background and
    # return so the webhook responds well within Teams' ~15s deadline.
    try:
        await context.send_activity(Activity(type="typing"))
    except Exception:
        pass

    task = asyncio.create_task(
        _run_and_reply(reference_dict, conversation_id, user_id, prompt, is_public_channel)
    )
    _bg_tasks.add(task)
    task.add_done_callback(_bg_tasks.discard)
    return True


@AGENT_APP.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState):
    for member in context.activity.members_added or []:
        if member.id != context.activity.recipient.id:
            await context.send_activity(
                Activity(type="message", text=f"Hi, I'm {BOT_DISPLAY_NAME}. Message me anytime.")
            )
    return True


@AGENT_APP.activity("invoke")
async def on_invoke(context: TurnContext, _state: TurnState):
    """Handle the FileConsentCard accept/decline invoke from Teams."""
    activity = context.activity
    if activity.name == "fileConsent/invoke":
        value = activity.value or {}
        action = value.get("action")
        if action == "accept":
            info = await teams_files.handle_file_consent_accept(value)
            if info:
                await context.send_activity(Activity(type="message", attachments=[info]))
        elif action == "decline":
            teams_files.handle_file_consent_decline(value)
    return True


# ---------------------------------------------------------------------------
# aiohttp host
# ---------------------------------------------------------------------------


async def _entry_point(req: Request) -> Response:
    # start_agent_process wires the request through the adapter + AgentApplication.
    return await start_agent_process(req, req.app["agent_app"], req.app["adapter"])


async def _health(_req: Request) -> Response:
    return json_response({"status": "ok", "bot": "claudette-teams"})


async def _on_startup(app: Application) -> None:
    app["cleanup"] = asyncio.create_task(claude_session.cleanup_idle_sessions())
    logger.info(f"{BOT_DISPLAY_NAME} ready on port {PORT} | project_dir={PROJECT_DIR}")


def create_app() -> Application:
    app = Application(middlewares=[jwt_authorization_middleware])
    app["agent_app"] = AGENT_APP
    app["adapter"] = ADAPTER
    app.router.add_post("/api/messages", _entry_point)
    app.router.add_get("/health", _health)
    app.on_startup.append(_on_startup)
    return app


# ---------------------------------------------------------------------------
# CLI proactive modes
# ---------------------------------------------------------------------------


async def _cli_send(conversation_id: str, message: str) -> None:
    reference_dict = get_reference(conversation_id)
    if not reference_dict:
        print(
            f"No stored conversation reference for '{conversation_id}'. Teams cannot "
            "cold-message — the user/channel must message the bot at least once first."
        )
        return
    await deliver(reference_dict, message)
    print("sent")


def main() -> None:
    parser = argparse.ArgumentParser(description="Claudette — Microsoft Teams bot powered by Claude Code")
    parser.add_argument("--send", nargs=2, metavar=("CONVERSATION_ID", "MESSAGE"),
                        help="Proactively message a known conversation and exit")
    parser.add_argument("--channel", nargs=2, metavar=("CONVERSATION_ID", "MESSAGE"),
                        help="Alias of --send (Teams treats both the same)")
    parser.add_argument("--send-result", metavar="CONVERSATION_ID",
                        help="Read Claude JSON from stdin and push to a conversation")
    args = parser.parse_args()

    if args.send:
        asyncio.run(_cli_send(args.send[0], args.send[1]))
        return
    if args.channel:
        asyncio.run(_cli_send(args.channel[0], args.channel[1]))
        return
    if args.send_result:
        raw = sys.stdin.read().strip()
        try:
            message = json.loads(raw).get("result", "") or "Job completed but produced no output."
        except json.JSONDecodeError:
            message = raw or "Job completed but produced no output."
        asyncio.run(_cli_send(args.send_result, message))
        return

    run_app(create_app(), host="0.0.0.0", port=PORT)


if __name__ == "__main__":
    main()
