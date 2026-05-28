# Claudette-Teams

Microsoft Teams bot powered by Claude Code — a port of the Slack [`claudette`](../claudette)
bot to the **Microsoft 365 Agents SDK**. It bridges Teams conversations to a
long-lived `claude` CLI process and streams replies back.

## How it works

```
Teams ──POST /api/messages──► aiohttp + AgentApplication (CloudAdapter, MSAL JWT)
                                  │  @AGENT_APP.activity("message")
                                  ├─ ack fast: typing indicator
                                  ├─ persist ConversationReference
                                  └─ background task:
                                       ├─ live `claude` subprocess per conversation (stream-json)
                                       └─ proactively push each text block + attach: files
```

Teams enforces a ~15s webhook deadline, so the handler returns immediately and
replies are delivered **proactively** via a stored conversation reference.

## Files

| File | Role |
|---|---|
| `bot.py` | Agents SDK wiring, message/invoke handlers, proactive delivery, CLI, aiohttp host |
| `claude_session.py` | Async live-session pool: `asyncio` subprocesses, stream-json reader, idle cleanup |
| `teams_files.py` | FileConsentCard upload (DM), Graph/SharePoint upload (channel), inbound download |
| `config.py` | Env, logging, auth tiers, session/reference persistence, markdown + chunking |
| `appManifest/` | Teams app manifest + icons (zip → sideload) |
| `setup-teams.md` | Azure Bot + App Service + manifest walkthrough |
| `CLAUDE.md.example` | Persona / ops manual for the Claude side — copy to `PROJECT_DIR` as `CLAUDE.md` |

## Run it locally (no Azure needed)

You can run and test the whole bot on your own machine with the **Bot Framework
Emulator** — no Azure, no Teams, no tunnel. This exercises the real code path
(message in → Claude subprocess → streamed reply out).

### Prerequisites

| Tool | Why | Check |
|---|---|---|
| **Python 3.10+** (3.11 recommended) | The Agents SDK needs 3.10+ | `python --version` |
| **Claude Code CLI** on your `PATH`, logged in | The bot shells out to `claude` | `claude --version` |
| **Bot Framework Emulator** | Local Teams-less chat client | [download](https://github.com/microsoft/BotFramework-Emulator/releases) |
| **git** | clone the repo | `git --version` |

> The bot does **not** call the Anthropic API directly — it drives your local
> `claude` CLI. If `claude` isn't installed and authenticated, nothing will reply.

### 1. Get the code and a virtual environment

```powershell
cd path\to\claudette-teams
python -m venv venv
.\venv\Scripts\Activate.ps1        # PowerShell. (cmd: venv\Scripts\activate.bat)
pip install -r requirements.txt
```

On macOS/Linux the activate line is `source venv/bin/activate`.

### 2. Create your `.env`

```powershell
Copy-Item .env.example .env
```

Edit `.env`. For a **local** run you only need these — leave the Azure creds blank
and turn on anonymous auth so the Emulator can connect without tokens:

```env
# Local auth bypass — Emulator sends no JWT. NEVER set this in Azure.
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__ANONYMOUS_ALLOWED=True
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=

# Where Claude runs. Point at the repo you want it to work in.
PROJECT_DIR=C:\path\to\your\working-directory

PORT=3978

# Leave AUTHORIZED_USERS BLANK locally = allow everyone (the Emulator has no AAD id).
AUTHORIZED_USERS=

# Must be a model your local `claude` CLI accepts. Change if needed.
CLAUDE_MODEL=claude-opus-4-6[1m]
CLAUDE_EFFORT=medium
CLAUDE_TIMEOUT=7200
```

> Why blank `AUTHORIZED_USERS`? Auth keys on the sender's Entra **object id**, which
> the Emulator doesn't provide — so an allowlist would reject every local message.

### 3. Set up the Claude-side persona (`CLAUDE.md`)

The bot launches `claude` with `PROJECT_DIR` as its working directory, so Claude
Code automatically picks up a `CLAUDE.md` in that folder. **This is how Claudette
gets her voice, file-sharing conventions, and channel/DM behavior** — without it
you'll get a generic, chatty Claude.

```powershell
Copy-Item CLAUDE.md.example $env:PROJECT_DIR\CLAUDE.md
```

Edit the copy if you want to change the persona, but keep the sections on
**Teams markdown**, **`attach:<path>`**, and **channel `SKIP` behavior** — the
bot depends on those conventions.

### 4. Confirm `claude` works from this folder

```powershell
claude -p "say hello" --model claude-opus-4-6[1m]
```

If that errors (not on PATH, not logged in, bad model name), fix it before
continuing — the bot can only be as healthy as the CLI it drives.

### 5. Start the bot

```powershell
python bot.py
```

Expected: a log line `Claudette ready on port 3978 | project_dir=...`. Sanity-check
the health route in another terminal:

```powershell
curl http://localhost:3978/health      # -> {"status":"ok","bot":"claudette-teams"}
```

### 6. Connect the Emulator

1. Open **Bot Framework Emulator → Open Bot**.
2. **Bot URL:** `http://localhost:3978/api/messages`
3. Leave **Microsoft App ID** and **App password** empty, then **Connect**.
4. Type a message. You should see a typing indicator, then Claude's reply stream in.

Each Emulator conversation maps to one live `claude` process; follow-up messages
resume the same session.

### 7. (Optional) Test from real Teams while still running locally

The Emulator covers everything except Teams-specific UI (FileConsentCard, channel
mentions). To hit the local bot from actual Teams you must expose port 3978 publicly
and register it:

1. Tunnel it: `cloudflared tunnel --url http://localhost:3978` (or `ngrok http 3978`,
   or VS Code dev tunnels). Copy the public `https://…` URL.
2. Create an **Azure Bot** registration (see [`setup-teams.md`](setup-teams.md) §1) and
   set its **Messaging endpoint** to `https://<tunnel>/api/messages`.
3. Put the real `CLIENTID` / `CLIENTSECRET` / `TENANTID` in `.env`, remove
   `ANONYMOUS_ALLOWED`, set `AUTHORIZED_USERS` to your AAD object id, restart.
4. Sideload the manifest (see [`setup-teams.md`](setup-teams.md) §4).

### Local troubleshooting

| Symptom | Fix |
|---|---|
| `ModuleNotFoundError: microsoft_agents...` | `pip install -r requirements.txt` inside the activated venv. Imports use **underscores** (`microsoft_agents`), not dots. |
| No reply, no error | `claude` isn't on PATH / not logged in / `CLAUDE_MODEL` invalid. Test step 3. |
| `401` in the Emulator | `ANONYMOUS_ALLOWED=True` missing, or you typed an App ID/password in the Emulator. Leave both blank. |
| `NotImplementedError` about subprocess on Windows | You're on a non-default event loop. Python's default Proactor loop supports subprocesses; don't override the loop policy. |
| Port 3978 in use | Change `PORT` in `.env` and the Emulator URL to match. |
| Reply is truncated | Expected for very long output — it's chunked into multiple messages (~17k chars each). |

## Azure / Teams deployment

See [`setup-teams.md`](setup-teams.md). TL;DR: provision an Azure Bot + App Service,
set the `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__*` app settings, deploy, point
the messaging endpoint at `/api/messages`, and sideload the manifest.

## Differences from the Slack bot

- **Auth keys on AAD object id** (`activity.from.aadObjectId`), not a Slack user id.
- **Sessions key on `conversation.id`**, not a Slack `thread_ts`.
- **No cold DM** — Teams can only be messaged back on a conversation it has already
  seen. CLI `--send <conversation_id>` requires a stored reference.
- **Files** use FileConsentCard (DMs) and Graph + SharePoint (channels) instead of
  Slack's upload API.

## Known limitations

- Channel inbound files arrive as HTML stubs unless Graph perms + SharePoint are set.
- Channel message-history context isn't fetched (would need `ChannelMessage.Read`);
  continuity relies on the resumed Claude session instead.
- `ALLOWED_CHANNELS` prefix matching is best-effort (depends on Teams supplying a name).

## Verification points

A few Agents SDK call sites are marked `VERIFY-AGAINST-SAMPLE` in the source — confirm
the exact `start_agent_process` / `jwt_authorization_middleware` / `continue_conversation`
signatures against https://github.com/microsoft/Agents `samples/python` on first run.
