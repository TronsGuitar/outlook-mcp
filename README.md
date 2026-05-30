# outlook-mcp

A small, self-hosted **MCP server** that gives an LLM (Claude Code, Claude Desktop, or any
MCP client) **read-only** access to a **personal Outlook.com** mailbox (outlook.com /
hotmail.com / live.com) through the Microsoft Graph API.

The python mcp server can be installed at https://pypi.org/project/outlook-personal-mcp/

It authenticates against Microsoft's **consumer** OAuth endpoint (`/consumers`), which is
what lets *personal* Microsoft accounts sign in. Most "Microsoft 365" connectors are
registered as work/school-only apps and will reject a personal account — this one is built
specifically for personal mailboxes.

> **Your data stays local.** The server runs on your machine and caches your OAuth token in
> `~/.outlook-mcp/token.json` (file permissions `600`). Nothing is sent anywhere except
> directly to Microsoft Graph.

## Tools

| Tool | Description |
| --- | --- |
| `search_mail(query, max_results)` | Full-text search across the mailbox. |
| `list_recent_mail(max_results, unread_only)` | Newest inbox messages. |
| `get_message(message_id)` | Full content of one message. |

All tools are read-only — the app requests only the `Mail.Read` scope. No send, no delete.

---

## Setup

### 1. Register your own Azure app (free — no paid subscription needed)

Each user brings their own app registration, so no shared infrastructure or credentials.

1. Go to <https://entra.microsoft.com> → **App registrations** → **New registration**.
2. **Name:** anything, e.g. `Outlook MCP`.
3. **Supported account types:** **Personal Microsoft accounts only**
   *(or "any org directory and personal Microsoft accounts" if you also use work accounts).*
4. **Redirect URI:** platform **Public client/native (mobile & desktop)**, value:
   `http://localhost:8765/callback`
5. **Register**, then copy the **Application (client) ID** from the Overview page.
6. **Authentication** → enable **Allow public client flows** → **Save**.
7. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** →
   add `Mail.Read` (also `offline_access` and `User.Read` if not auto-added). No admin
   consent is required for a personal account.

### 2. Install

With [uv](https://docs.astral.sh/uv/) (recommended):

```sh
# clone, then from the repo root:
uv sync
```

Or install into any environment with pip:

```sh
pip install .
```

### 3. Configure and sign in (one time)

```sh
# PowerShell
$env:OUTLOOK_MCP_CLIENT_ID = "<your-application-client-id>"
uv run outlook-mcp-auth
```

```sh
# bash / zsh
export OUTLOOK_MCP_CLIENT_ID="<your-application-client-id>"
uv run outlook-mcp-auth
```

A browser opens; sign in and approve. Your refresh token is cached and auto-renews after that.

### 4. Register with your MCP client

**Claude Code:**

```sh
claude mcp add --scope user outlook \
  --env OUTLOOK_MCP_CLIENT_ID=<client-id> \
  -- uv --directory /path/to/outlook-mcp run outlook-mcp
```

**Claude Desktop / other clients** — add to the MCP config file:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["--directory", "/path/to/outlook-mcp", "run", "outlook-mcp"],
      "env": { "OUTLOOK_MCP_CLIENT_ID": "<client-id>" }
    }
  }
}
```

A one-click **Desktop Extension** bundle (`manifest.json`) is also included — see
[Desktop Extension](#desktop-extension) below.

---

## Configuration

| Env var | Required | Default | Purpose |
| --- | --- | --- | --- |
| `OUTLOOK_MCP_CLIENT_ID` | yes | — | Application (client) ID from your Azure app registration. |
| `OUTLOOK_MCP_REDIRECT_PORT` | no | `8765` | Loopback port; must match the redirect URI port you registered. |

## Desktop Extension

`manifest.json` follows the [MCP Bundle / Desktop Extension](https://github.com/anthropics/mcpb)
spec. It prompts the user for their client ID on install. To build the installable bundle:

```sh
npx @anthropic-ai/mcpb pack
```

Then open the resulting `.mcpb` file with Claude Desktop.

## Re-authenticating

If the cached token is revoked or expires beyond refresh, just run the sign-in again:

```sh
uv run outlook-mcp-auth
```

## Security notes

- Read-only (`Mail.Read`) — the server cannot send, move, or delete mail.
- Tokens never leave your machine; `~/.outlook-mcp/token.json` is excluded from git.
- Each user uses their own Azure app registration, so there is no shared client secret.

## License

MIT — see [LICENSE](LICENSE).
