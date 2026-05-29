"""FastMCP server exposing read-only Outlook.com mail tools via Microsoft Graph."""

from __future__ import annotations

from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

mcp = FastMCP("outlook-mcp")


def _graph_get(path: str, params: dict[str, Any] | None = None) -> dict:
    token = get_access_token()
    resp = httpx.get(
        f"{GRAPH_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    if resp.status_code == 401:
        raise RuntimeError(
            "Outlook rejected the access token (401). Try re-running `uv run outlook-mcp-auth`."
        )
    if resp.status_code >= 400:
        raise RuntimeError(f"Graph API error ({resp.status_code}): {resp.text}")
    return resp.json()


def _fmt_message(m: dict) -> dict:
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "received": m.get("receivedDateTime"),
        "isRead": m.get("isRead"),
        "hasAttachments": m.get("hasAttachments"),
        "preview": m.get("bodyPreview"),
        "webLink": m.get("webLink"),
    }


@mcp.tool()
def search_mail(query: str, max_results: int = 15) -> list[dict]:
    """Search the Outlook mailbox for messages matching a free-text query.

    Searches across subject, body, sender and recipients (Graph $search). Returns
    message summaries; use get_message with an id to read full content.

    Args:
        query: Free-text search, e.g. 'movie tickets' or 'Cinemark confirmation'.
        max_results: Max messages to return (1-50).
    """
    max_results = max(1, min(max_results, 50))
    data = _graph_get(
        "/me/messages",
        params={
            "$search": f'"{query}"',
            "$top": max_results,
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink",
        },
    )
    return [_fmt_message(m) for m in data.get("value", [])]


@mcp.tool()
def list_recent_mail(max_results: int = 15, unread_only: bool = False) -> list[dict]:
    """List the most recent messages in the inbox, newest first.

    Args:
        max_results: Max messages to return (1-50).
        unread_only: If true, only return unread messages.
    """
    max_results = max(1, min(max_results, 50))
    params: dict[str, Any] = {
        "$top": max_results,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink",
    }
    if unread_only:
        params["$filter"] = "isRead eq false"
    data = _graph_get("/me/mailFolders/inbox/messages", params=params)
    return [_fmt_message(m) for m in data.get("value", [])]


@mcp.tool()
def get_message(message_id: str) -> dict:
    """Fetch the full content of a single message by its id.

    Args:
        message_id: The message id returned by search_mail or list_recent_mail.
    """
    m = _graph_get(
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "isRead,hasAttachments,body,webLink"
        },
    )
    out = _fmt_message(m)
    out["to"] = [r.get("emailAddress", {}).get("address") for r in m.get("toRecipients", [])]
    out["cc"] = [r.get("emailAddress", {}).get("address") for r in m.get("ccRecipients", [])]
    body = m.get("body") or {}
    out["bodyContentType"] = body.get("contentType")
    out["body"] = body.get("content")
    return out


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
