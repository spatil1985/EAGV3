"""MCP server — 14 tools over stdio transport. Start with: uv run python mcp_server.py"""
from __future__ import annotations

import os
import re
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import httpx
from bs4 import BeautifulSoup
from duckduckgo_search import DDGS
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

import persistent_store

load_dotenv()

mcp = FastMCP("padma-tools")


@mcp.tool()
def web_search(query: str, max_results: int = 5) -> str:
    """Search the web using Tavily (falls back to DuckDuckGo) and return results."""
    api_key = os.getenv("TAVILY_API_KEY", "")
    if api_key:
        try:
            resp = httpx.post(
                "https://api.tavily.com/search",
                json={
                    "api_key": api_key,
                    "query": query,
                    "max_results": max_results,
                    "search_depth": "basic",
                },
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            results = data.get("results", [])
            if not results:
                return "No results found."
            lines = []
            for r in results:
                content = r.get("content") or r.get("snippet", "")
                lines.append(f"- [{r['title']}]({r['url']})\n  {content[:400]}")
            return "\n\n".join(lines)
        except Exception as e:
            return f"Tavily search error: {e}"

    # DuckDuckGo fallback
    results = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(query, max_results=max_results):
                results.append(f"- [{r['title']}]({r['href']})\n  {r['body']}")
    except Exception as e:
        return f"DuckDuckGo search error: {e}"
    return "\n\n".join(results) if results else "No results found."


@mcp.tool()
def fetch_url(url: str) -> str:
    """Fetch a URL and return its text content (strips HTML tags)."""
    headers = {"User-Agent": "Mozilla/5.0 (compatible; PADMA-Agent/1.0)"}
    try:
        resp = httpx.get(url, headers=headers, timeout=20, follow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        text = soup.get_text(separator="\n", strip=True)
        # Trim to 8000 chars to avoid context overflow
        return text[:8000] if len(text) > 8000 else text
    except Exception as e:
        return f"Error fetching {url}: {e}"


@mcp.tool()
def get_time(timezone: str = "UTC") -> str:
    """Get the current date and time in the specified timezone."""
    try:
        tz = ZoneInfo(timezone)
    except ZoneInfoNotFoundError:
        tz = ZoneInfo("UTC")
    now = datetime.now(tz)
    return now.strftime(f"%Y-%m-%d %H:%M:%S %Z (timezone: {timezone})")


@mcp.tool()
def currency_convert(amount: float, from_currency: str, to_currency: str) -> str:
    """Convert an amount between currencies using a free exchange rate API."""
    try:
        url = f"https://api.frankfurter.app/latest?amount={amount}&from={from_currency.upper()}&to={to_currency.upper()}"
        resp = httpx.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        rate_info = data.get("rates", {})
        to_upper = to_currency.upper()
        if to_upper in rate_info:
            converted = rate_info[to_upper]
            return f"{amount} {from_currency.upper()} = {converted} {to_upper}"
        return f"Conversion result: {rate_info}"
    except Exception as e:
        return f"Currency conversion error: {e}"


@mcp.tool()
def read_file(path: str) -> str:
    """Read and return the contents of a file."""
    try:
        return Path(path).read_text(encoding="utf-8")
    except Exception as e:
        return f"Error reading {path}: {e}"


@mcp.tool()
def create_file(path: str, content: str) -> str:
    """Create a new file with the given content. Fails if file already exists."""
    p = Path(path)
    if p.exists():
        return f"Error: {path} already exists. Use update_file to overwrite."
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
        return f"Created {path} ({len(content)} chars)"
    except Exception as e:
        return f"Error creating {path}: {e}"


@mcp.tool()
def update_file(path: str, content: str) -> str:
    """Overwrite a file with new content (creates it if it doesn't exist)."""
    try:
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
        return f"Updated {path} ({len(content)} chars)"
    except Exception as e:
        return f"Error updating {path}: {e}"


@mcp.tool()
def edit_file(path: str, old_text: str, new_text: str) -> str:
    """Replace the first occurrence of old_text with new_text in a file."""
    try:
        p = Path(path)
        original = p.read_text(encoding="utf-8")
        if old_text not in original:
            return f"Error: old_text not found in {path}"
        updated = original.replace(old_text, new_text, 1)
        p.write_text(updated, encoding="utf-8")
        return f"Edited {path}: replaced {len(old_text)} chars"
    except Exception as e:
        return f"Error editing {path}: {e}"


@mcp.tool()
def list_dir(path: str = ".") -> str:
    """List the contents of a directory."""
    try:
        entries = list(Path(path).iterdir())
        lines = []
        for e in sorted(entries):
            kind = "DIR" if e.is_dir() else "FILE"
            lines.append(f"{kind}  {e.name}")
        return "\n".join(lines) if lines else "(empty)"
    except Exception as e:
        return f"Error listing {path}: {e}"


# ── Persistent long-term memory tools ────────────────────────────────────────

@mcp.tool()
def save_memory(key: str, value: str, category: str = "fact", tags: str = "", notes: str = "") -> str:
    """Save or update a named fact in persistent long-term memory.

    key      : unique name (e.g. 'mom_birthday', 'user_city')
    value    : the data to store (any text)
    category : 'fact' | 'date' | 'preference' | 'reminder' | 'contact'
    tags     : comma-separated keywords for search (e.g. 'birthday,family')
    notes    : optional extra context
    """
    tag_list = [t.strip() for t in tags.split(",") if t.strip()]
    entry = persistent_store.save(key, value, category, tag_list, notes)
    return f"Saved to persistent memory: key='{key}' value='{value}' category='{category}' (id={entry['id']})"


@mcp.tool()
def load_memory(key: str) -> str:
    """Retrieve a named fact from persistent long-term memory by its key."""
    entry = persistent_store.load(key)
    if entry is None:
        return f"No memory found for key='{key}'"
    return (
        f"key: {entry['key']}\n"
        f"value: {entry['value']}\n"
        f"category: {entry['category']}\n"
        f"tags: {', '.join(entry.get('tags', []))}\n"
        f"notes: {entry.get('notes', '')}\n"
        f"saved: {entry.get('created_at', '?')}  updated: {entry.get('updated_at', '?')}"
    )


@mcp.tool()
def search_memories(query: str) -> str:
    """Search persistent long-term memory by keyword. Returns matching entries."""
    results = persistent_store.search(query)
    if not results:
        return f"No memories found matching '{query}'"
    lines = [f"Found {len(results)} result(s) for '{query}':\n"]
    for r in results[:10]:
        lines.append(f"  [{r['category']}] {r['key']} = {r['value']}  (tags: {', '.join(r.get('tags',[]))})")
    return "\n".join(lines)


@mcp.tool()
def list_memories(category: str = "") -> str:
    """List all entries in persistent long-term memory. Optionally filter by category."""
    items = persistent_store.list_all()
    if category:
        items = [i for i in items if i.get("category") == category]
    if not items:
        return "Persistent memory is empty." if not category else f"No memories in category '{category}'"
    lines = [f"{len(items)} memory entries (newest first):\n"]
    for item in items:
        lines.append(
            f"  [{item['category']}] {item['key']} = {item['value']}"
            + (f"  | tags: {', '.join(item.get('tags', []))}" if item.get("tags") else "")
            + f"  | {item.get('updated_at','?')[:10]}"
        )
    return "\n".join(lines)


# ── Google Calendar tool ──────────────────────────────────────────────────────

_GCAL_SCOPES     = ["https://www.googleapis.com/auth/calendar"]
_CREDS_PATH      = Path("persistent_mem/credentials.json")
_TOKEN_PATH      = Path("persistent_mem/token.json")
_EVENTS_DIR      = Path("persistent_mem/events")


def _build_gcal_service():
    """Return an authorised Google Calendar service, or None if not configured."""
    if not _TOKEN_PATH.exists():
        return None
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build

        creds = Credentials.from_authorized_user_file(str(_TOKEN_PATH), _GCAL_SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            _TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
        if creds.valid:
            return build("calendar", "v3", credentials=creds)
    except Exception:
        pass
    return None


def _save_ics(title: str, start: datetime, end: datetime, description: str, reminder_minutes: int) -> str:
    """Create an ICS file and return the path."""
    _EVENTS_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w\-]", "_", title)[:40]
    ics_path = _EVENTS_DIR / f"{safe}_{start.strftime('%Y%m%d')}.ics"

    def fmt(dt: datetime) -> str:
        return dt.strftime("%Y%m%dT%H%M%S")

    ics = (
        "BEGIN:VCALENDAR\r\n"
        "VERSION:2.0\r\n"
        "PRODID:-//PADMA Agent//EN\r\n"
        "BEGIN:VEVENT\r\n"
        f"UID:{uuid.uuid4()}@padma\r\n"
        f"DTSTART:{fmt(start)}\r\n"
        f"DTEND:{fmt(end)}\r\n"
        f"SUMMARY:{title}\r\n"
        f"DESCRIPTION:{description}\r\n"
        "BEGIN:VALARM\r\n"
        f"TRIGGER:-PT{reminder_minutes}M\r\n"
        "ACTION:DISPLAY\r\n"
        f"DESCRIPTION:{title}\r\n"
        "END:VALARM\r\n"
        "END:VEVENT\r\n"
        "END:VCALENDAR\r\n"
    )
    ics_path.write_text(ics, encoding="utf-8")
    return str(ics_path)


@mcp.tool()
def create_calendar_event(
    title: str,
    date: str,
    time: str = "09:00",
    description: str = "",
    reminder_minutes: int = 30,
    duration_minutes: int = 60,
) -> str:
    """Create a Google Calendar event or reminder.

    date     : ISO date string  e.g. '2026-05-01'
    time     : 24h time string  e.g. '09:00'  (default 09:00)
    reminder_minutes : popup reminder N minutes before (default 30)
    duration_minutes : event length in minutes (default 60)

    Requires Google OAuth setup (run setup_gcal.py once).
    Falls back to creating an .ics file in persistent_mem/events/ if not authenticated.
    """
    try:
        start = datetime.fromisoformat(f"{date}T{time}:00")
    except ValueError:
        return f"Invalid date/time: date='{date}' time='{time}'. Use YYYY-MM-DD and HH:MM."
    end = start + timedelta(minutes=duration_minutes)

    # ── Try Google Calendar API ──
    service = _build_gcal_service()
    if service:
        try:
            event_body = {
                "summary": title,
                "description": description,
                "start": {"dateTime": start.isoformat(), "timeZone": "UTC"},
                "end":   {"dateTime": end.isoformat(),   "timeZone": "UTC"},
                "reminders": {
                    "useDefault": False,
                    "overrides": [{"method": "popup", "minutes": reminder_minutes}],
                },
            }
            created = service.events().insert(calendarId="primary", body=event_body).execute()
            link = created.get("htmlLink", "")
            return f"Google Calendar event created!\nTitle: {title}\nDate: {date} {time}\nLink: {link}"
        except Exception as exc:
            pass  # fall through to ICS

    # ── ICS fallback ──
    ics_path = _save_ics(title, start, end, description, reminder_minutes)
    return (
        f"ICS calendar file created: {ics_path}\n"
        f"Open this file to import the event into any calendar app (Google, Outlook, Apple).\n"
        f"Title: {title} | Date: {date} {time} | Reminder: {reminder_minutes}min before\n"
        f"For automatic Google Calendar sync: run  uv run python setup_gcal.py"
    )


if __name__ == "__main__":
    mcp.run(transport="stdio")
