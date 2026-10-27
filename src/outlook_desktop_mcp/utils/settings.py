"""User settings for the server: env var > settings file > built-in default.

Each setting has an environment variable (for people who configure the server in
Claude's config file) and a key in a small JSON settings file that the ``set_setting``
tool writes, so a user can change a setting by asking Claude and it sticks across
restarts and ``git pull``. The file lives outside the repo:

    %LOCALAPPDATA%\\outlook-desktop-mcp\\settings.json   (or OUTLOOK_MCP_SETTINGS_FILE)

Values are read on every lookup (the file is re-read when its mtime changes), so a
change takes effect immediately in the running server. An environment variable always
wins over the file; ``get_with_source`` says which one applied. ``open_in_outlook_mode``
and ``should_open`` resolve which saved drafts open in an Outlook window.
"""
from __future__ import annotations

import json
import os
import pathlib
import threading
from typing import Any

_TRUE = {"1", "true", "yes", "on"}

SETTINGS: dict[str, dict] = {
    "default_account": {
        "env": "OUTLOOK_MCP_DEFAULT_ACCOUNT", "type": "str", "default": "",
        "help": "Account (display name or unique substring) used when a tool is not given one. "
                "Empty = Outlook's own default mailbox.",
    },
    "hide_accounts": {
        "env": "OUTLOOK_MCP_HIDE_ACCOUNTS", "type": "list", "default": [],
        "help": "Accounts (display name substrings, semicolon-separated) left out of list_accounts "
                "and of account=\"all\" searches. They can still be named explicitly.",
    },
    "open_in_outlook": {
        "env": "OUTLOOK_MCP_OPEN_IN_OUTLOOK", "type": "choice", "choices": ("none", "meetings", "all"),
        "default": "meetings",
        "help": "Which saved drafts open in an Outlook window: 'meetings' (default) opens unsent meeting "
                "invitations only, so they are not lost on the calendar; 'all' opens every draft; 'none' "
                "opens nothing. Email drafts always land in Drafts. The open_in_outlook tool opens any "
                "item on request.",
    },
    "open_drafts": {
        "env": "OUTLOOK_MCP_OPEN_DRAFTS", "type": "bool", "default": False,
        "help": "Older name for open_in_outlook: true = all, false = none. Ignored when open_in_outlook is set.",
    },
    "font_family": {"env": "OUTLOOK_MCP_FONT_FAMILY", "type": "str", "default": "Aptos",
                    "help": "Font for text the server writes into drafts."},
    "font_size_pt": {"env": "OUTLOOK_MCP_FONT_SIZE_PT", "type": "str", "default": "12",
                     "help": "Font size in points for draft text."},
    "font_color": {"env": "OUTLOOK_MCP_FONT_COLOR", "type": "str", "default": "#000000",
                   "help": "Font colour for draft text."},
    "plain_text": {"env": "OUTLOOK_MCP_PLAIN_TEXT", "type": "bool", "default": False,
                   "help": "Write drafts as plain text instead of styled HTML."},
    "include_shared_in_all": {
        "env": "OUTLOOK_MCP_INCLUDE_SHARED_IN_ALL", "type": "bool", "default": False,
        "help": "Include shared Exchange mailboxes that are not cached locally in account=\"all\" "
                "searches. Off = skip them (every folder listing in such a mailbox is a server round "
                "trip and can take tens of seconds); they can still be named explicitly.",
    },
    "call_timeout_s": {
        "env": "OUTLOOK_MCP_CALL_TIMEOUT_S", "type": "str", "default": "60",
        "help": "Seconds a tool call may take. Long searches stop a few seconds before this and "
                "return what they found so far, marked partial.",
    },
    "legacy_mode": {
        "env": "OUTLOOK_MCP_LEGACY_MODE", "type": "bool", "default": False,
        "help": "Use the pre-0.6 search code paths (Python walks every message instead of asking "
                "Outlook to filter). Much slower; only for troubleshooting the new paths.",
    },
}

_lock = threading.Lock()
_cache: dict[str, Any] = {"path": None, "mtime": None, "data": {}}


def settings_path() -> pathlib.Path:
    override = os.environ.get("OUTLOOK_MCP_SETTINGS_FILE", "").strip()
    if override:
        return pathlib.Path(override)
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CONFIG_HOME") or str(pathlib.Path.home() / ".config")
    return pathlib.Path(base) / "outlook-desktop-mcp" / "settings.json"


def _coerce(key: str, raw: Any, strict: bool = False) -> Any:
    """Turn a raw env, file or user value into the setting's type.

    ``strict`` (used when the user sets a value) raises ValueError for a choice outside
    the allowed set; otherwise an unknown choice falls back to the default, so a stale
    file or env var cannot break every lookup.
    """
    spec = SETTINGS[key]
    kind = spec["type"]
    if kind == "choice":
        text = str(raw or "").strip().lower()
        if text in spec["choices"]:
            return text
        if strict:
            raise ValueError(f"{key} must be one of {', '.join(spec['choices'])}; got {raw!r}")
        return spec["default"]
    if kind == "bool":
        if isinstance(raw, bool):
            return raw
        return str(raw).strip().lower() in _TRUE
    if kind == "list":
        if isinstance(raw, list):
            items = [str(x) for x in raw]
        else:
            text = str(raw).strip()
            if text.startswith("["):
                try:
                    items = [str(x) for x in json.loads(text)]
                except Exception:
                    items = [text]
            else:
                items = text.split(";")
        return [x.strip() for x in items if x.strip()]
    return str(raw).strip()


def _file_data() -> dict:
    path = settings_path()
    with _lock:
        try:
            mtime = path.stat().st_mtime
        except OSError:
            _cache.update(path=path, mtime=None, data={})
            return {}
        if _cache["path"] == path and _cache["mtime"] == mtime:
            return _cache["data"]
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig") or "{}")
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
        _cache.update(path=path, mtime=mtime, data=data)
        return data


def get_with_source(key: str) -> tuple[Any, str]:
    """``(value, source)`` where source is "env", "file" or "default"."""
    spec = SETTINGS[key]
    env = os.environ.get(spec["env"])
    if env is not None and env.strip() != "":
        return _coerce(key, env), "env"
    data = _file_data()
    if key in data and data[key] not in (None, ""):
        return _coerce(key, data[key]), "file"
    return spec["default"], "default"


def get(key: str) -> Any:
    return get_with_source(key)[0]


def all_settings() -> dict:
    out = {}
    for key, spec in SETTINGS.items():
        value, source = get_with_source(key)
        out[key] = {"value": value, "source": source, "env_var": spec["env"], "help": spec["help"]}
    return out


def set_value(key: str, value: Any) -> Any:
    """Write one setting to the file (creating it) and return the coerced value."""
    if key not in SETTINGS:
        raise KeyError(f"Unknown setting {key!r}. Known: {', '.join(SETTINGS)}")
    coerced = _coerce(key, value, strict=True)
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = dict(_file_data())
    data[key] = coerced
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    os.replace(tmp, path)
    with _lock:
        _cache.update(path=path, mtime=None, data={})  # force re-read
    return coerced


def clear_value(key: str) -> None:
    """Remove one setting from the file so the default (or env var) applies again."""
    if key not in SETTINGS:
        raise KeyError(f"Unknown setting {key!r}. Known: {', '.join(SETTINGS)}")
    path = settings_path()
    data = dict(_file_data())
    if key in data:
        del data[key]
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        os.replace(tmp, path)
    with _lock:
        _cache.update(path=path, mtime=None, data={})


def open_in_outlook_mode() -> str:
    """``"none"``, ``"meetings"`` or ``"all"``: which saved drafts open in an Outlook window.

    ``open_in_outlook`` wins when it is set (env or file). Otherwise the older
    ``open_drafts`` boolean maps to ``all`` / ``none`` so existing installs keep behaving.
    Otherwise the default, ``meetings``.
    """
    value, source = get_with_source("open_in_outlook")
    if source != "default":
        return value
    legacy, legacy_source = get_with_source("open_drafts")
    if legacy_source != "default":
        return "all" if legacy else "none"
    return value


def should_open(kind: str) -> bool:
    """Whether a draft of ``kind`` ("email", "reply", "reply_all", "forward", "update",
    "meeting") opens in an Outlook window after it is saved."""
    mode = open_in_outlook_mode()
    if mode == "all":
        return True
    if mode == "meetings":
        return kind == "meeting"
    return False
