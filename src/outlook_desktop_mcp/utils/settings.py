"""User settings for the server: env var > settings file > built-in default.

Each setting has an environment variable (for people who configure the server in
Claude's config file) and a key in a small JSON settings file that the ``set_setting``
tool writes, so a user can change a setting by asking Claude and it sticks across
restarts and ``git pull``. The file lives outside the repo:

    %LOCALAPPDATA%\\outlook-desktop-mcp\\settings.json   (or OUTLOOK_MCP_SETTINGS_FILE)

Values are read on every lookup (the file is re-read when its mtime changes), so a
change takes effect immediately in the running server. An environment variable always
wins over the file; ``get_with_source`` says which one applied.
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
    "open_drafts": {
        "env": "OUTLOOK_MCP_OPEN_DRAFTS", "type": "bool", "default": False,
        "help": "Open each new draft in an Outlook window. Off = just save it to Drafts.",
    },
    "font_family": {"env": "OUTLOOK_MCP_FONT_FAMILY", "type": "str", "default": "Aptos",
                    "help": "Font for text the server writes into drafts."},
    "font_size_pt": {"env": "OUTLOOK_MCP_FONT_SIZE_PT", "type": "str", "default": "12",
                     "help": "Font size in points for draft text."},
    "font_color": {"env": "OUTLOOK_MCP_FONT_COLOR", "type": "str", "default": "#000000",
                   "help": "Font colour for draft text."},
    "plain_text": {"env": "OUTLOOK_MCP_PLAIN_TEXT", "type": "bool", "default": False,
                   "help": "Write drafts as plain text instead of styled HTML."},
}

_lock = threading.Lock()
_cache: dict[str, Any] = {"path": None, "mtime": None, "data": {}}


def settings_path() -> pathlib.Path:
    override = os.environ.get("OUTLOOK_MCP_SETTINGS_FILE", "").strip()
    if override:
        return pathlib.Path(override)
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CONFIG_HOME") or str(pathlib.Path.home() / ".config")
    return pathlib.Path(base) / "outlook-desktop-mcp" / "settings.json"


def _coerce(key: str, raw: Any) -> Any:
    kind = SETTINGS[key]["type"]
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
    coerced = _coerce(key, value)
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
