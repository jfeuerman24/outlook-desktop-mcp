"""Register (or remove) outlook-desktop-mcp in the Claude desktop app's settings file.

The Claude desktop app keeps its local MCP servers in
``%APPDATA%\\Claude\\claude_desktop_config.json``. This script adds or updates
one entry, ``mcpServers.outlook-desktop``, and leaves every other line of the
file exactly as it was. Before writing it saves a timestamped backup next to
the original. If the file is not valid JSON it stops and changes nothing.

Usage (normally run by install.ps1, not by hand):

    python scripts/configure_claude_desktop.py            # add / update
    python scripts/configure_claude_desktop.py --remove   # take it out again
    python scripts/configure_claude_desktop.py --show     # print the file path and entry only

Exit codes: 0 done (or nothing to do), 1 bad arguments or missing exe,
2 settings file unreadable, 3 Claude desktop app folder not found.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import pathlib
import shutil
import sys

SERVER_NAME = "outlook-desktop"


def default_exe() -> pathlib.Path:
    repo = pathlib.Path(__file__).resolve().parent.parent
    return repo / ".venv" / "Scripts" / "outlook-desktop-mcp.exe"


def candidate_dirs() -> list[pathlib.Path]:
    """Folders where the Claude desktop app may keep claude_desktop_config.json, best first.

    - Microsoft Store build: %LOCALAPPDATA%\\Claude-Data\\<package id>\\ (this is where its
      Settings > Developer > Edit Config button opens).
    - Download from claude.ai: %APPDATA%\\Claude\\.
    - Older Store builds: %LOCALAPPDATA%\\Packages\\<package>\\LocalCache\\Roaming\\Claude\\.
    """
    out: list[pathlib.Path] = []
    local = os.environ.get("LOCALAPPDATA")
    appdata = os.environ.get("APPDATA")
    if local:
        data = pathlib.Path(local) / "Claude-Data"
        if data.is_dir():
            out.extend(sorted(p for p in data.iterdir() if p.is_dir()))
            out.append(data)
    if appdata:
        out.append(pathlib.Path(appdata) / "Claude")
    if local:
        pkgs = pathlib.Path(local) / "Packages"
        if pkgs.is_dir():
            out.extend(sorted(pkgs.glob("*/LocalCache/Roaming/Claude")))
    return out


def default_config() -> pathlib.Path | None:
    """Pick the settings file the running Claude app actually uses.

    Candidates are in priority order. The first one that already holds the file wins;
    failing that, the first folder that exists. None if no candidate folder exists.
    """
    dirs = [d for d in candidate_dirs() if d.is_dir()]
    if not dirs:
        return None
    for d in dirs:
        if (d / "claude_desktop_config.json").exists():
            return d / "claude_desktop_config.json"
    return dirs[0] / "claude_desktop_config.json"


def load(path: pathlib.Path) -> dict:
    """Return the parsed file, or {} when it is missing or blank."""
    if not path.exists() or path.stat().st_size == 0:
        return {}
    text = path.read_text(encoding="utf-8-sig")
    if not text.strip():
        return {}
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("top level of the settings file is not a JSON object")
    return data


def backup(path: pathlib.Path) -> pathlib.Path | None:
    if not path.exists():
        return None
    stamp = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = path.with_name(f"claude_desktop_config.backup-{stamp}.json")
    n = 1
    while dest.exists():
        n += 1
        dest = path.with_name(f"claude_desktop_config.backup-{stamp}-{n}.json")
    shutil.copy2(path, dest)
    return dest


def write(path: pathlib.Path, data: dict) -> None:
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    os.replace(tmp, path)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--exe", help="path to outlook-desktop-mcp.exe (default: this repo's .venv)")
    ap.add_argument("--config", help="settings file to edit (default: %%APPDATA%%\\Claude\\claude_desktop_config.json)")
    ap.add_argument("--remove", action="store_true", help="remove the entry instead of adding it")
    ap.add_argument("--show", action="store_true", help="print the paths and current entry, change nothing")
    args = ap.parse_args(argv)

    cfg = pathlib.Path(args.config) if args.config else default_config()
    if cfg is None:
        looked = "\n  ".join(str(d) for d in candidate_dirs()) or "(APPDATA/LOCALAPPDATA not set; is this Windows?)"
        print("Could not find the Claude desktop app's settings folder. Looked in:\n  " + looked +
              "\nInstall the Claude desktop app, open it once, and run this again. If the app keeps its settings "
              "somewhere else (Settings > Developer > Edit Config opens that folder), pass --config <path>.",
              file=sys.stderr)
        return 3
    if not cfg.parent.exists():
        print(f"{cfg.parent} does not exist.", file=sys.stderr)
        return 3

    exe = pathlib.Path(args.exe).resolve() if args.exe else default_exe()

    try:
        data = load(cfg)
    except (ValueError, json.JSONDecodeError) as e:
        print(f"{cfg} is not valid JSON ({e}). Nothing was changed.\n"
              f"Open it from Claude's Settings > Developer > Edit Config, fix or clear it, and run this again.",
              file=sys.stderr)
        return 2

    servers = data.setdefault("mcpServers", {})
    if not isinstance(servers, dict):
        print(f'"mcpServers" in {cfg} is not an object. Nothing was changed.', file=sys.stderr)
        return 2

    current = servers.get(SERVER_NAME)

    if args.show:
        print(f"settings file: {cfg}")
        print(f"entry {SERVER_NAME!r}: {json.dumps(current) if current is not None else '(not present)'}")
        print(f"exe: {exe} ({'exists' if exe.exists() else 'MISSING'})")
        return 0

    if args.remove:
        if current is None:
            print(f"{SERVER_NAME} is not in {cfg}; nothing to do.")
            return 0
        bak = backup(cfg)
        del servers[SERVER_NAME]
        write(cfg, data)
        print(f"Removed {SERVER_NAME} from {cfg}" + (f" (backup: {bak.name})" if bak else ""))
        return 0

    if not exe.exists():
        print(f"{exe} does not exist. Run install.ps1 first.", file=sys.stderr)
        return 1

    desired = {"command": str(exe)}
    if current == desired:
        print(f"{SERVER_NAME} is already registered in {cfg}; nothing to do.")
        return 0

    bak = backup(cfg)
    servers[SERVER_NAME] = desired
    write(cfg, data)
    verb = "Updated" if current is not None else "Added"
    print(f"{verb} {SERVER_NAME} in {cfg}")
    print(f"  command: {exe}")
    if bak:
        print(f"  backup of the previous file: {bak.name}")
    print("Quit the Claude desktop app from the system tray and start it again for this to take effect.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
