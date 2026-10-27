"""Platform-aware entry point for outlook-desktop-mcp.

Detects the OS and imports the correct server module:
- macOS → server_mac (AppleScript automation)
- Windows → server (COM automation)

Flags:
    --read-only   register only tools that read; no send/reply/move/create/delete tools.
                  Same as setting OUTLOOK_MCP_READ_ONLY=1.
"""
import argparse
import sys

from outlook_desktop_mcp.readonly import ENV_VAR, enable_read_only


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="outlook-desktop-mcp",
        description="Expose the locally running Outlook Desktop as an MCP server over stdio.",
    )
    parser.add_argument(
        "--read-only", action="store_true",
        help=f"expose read tools only (no send, reply, move, mark, create, update or delete); same as {ENV_VAR}=1",
    )
    args, _unknown = parser.parse_known_args(argv)
    if args.read_only:
        enable_read_only()  # must happen before the server module registers its tools

    if sys.platform == "darwin":
        from outlook_desktop_mcp.server_mac import main as _main
    else:
        from outlook_desktop_mcp.server import main as _main
    _main()
