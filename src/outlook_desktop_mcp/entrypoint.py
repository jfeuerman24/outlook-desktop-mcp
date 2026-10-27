"""Entry point for outlook-desktop-mcp (Windows only)."""
import sys


def main() -> None:
    if sys.platform != "win32":
        sys.stderr.write(
            "outlook-desktop-mcp only runs on Windows: it drives the classic Outlook desktop "
            "app through COM, which exists nowhere else.\n"
        )
        sys.exit(1)

    from outlook_desktop_mcp.server import main as _main
    _main()
