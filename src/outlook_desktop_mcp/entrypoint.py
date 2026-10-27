"""Entry point for outlook-desktop-mcp (Windows only)."""
import sys


def main() -> None:
    if "--version" in sys.argv[1:] or "-V" in sys.argv[1:]:
        from outlook_desktop_mcp import __version__
        print(f"outlook-desktop-mcp {__version__}")
        return
    if sys.platform != "win32":
        sys.stderr.write(
            "outlook-desktop-mcp only runs on Windows: it drives the classic Outlook desktop "
            "app through COM, which exists nowhere else.\n"
        )
        sys.exit(1)

    from outlook_desktop_mcp.server import main as _main
    _main()
