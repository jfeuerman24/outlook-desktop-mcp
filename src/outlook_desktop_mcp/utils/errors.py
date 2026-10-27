"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")


def format_com_error(e: Exception) -> str:
    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            details = exc[2] if exc else "No details"
            _logger.debug("COM error detail: %s", details)  # internal only
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}"
    except Exception:
        pass
    # Non-COM errors (a bad argument, an unknown calendar, a parse failure) are the
    # caller's to fix, so say what went wrong instead of a generic message.
    _logger.warning("Non-COM exception: %s: %s", type(e).__name__, e)
    return f"{type(e).__name__}: {e}"
