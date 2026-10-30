"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")

DISP_E_EXCEPTION = 0x80020009

NO_DETAIL_HINT = (
    "Outlook gave no detail. Common causes: the item is open in an Outlook window or selected in "
    "the reading pane, or the file is in use by another program."
)


def _hex(code) -> str:
    return f"0x{int(code) & 0xFFFFFFFF:08X}"


def format_com_error(e: Exception) -> str:
    """Outlook's own words for a COM failure.

    pywin32 reports every exception Outlook raises as DISP_E_EXCEPTION, "Exception occurred.";
    the sentence Outlook wrote ("This method can't be used with an inline response mail item.")
    and its real error code sit in ``excepinfo``. That sentence is what the caller can act on,
    so it leads and the code follows.
    """
    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            description = ""
            scode = None
            if exc:
                # (wCode, source, description, helpFile, helpContext, scode)
                description = str(exc[2] or "").strip()
                scode = exc[5] if len(exc) > 5 else None
            if description:
                return f"{description.rstrip('.')}. (Outlook error {_hex(scode or hr)})"
            text = f"COM Error ({_hex(hr)}): {msg}"
            if (int(hr) & 0xFFFFFFFF) == DISP_E_EXCEPTION:
                text += " " + NO_DETAIL_HINT
            return text
    except Exception:
        pass
    # Non-COM errors (a bad argument, an unknown calendar, a parse failure) are the
    # caller's to fix, so say what went wrong instead of a generic message.
    _logger.warning("Non-COM exception: %s: %s", type(e).__name__, e)
    return f"{type(e).__name__}: {e}"
