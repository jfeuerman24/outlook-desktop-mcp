"""Resolve Exchange addresses to SMTP.

Inside an Exchange organisation ``SenderEmailAddress`` is a directory name such as
``/o=ExchangeLabs/ou=.../cn=...`` rather than an email address. MAPI stores the real
address in ``PR_SMTP_ADDRESS``; ``GetExchangeUser()`` is the slower fallback. Every
lookup is wrapped: an unreadable address yields "" rather than an exception.
"""
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

RECIPIENT_TYPES = {1: "to", 2: "cc", 3: "bcc"}


def _is_dn(address: str) -> bool:
    return address.lower().startswith("/o=")  # Exchange returns /O= or /o=


def get_smtp(item) -> str:
    """The sender's SMTP address, trying three sources in order."""
    try:
        addr = item.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if addr:
            return str(addr).lower()
    except Exception:
        pass
    try:
        addr = str(item.SenderEmailAddress or "")
        if addr and not _is_dn(addr):
            return addr.lower()
    except Exception:
        pass
    try:
        return str(item.Sender.GetExchangeUser().PrimarySmtpAddress).lower()
    except Exception:
        return ""


def recipient_address(recipient) -> str:
    try:
        addr = str(recipient.Address or "")
        if "@" in addr and not _is_dn(addr):
            return addr.lower()
    except Exception:
        pass
    try:
        addr = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if addr:
            return str(addr).lower()
    except Exception:
        pass
    try:
        return str(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress).lower()
    except Exception:
        return ""


def format_recipients(item) -> list[dict]:
    """``[{"name", "address", "type"}]`` for To/CC/BCC, addresses resolved to SMTP."""
    out: list[dict] = []
    try:
        recipients = item.Recipients
        for i in range(recipients.Count):
            recipient = recipients.Item(i + 1)
            try:
                out.append({
                    "name": str(recipient.Name or ""),
                    "address": recipient_address(recipient),
                    "type": RECIPIENT_TYPES.get(int(recipient.Type), "to"),
                })
            except Exception:
                continue
    except Exception:
        pass
    return out


def sender_matches(item, needle: str, smtp: str | None = None) -> bool:
    """Case-insensitive substring match of ``needle`` against the sender's address or name."""
    needle = needle.lower().strip()
    if not needle:
        return True
    address = smtp if smtp is not None else get_smtp(item)
    if needle in address:
        return True
    try:
        return needle in str(item.SenderName or "").lower()
    except Exception:
        return False
