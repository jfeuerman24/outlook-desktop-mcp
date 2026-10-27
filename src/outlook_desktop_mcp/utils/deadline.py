"""Cooperative deadlines for long Outlook loops. Pure Python, unit-tested.

The COM bridge gives each tool call a hard timeout, after which the client gives up. A
loop that keeps running past that point does nobody any good: the result is thrown away
and the next call queues behind it. So the bridge also sets a *soft* deadline a few
seconds earlier on the COM thread, and long loops check ``expired()`` and return what
they have, marked partial, while the client is still listening.
"""
from __future__ import annotations

import threading
import time


class Deadline:
    """A point in ``time.monotonic()`` seconds, or never when ``at`` is None."""

    def __init__(self, seconds: float | None = None, at: float | None = None):
        if at is not None:
            self.at = at
        elif seconds is None:
            self.at = None
        else:
            self.at = time.monotonic() + max(0.0, float(seconds))

    @property
    def expired(self) -> bool:
        return self.at is not None and time.monotonic() >= self.at

    def remaining(self) -> float | None:
        if self.at is None:
            return None
        return max(0.0, self.at - time.monotonic())


NEVER = Deadline(None)
_local = threading.local()


def set_current(deadline: Deadline | None) -> None:
    """Install the deadline for the running thread (the bridge does this per call)."""
    _local.current = deadline


def current() -> Deadline:
    return getattr(_local, "current", None) or NEVER


def expired() -> bool:
    """True once the current call should stop and return partial results."""
    return current().expired


def partial_note(what: str = "The search") -> str:
    return (f"{what} stopped early to answer within the time limit; results may be incomplete. "
            "Narrow it (one account, a date range, a more specific name) or raise the call_timeout_s setting.")
