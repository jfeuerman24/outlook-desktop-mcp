"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.

Timeouts
--------
Each call has a hard timeout (the ``call_timeout_s`` setting, or ``_timeout=``).
Outlook work cannot be interrupted, so three things keep one slow call from
taking the whole session down with it:

* a *soft* deadline a few seconds before the hard one is installed on the COM
  thread (``utils.deadline``); long loops check it and return partial results
  while the client is still listening;
* a request whose caller has already given up is skipped when its turn comes,
  instead of running for nothing;
* a new call made while an earlier one has overrun its own deadline fails at
  once with a clear "Outlook is busy" message rather than waiting to time out.
"""
import threading
import queue
import asyncio
import sys
import logging
import time

from outlook_desktop_mcp.utils import deadline as _deadline
from outlook_desktop_mcp.utils import settings

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

DEFAULT_TIMEOUT = 60.0
MIN_TIMEOUT, MAX_TIMEOUT = 10.0, 600.0
SOFT_MARGIN = 5.0  # seconds before the hard timeout at which loops stop


def default_timeout() -> float:
    try:
        value = float(settings.get("call_timeout_s"))
    except Exception:
        value = DEFAULT_TIMEOUT
    return min(MAX_TIMEOUT, max(MIN_TIMEOUT, value))


class _Request:
    __slots__ = ("func", "args", "kwargs", "event", "holder", "name", "enqueued", "deadline_at", "soft_at", "started")

    def __init__(self, func, args, kwargs, timeout: float):
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.event = threading.Event()
        self.holder: dict = {}
        self.name = getattr(func, "__name__", "call")
        self.enqueued = time.monotonic()
        self.deadline_at = self.enqueued + timeout
        self.soft_at = self.deadline_at - min(SOFT_MARGIN, timeout * 0.25)
        self.started: float | None = None


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._running: _Request | None = None

    def start(self):
        """Start the COM thread. Call once at server startup."""
        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=15):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                "Outlook COM thread failed to initialize within 15s. "
                "Is Outlook Desktop (Classic) running?"
            )

    def _com_thread_main(self):
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    req: _Request = self._request_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                now = time.monotonic()
                if now >= req.deadline_at:
                    # The caller stopped waiting while this sat behind a slow call.
                    req.holder["error"] = TimeoutError(
                        "Skipped: the request waited so long behind an earlier Outlook request that "
                        "its caller had already given up. Try again."
                    )
                    req.event.set()
                    continue
                req.started = now
                self._running = req
                _deadline.set_current(_deadline.Deadline(at=req.soft_at))
                try:
                    req.holder["value"] = req.func(self._outlook, self._namespace, *req.args, **req.kwargs)
                except Exception as e:
                    req.holder["error"] = e
                finally:
                    _deadline.set_current(None)
                    self._running = None
                    took = time.monotonic() - now
                    if took > req.soft_at - now:
                        logger.warning("%s ran %.1fs (limit %.0fs)", req.name, took, req.deadline_at - req.enqueued)
                    req.event.set()
        except Exception as e:
            self._init_error = e
            self._ready.set()  # Unblock the caller so they see the error
            logger.error("COM thread init failed: %s", e)
        finally:
            pythoncom.CoUninitialize()

    def busy_for(self) -> float | None:
        """Seconds the current call has been running past its own deadline, else None."""
        running = self._running
        if running is None or running.started is None:
            return None
        over = time.monotonic() - running.deadline_at
        return over if over > 0 else None

    async def call(self, func, *args, **kwargs):
        """
        Schedule a function to run on the COM thread and await its result.

        The function signature must be: func(outlook, namespace, *args, **kwargs)
        ``_timeout=`` overrides the call_timeout_s setting for this call.
        """
        timeout = kwargs.pop("_timeout", None)
        timeout = float(timeout) if timeout else default_timeout()

        over = self.busy_for()
        if over is not None:
            running = self._running
            name = running.name if running else "an earlier request"
            raise TimeoutError(
                f"Outlook is still busy finishing an earlier request ({name}) that has overrun its time "
                f"limit by {over:.0f}s. Try again in a moment, or make the request narrower (one account, "
                "a date range, a more specific name)."
            )

        req = _Request(func, args, kwargs, timeout)
        self._request_queue.put(req)

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(None, lambda: req.event.wait(timeout=timeout + 0.5))
        if not signaled:
            if req.started is None:
                raise TimeoutError(
                    f"Waited {timeout:.0f}s for an earlier Outlook request to finish and gave up. "
                    "Try again in a moment."
                )
            raise TimeoutError(
                f"Outlook COM operation timed out after {timeout:.0f} seconds. Outlook may be waiting "
                "for user input (e.g., a dialog box), or the request is too broad for the time limit."
            )

        if "error" in req.holder:
            raise req.holder["error"]
        return req.holder.get("value")

    def stop(self):
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
