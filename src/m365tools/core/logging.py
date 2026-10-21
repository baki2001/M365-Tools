"""Run logging: ``run.log`` in the run folder plus a handler that forwards records to the UI as ``LogLine``s.

There is deliberately no console handler; the TUI/CLI render ``LogLine`` events themselves.
"""

from __future__ import annotations

import logging
from pathlib import Path

from m365tools.core.events import EventSink, Level, LogLine

LOGGER_NAME = "m365tools"
LOG_FORMAT = "%(asctime)s [%(levelname)-5s] %(message)s"

_LEVEL_MAP: dict[int, Level] = {
    logging.DEBUG: "debug",
    logging.INFO: "info",
    logging.WARNING: "warn",
    logging.ERROR: "error",
    logging.CRITICAL: "error",
}


def _to_level(levelno: int) -> Level:
    if levelno >= logging.ERROR:
        return "error"
    if levelno >= logging.WARNING:
        return "warn"
    if levelno >= logging.INFO:
        return "info"
    return _LEVEL_MAP.get(levelno, "debug")


class SinkHandler(logging.Handler):
    """Converts log records into :class:`LogLine` events for an :class:`EventSink`."""

    def __init__(self, sink: EventSink, level: int | str = logging.INFO) -> None:
        super().__init__(level)
        self.sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.sink(LogLine(level=_to_level(record.levelno), text=record.getMessage()))
        except Exception:  # noqa: BLE001 - a broken UI must never break the run
            self.handleError(record)


def setup_logging(run_dir: Path, level: int | str = "INFO") -> logging.Logger:
    """Configure the ``m365tools`` logger to write ``run_dir/run.log``.

    Previous file handlers installed by this function are removed, so calling it for each run is safe.
    Attach a :class:`SinkHandler` separately to forward records to the UI.
    """
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)
    logger.propagate = False
    for h in list(logger.handlers):
        if getattr(h, "_m365_run_file", False):
            logger.removeHandler(h)
            h.close()
    run_dir.mkdir(parents=True, exist_ok=True)
    fh = logging.FileHandler(run_dir / "run.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter(LOG_FORMAT))
    fh.setLevel(level)
    fh._m365_run_file = True  # type: ignore[attr-defined]
    logger.addHandler(fh)
    return logger
