"""Exception hierarchy shared by core and tools. UI layers map these to messages."""

from __future__ import annotations


class M365Error(Exception):
    """Base class for all toolkit errors."""


class ConfigError(M365Error):
    """Missing or invalid configuration (tenant, profile, paths)."""


class AuthError(M365Error):
    """Token acquisition failed or was cancelled."""


class NotFoundError(M365Error):
    """Site, library, folder, file or user not found."""

    def __init__(self, message: str, *, available: list[str] | None = None) -> None:
        super().__init__(message)
        self.available = available or []


class AccessDeniedError(M365Error):
    """403 from SharePoint / Graph."""


class ThrottledError(M365Error):
    """Gave up after retrying 429/503 responses."""

    def __init__(self, message: str, *, retry_after: float | None = None) -> None:
        super().__init__(message)
        self.retry_after = retry_after


class TransferError(M365Error):
    """Download/upload failed after retries."""


class ValidationError(M365Error):
    """A plan, option or user map is invalid."""
