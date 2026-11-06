"""Stable error taxonomy shared by the library, CLI, and MCP surface."""

from __future__ import annotations


class ResearchMCPError(RuntimeError):
    """Base class for expected product errors."""


class DocumentNotFound(ResearchMCPError):
    pass


class AmbiguousDocument(ResearchMCPError):
    pass


class RefusedUnsafe(ResearchMCPError):
    pass


class VerificationFailed(ResearchMCPError):
    def __init__(self, message: str, *, expected=None, actual=None):
        super().__init__(message)
        self.expected = expected
        self.actual = actual


class DocumentDamaged(ResearchMCPError):
    pass


class ZoteroUnavailable(ResearchMCPError):
    pass


class ZoteroModal(ResearchMCPError):
    pass


class ZoteroItemNotFound(ResearchMCPError):
    pass


class IntegrationBusy(ResearchMCPError):
    pass


class IntegrationRefused(ResearchMCPError):
    pass
