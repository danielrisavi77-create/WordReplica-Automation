class WordReplicaError(Exception):
    """Base exception for Word Replica."""


class CriticalRebuildError(WordReplicaError):
    """A reconstruction error that must stop the run."""


class SourceIntegrityError(CriticalRebuildError):
    """Raised when the protected source changes during a run."""


class RendererUnavailableError(CriticalRebuildError):
    """Raised when the requested renderer is unavailable."""


class PackageReadError(CriticalRebuildError):
    """Raised when the source DOCX package cannot be read."""
