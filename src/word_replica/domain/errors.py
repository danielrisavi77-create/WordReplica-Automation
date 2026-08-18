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


class RepairPackageError(WordReplicaError):
    """Raised when a Lekta repair-contract package fails preflight validation.

    Always fail-closed: raised before Word is ever started.
    """

    def __init__(self, code: str, message: str = "") -> None:
        super().__init__(f"{code}: {message}" if message else code)
        self.code = code
