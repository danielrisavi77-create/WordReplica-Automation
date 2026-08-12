from enum import StrEnum


class RendererChoice(StrEnum):
    AUTO = "auto"
    WORD = "word"
    DOCX = "docx"


class VisibilityMode(StrEnum):
    VISIBLE = "visible"
    BACKGROUND = "background"


class FidelityMode(StrEnum):
    CLEAN = "clean"
    FULL = "full"


class MetadataMode(StrEnum):
    FRESH = "fresh"
    PRESERVE = "preserve"


class RunStatus(StrEnum):
    PASS = "PASS"
    WARN = "WARN"
    FAIL = "FAIL"


class ReconstructionMode(StrEnum):
    INSTANT = "instant"
    INTERACTIVE = "interactive"


class InteractiveSpeedMode(StrEnum):
    SLOW = "slow"
    FAST = "fast"
    CUSTOM = "custom"
    MAXIMUM = "maximum"


class InteractiveFidelity(StrEnum):
    STANDARD = "standard"
    MAXIMUM = "maximum"


class CapabilityClass(StrEnum):
    RECONSTRUCTED = "RECONSTRUCTED"
    PRESERVED = "PRESERVED"
    UNSUPPORTED = "UNSUPPORTED"


class InteractiveRunState(StrEnum):
    CREATED = "CREATED"
    PREFLIGHT = "PREFLIGHT"
    RUNNING = "RUNNING"
    PAUSED = "PAUSED"
    STOPPED = "STOPPED"
    VERIFYING = "VERIFYING"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
