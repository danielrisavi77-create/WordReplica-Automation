"""Compatibility re-export. The implementation now lives in the installable
package at word_replica.qa.golden_audit so it can be reused outside the
codex_automation harness (e.g. by the Lekta repair-package service).
"""
from __future__ import annotations

from word_replica.qa.golden_audit import (
    DEFAULT_GATE_NAMES,
    GateResult,
    audit_docx_pair,
    build_golden_report,
    build_model_gates,
    build_visual_gate,
    compare_page_text_partitions,
)

__all__ = [
    "DEFAULT_GATE_NAMES",
    "GateResult",
    "build_golden_report",
    "build_model_gates",
    "build_visual_gate",
    "audit_docx_pair",
    "compare_page_text_partitions",
]
