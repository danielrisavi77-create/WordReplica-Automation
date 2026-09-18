"""Reading lxml nodes that are not elements.

An lxml tree contains comments and processing instructions alongside elements,
and their ``.tag`` is a *callable* (``etree.Comment``), not a string. Code that
does ``node.tag.rsplit("}", 1)`` to get a local name crashes on them with an
AttributeError -- which is how a real document (Apache POI's
``Bug66263-table.docx``, whose body opens with a Word-generated conditional
comment) took the parser down.
"""
from __future__ import annotations

__all__ = ["local_name"]


def local_name(node) -> str:
    """Namespace-stripped tag name, or "" for comments and processing instructions.

    Returning "" rather than raising is deliberate: every caller dispatches on
    the local name, so a non-element simply matches no branch and is skipped.
    """
    tag = getattr(node, "tag", None)
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]
