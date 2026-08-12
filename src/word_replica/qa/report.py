from __future__ import annotations

from pathlib import Path
from typing import Any

from jinja2 import Environment, select_autoescape

_TEMPLATE = """<!doctype html>
<html><head><meta charset='utf-8'><title>Word Replica QA — {{ status }}</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:2rem;line-height:1.45;max-width:1100px}
code{overflow-wrap:anywhere}.PASS{font-weight:700}.finding,.warning{margin:.4rem 0;padding:.4rem .6rem;background:#f5f5f5;border-radius:.35rem}
table{border-collapse:collapse}td,th{border:1px solid #ddd;padding:.35rem .55rem;text-align:left}
</style></head><body>
<h1>Word Replica QA — {{ status }}</h1>
{% if source_sha256 %}<p>Source SHA-256: <code>{{ source_sha256 }}</code></p>{% endif %}
{% if output_sha256 %}<p>Output SHA-256: <code>{{ output_sha256 }}</code></p>{% endif %}
{% if renderer or fidelity or metadata %}<p>Renderer: {{ renderer or '-' }} | Fidelity: {{ fidelity or '-' }} | Metadata: {{ metadata or '-' }}</p>{% endif %}
<h2>Actual saves ({{ saves|length }})</h2><ol>{% for save in saves %}<li>#{{ value(save, 'sequence') }} — {{ value(save, 'reason') }} — {{ value(save, 'timestamp_local') }}</li>{% endfor %}</ol>
<h2>QA</h2>{% for level, result in levels.items() %}<section><h3>{{ level }} — {{ 'PASS' if value(result, 'passed') else 'DIFF' }}</h3>{% for f in value(result, 'findings', []) %}<div class='finding'>{{ value(f, 'code', 'FINDING') }} @ {{ value(f, 'path', '-') }}</div>{% endfor %}</section>{% endfor %}
<h2>L4 visual comparison</h2>{% if render_result %}<p>{{ 'Within tolerance' if value(render_result, 'within_tolerance') else 'Outside tolerance' }}; pages {{ value(render_result, 'source_page_count', 0) }}/{{ value(render_result, 'rebuilt_page_count', 0) }}.</p>{% else %}<p>Not verified in this run.</p>{% endif %}
<h2>Warnings / limitations</h2>{% for w in warnings %}<div class='warning'>{{ value(w, 'code') }} — {{ value(w, 'message') }}</div>{% else %}<p>None.</p>{% endfor %}
<p><strong>Integrity:</strong> This document was reconstructed by an automated process. Interactive mode may insert content character by character and object by object, but that is not proof of manual authorship. Word Replica does not fabricate timestamps, editing time, save history, or provenance.</p>
</body></html>"""


def _value(obj: Any, name: str, default: Any = "") -> Any:
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def write_qa_report(path: Path, **context: Any) -> Path:
    env = Environment(autoescape=select_autoescape(default_for_string=True))
    env.globals["value"] = _value
    defaults = {
        "status": "UNKNOWN",
        "source_sha256": "",
        "output_sha256": "",
        "renderer": "",
        "fidelity": "",
        "metadata": "",
        "saves": [],
        "levels": {},
        "warnings": [],
        "render_result": None,
    }
    defaults.update(context)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(env.from_string(_TEMPLATE).render(**defaults), encoding="utf-8")
    return path
