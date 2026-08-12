from __future__ import annotations

import csv
from collections import defaultdict
from datetime import datetime, timezone
import json
from pathlib import Path

from scripts.remote_harness.contracts import normalize_failure


def _load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"status":"HARNESS_FAIL","reasons":[f"unreadable result.json: {exc}"]}


def aggregate_results(result_root: Path) -> dict:
    result_root=Path(result_root)
    rows=[]
    groups: dict[str, dict] = {}
    document_dirs=sorted((result_root/"documents").glob("*")) if (result_root/"documents").exists() else []
    for doc_dir in document_dirs:
        if not doc_dir.is_dir():
            continue
        for stage_dir in sorted(p for p in doc_dir.iterdir() if p.is_dir()):
            result_path=stage_dir/"result.json"
            if not result_path.exists():
                continue
            payload=_load_json(result_path)
            document=str(payload.get("document") or doc_dir.name)
            stage=str(payload.get("stage") or stage_dir.name)
            status=str(payload.get("status") or "HARNESS_FAIL")
            reasons=payload.get("reasons") or []
            reason="; ".join(str(x) for x in reasons[:5])
            fingerprint=""
            if status not in {"PASS","WARN","EXPECTED_BLOCK"} or reason:
                fingerprint=normalize_failure(reason or status)
                group=groups.setdefault(fingerprint,{"runs":0,"documents_set":set(),"stages":set(),"examples":[]})
                group["runs"] += 1
                group["documents_set"].add(document)
                group["stages"].add(stage)
                if reason and len(group["examples"]) < 3:
                    group["examples"].append(reason)
            rows.append({
                "document":document,
                "stage":stage,
                "status":status,
                "elapsed_seconds":payload.get("elapsed_seconds",0),
                "failure_fingerprint":fingerprint,
                "reason":reason,
                "project_id":payload.get("project_id"),
                "save_count":payload.get("save_count",0),
            })
    failure_groups={
        key:{
            "runs":value["runs"],
            "documents":len(value["documents_set"]),
            "document_names":sorted(value["documents_set"]),
            "stages":sorted(value["stages"]),
            "examples":value["examples"],
        }
        for key,value in sorted(groups.items(), key=lambda item:(-item[1]["runs"],item[0]))
    }
    summary={
        "schema_version":1,
        "generated_at_utc":datetime.now(timezone.utc).isoformat(),
        "documents":len({row["document"] for row in rows}),
        "runs":len(rows),
        "status_counts":{},
        "failure_groups":failure_groups,
        "rows":rows,
    }
    for row in rows:
        summary["status_counts"][row["status"]]=summary["status_counts"].get(row["status"],0)+1
    result_root.mkdir(parents=True,exist_ok=True)
    (result_root/"summary.json").write_text(json.dumps(summary,indent=2,ensure_ascii=False),encoding="utf-8")
    with (result_root/"summary.csv").open("w",newline="",encoding="utf-8-sig") as handle:
        fieldnames=["document","stage","status","elapsed_seconds","failure_fingerprint","reason","project_id","save_count"]
        writer=csv.DictWriter(handle,fieldnames=fieldnames)
        writer.writeheader(); writer.writerows(rows)
    return summary
