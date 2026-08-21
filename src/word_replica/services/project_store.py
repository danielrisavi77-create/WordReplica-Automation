from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
import sqlite3
import uuid

from platformdirs import user_documents_dir

from word_replica.config import RebuildOptions
from word_replica.services.source_guard import capture_source


@dataclass(frozen=True, slots=True)
class ProjectPaths:
    project_id: str
    root: Path
    source_snapshot_dir: Path
    working_dir: Path
    output_dir: Path
    backups_dir: Path
    logs_dir: Path
    qa_dir: Path
    project_json: Path


class ProjectStore:
    def __init__(
        self, app_root: Path | None = None, *, projects_under_app_root: bool = False
    ) -> None:
        """``projects_under_app_root`` keeps projects out of the source's tree.

        A project normally lands beside the document it came from, which is
        where someone rebuilding their own file expects to find it. A caller
        reading a corpus it does not own -- the fidelity lab -- needs the
        opposite, or every pass leaves a full project next to every document it
        read. The default is unchanged.
        """
        self.app_root = app_root or Path(user_documents_dir()) / "WordReplica" / "Projects"
        self.app_root.mkdir(parents=True, exist_ok=True)
        self.projects_under_app_root = projects_under_app_root
        self.db_path = self.app_root / "projects.sqlite3"
        self._init_db()

    def _init_db(self) -> None:
        with sqlite3.connect(self.db_path) as db:
            db.execute(
                "CREATE TABLE IF NOT EXISTS projects ("
                "id TEXT PRIMARY KEY, source_path TEXT NOT NULL, root_path TEXT NOT NULL, "
                "created_utc TEXT NOT NULL, status TEXT NOT NULL)"
            )

    def create_project(self, source: Path, options: RebuildOptions) -> ProjectPaths:
        source = source.resolve()
        snapshot = capture_source(source)
        project_id = uuid.uuid4().hex
        # project_id is unique per call, so two documents of the same name from
        # different directories cannot collide under a shared root.
        home = self.app_root / "projects" if self.projects_under_app_root else source.parent
        root = home / f"{source.stem}_rebuild" / project_id
        dirs = {
            name: root / name
            for name in ("source_snapshot", "working", "output", "backups", "logs", "qa")
        }
        for path in dirs.values():
            path.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, dirs["source_snapshot"] / source.name)

        project_json = root / "project.json"
        payload = {
            "project_id": project_id,
            "source_path": str(source),
            "source_sha256": snapshot.sha256,
            "source_size": snapshot.size,
            "options": asdict(options),
        }
        project_json.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        created_utc = datetime.now(timezone.utc).isoformat()
        with sqlite3.connect(self.db_path) as db:
            db.execute(
                "INSERT INTO projects VALUES (?, ?, ?, ?, ?)",
                (project_id, str(source), str(root), created_utc, "CREATED"),
            )
        return ProjectPaths(
            project_id,
            root,
            dirs["source_snapshot"],
            dirs["working"],
            dirs["output"],
            dirs["backups"],
            dirs["logs"],
            dirs["qa"],
            project_json,
        )

    def get_project(self, project_id: str) -> dict[str, str]:
        with sqlite3.connect(self.db_path) as db:
            row = db.execute(
                "SELECT id, source_path, root_path, created_utc, status FROM projects WHERE id = ?",
                (project_id,),
            ).fetchone()
        if row is None:
            raise KeyError(project_id)
        return dict(zip(("id", "source_path", "root_path", "created_utc", "status"), row, strict=True))

    def paths_for_project(self, project_id: str) -> ProjectPaths:
        project = self.get_project(project_id)
        root = Path(project["root_path"])
        return ProjectPaths(
            project_id=project_id,
            root=root,
            source_snapshot_dir=root / "source_snapshot",
            working_dir=root / "working",
            output_dir=root / "output",
            backups_dir=root / "backups",
            logs_dir=root / "logs",
            qa_dir=root / "qa",
            project_json=root / "project.json",
        )

    def list_resumable_projects(self) -> list[dict[str, object]]:
        resumable: list[dict[str, object]] = []
        for project in self.list_projects():
            if project["status"] not in {"STOPPED", "PAUSED", "INTERRUPTED"}:
                continue
            paths = self.paths_for_project(project["id"])
            checkpoint_path = paths.logs_dir / "interactive_checkpoint.json"
            if not checkpoint_path.exists():
                continue
            try:
                checkpoint = json.loads(checkpoint_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            item: dict[str, object] = dict(project)
            item.update(
                checkpoint_path=str(checkpoint_path),
                last_completed_event_index=int(checkpoint.get("last_completed_event_index", -1)),
                checkpoint_timestamp_utc=str(checkpoint.get("timestamp_utc", "")),
                checkpoint_status=str(checkpoint.get("status", project["status"])),
            )
            resumable.append(item)
        return resumable

    def set_status(self, project_id: str, status: str) -> None:
        with sqlite3.connect(self.db_path) as db:
            db.execute("UPDATE projects SET status = ? WHERE id = ?", (status, project_id))

    def list_projects(self) -> list[dict[str, str]]:
        with sqlite3.connect(self.db_path) as db:
            rows = db.execute(
                "SELECT id, source_path, root_path, created_utc, status "
                "FROM projects ORDER BY created_utc DESC"
            ).fetchall()
        keys = ("id", "source_path", "root_path", "created_utc", "status")
        return [dict(zip(keys, row, strict=True)) for row in rows]
