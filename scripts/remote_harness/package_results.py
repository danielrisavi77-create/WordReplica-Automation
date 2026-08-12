from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

_ARTIFACT_EXTS={".docx",".pdf",".png",".jpg",".jpeg",".webp"}


def package_results(result_root: Path, destination_zip: Path, logs_only: bool) -> Path:
    result_root=Path(result_root).resolve(); destination_zip=Path(destination_zip).resolve()
    destination_zip.parent.mkdir(parents=True,exist_ok=True)
    if destination_zip.exists(): destination_zip.unlink()
    prefix=Path("remote_results")/result_root.name
    with ZipFile(destination_zip,"w",ZIP_DEFLATED,allowZip64=True) as zf:
        for path in sorted(result_root.rglob("*")):
            if not path.is_file():
                continue
            if path.resolve()==destination_zip:
                continue
            if logs_only and path.suffix.lower() in _ARTIFACT_EXTS:
                continue
            rel=path.relative_to(result_root)
            if any(part.lower() in {"realworld_input","source_input"} for part in rel.parts):
                continue
            zf.write(path, (prefix/rel).as_posix())
    return destination_zip
