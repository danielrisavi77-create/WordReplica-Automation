from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict, replace
import hashlib
import json
import os
import posixpath
from pathlib import Path, PurePosixPath
import tempfile
import time
from typing import Any, Callable
from zipfile import ZIP_DEFLATED, ZipFile

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import InteractiveFidelity, InteractiveSpeedMode, MetadataMode, ReconstructionMode, RunStatus, VisibilityMode
from word_replica.domain.reconstruction import PreparedInteractiveRun
from word_replica.domain.results import RunResult, WarningItem
from word_replica.interactive.blueprint import BlueprintCompiler
from word_replica.interactive.capabilities import detect_image_extension
from word_replica.interactive.checkpoints import (
    InteractiveCheckpointCoordinator,
    InteractiveCheckpointStore,
)
from word_replica.interactive.control import InteractiveRunControl
from word_replica.interactive.preflight import analyze_preflight
from word_replica.interactive.verification import InteractiveProgressTracker, LiveVerifier
from word_replica.parser.parser import DocxParser
from word_replica.opc.properties import DocumentProperties, build_output_metadata
from word_replica.qa.environment import capture_environment_fingerprint
from word_replica.qa.policy import classify_run, run_l0_l3
from word_replica.qa.render import compare_pdfs
from word_replica.qa.report import write_qa_report
from word_replica.qa.word_render import export_docx_to_pdf_with_word
from word_replica.renderers.interactive_word import InteractiveWordController, InteractiveWordRenderer
from word_replica.services.audit import AuditLog
from word_replica.services.checkpoints import CheckpointManager
from word_replica.services.project_store import ProjectStore
from word_replica.services.source_guard import sha256_file


def _replace_with_retry(source: Path, destination: Path, *, attempts: int = 60, delay_seconds: float = 0.5) -> None:
    """Finalization runs a long chain of package-cleanup passes, each
    rewriting `destination` via a temp-file rename. Windows can transiently
    deny that rename right after a file is written - antivirus/indexer
    scanning is the common cause - even though no WordReplica-owned process
    still holds a handle on it. Retry briefly before giving up so a real
    permission problem still surfaces as an error.
    """
    last: OSError | None = None
    for attempt in range(attempts):
        try:
            source.replace(destination)
            return
        except OSError as exc:
            last = exc
            if attempt + 1 >= attempts:
                raise
            time.sleep(delay_seconds)
    if last is not None:
        raise last


class _CompositeObserver:
    def __init__(self, *observers) -> None:
        self.observers = tuple(observer for observer in observers if observer is not None)

    def event_started(self, index, event, snapshot) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_started", None)
            if callback is not None:
                callback(index, event, snapshot)

    def event_completed(self, index, event) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_completed", None)
            if callback is not None:
                callback(index, event)

    def event_finished(self, index, event, snapshot) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_finished", None)
            if callback is not None:
                callback(index, event, snapshot)

    def event_failed(self, index, event, snapshot, exc) -> None:
        for observer in self.observers:
            callback = getattr(observer, "event_failed", None)
            if callback is not None:
                callback(index, event, snapshot, exc)

    def table_batch_profile(self, index, event, metrics) -> None:
        for observer in self.observers:
            callback = getattr(observer, "table_batch_profile", None)
            if callback is not None:
                callback(index, event, metrics)

    @property
    def halt_status(self):
        for observer in self.observers:
            status = getattr(observer, "halt_status", None)
            if status:
                return status
        return None

    def state_mismatch(self, index, event, expected, actual) -> None:
        for observer in self.observers:
            callback = getattr(observer, "state_mismatch", None)
            if callback is not None:
                callback(index, event, expected, actual)


class _InteractiveServiceObserver:
    SAFE_VERIFY_EVENTS = {"EndTable", "InsertTableBatch", "SetImageZOrder", "EndSection"}

    def __init__(self, blueprint, control, audit, downstream=None, *, live_verifier=None,
                 source_model=None, output_path=None, coordinator=None, settings=None) -> None:
        self.blueprint = blueprint
        self.tracker = InteractiveProgressTracker(blueprint)
        self.control = control
        self.audit = audit
        self.downstream = downstream
        self.live_verifier = live_verifier
        self.source_model = source_model
        self.output_path = Path(output_path) if output_path is not None else None
        self.coordinator = coordinator
        self.settings = dict(settings or {})
        self.halt_status = None
        self.halt_reasons: tuple[str, ...] = ()
        self._current_story = "body"
        self._expected_body_text: list[str] = []

    def _track_expected_body_text(self, event) -> None:
        et = event.event_type
        if et == "BeginHeader": self._current_story = "header"
        elif et == "BeginFooter": self._current_story = "footer"
        elif et == "BeginFootnoteStory": self._current_story = "footnote"
        elif et == "BeginEndnoteStory": self._current_story = "endnote"
        elif et in {"EndHeader", "EndFooter", "EndFootnoteStory", "EndEndnoteStory"}: self._current_story = "body"
        elif et == "InsertCharacter" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("character", "")))
        elif et == "InsertText" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("text", "")))
        elif et == "InsertTableBatch" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("text_projection", "")))
        elif et == "InsertTab" and self._current_story == "body": self._expected_body_text.append("\t")
        elif et in {"InsertLineBreak", "InsertPageBreak"} and self._current_story == "body": self._expected_body_text.append("\n")
        elif et == "CreateField" and self._current_story == "body":
            self._expected_body_text.append(str(event.payload.get("cached_result", "")))

    def _verify_safe_boundary(self, index, event) -> None:
        if not self.settings.get("verify_during_run", False): return
        if event.event_type not in self.SAFE_VERIFY_EVENTS: return
        if self.live_verifier is None or self.output_path is None or self.source_model is None: return
        # Verification reads the saved DOCX. Force a truthful save when the configured
        # checkpoint policy did not already save this exact boundary.
        if self.coordinator is not None and self.coordinator.last_checkpoint_event_index != index:
            checkpoint = self.coordinator.milestone(index, "live_verification_boundary")
            self.tracker.checkpoint_saved(index, checkpoint.timestamp_utc)
        boundary = {
            "name": f"{event.event_type}:{event.source_element_id}",
            "expected_text_prefix": "".join(self._expected_body_text),
            "completed_tables": self.tracker.completed_tables,
        }
        result = self.live_verifier.verify_boundary(self.source_model, self.output_path, boundary)
        self.tracker.verification_updated(result.status)
        self.audit.append("INTERACTIVE_LIVE_VERIFICATION", {
            "event_index": index, "boundary": boundary["name"], "status": result.status,
            "passed": result.passed, "reasons": list(result.reasons),
        })
        if not result.passed:
            self.halt_status = "PAUSED_FIDELITY_MISMATCH"
            self.halt_reasons = tuple(result.reasons)
            self.control.pause()

    def prime_through(self, last_completed_index: int) -> None:
        if last_completed_index < 0:
            return
        end = min(last_completed_index, self.blueprint.total_events - 1)
        for index in range(0, end + 1):
            event = self.blueprint.events[index]
            self._track_expected_body_text(event)
            self.tracker.event_completed(index, event)

    def event_started(self, index, event, snapshot) -> None:
        callback = getattr(self.downstream, "event_started", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot)

    def event_failed(self, index, event, snapshot, exc) -> None:
        self.audit.append("INTERACTIVE_EVENT_FAILED", {
            "event_index": index,
            "event_type": event.event_type,
            "source_element_id": event.source_element_id,
            "error_type": type(exc).__name__,
            "error_hresult": getattr(exc, "hresult", None),
            "error": str(exc),
        })
        callback = getattr(self.downstream, "event_failed", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot, exc)

    def table_batch_profile(self, index, event, metrics) -> None:
        callback = getattr(self.downstream, "table_batch_profile", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, metrics)

    def event_finished(self, index, event, snapshot) -> None:
        callback = getattr(self.downstream, "event_finished", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, snapshot)

    def event_completed(self, index, event) -> None:
        self._track_expected_body_text(event)
        self.tracker.event_completed(index, event)
        self._verify_safe_boundary(index, event)
        callback = getattr(self.downstream, "event_completed", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event)
        progress = getattr(self.downstream, "progress", None) if self.downstream is not None else None
        if progress is not None:
            progress(self.tracker.snapshot(state=self.control.state))

    def state_mismatch(self, index, event, expected, actual) -> None:
        from dataclasses import asdict, is_dataclass
        left = asdict(expected) if is_dataclass(expected) else dict(expected)
        right = asdict(actual) if is_dataclass(actual) else dict(actual)
        self.tracker.verification_updated("WORD_STATE_MISMATCH")
        self.audit.append("WORD_STATE_MISMATCH", {
            "event_index": index, "event_type": event.event_type,
            "source_element_id": event.source_element_id,
            "expected": left, "actual": right,
        })
        callback = getattr(self.downstream, "state_mismatch", None) if self.downstream is not None else None
        if callback is not None:
            callback(index, event, expected, actual)
        progress = getattr(self.downstream, "progress", None) if self.downstream is not None else None
        if progress is not None:
            progress(self.tracker.snapshot(state=self.control.state))


class InteractiveRebuildService:
    def __init__(
        self,
        *,
        parser=None,
        word_probe: Callable[[], bool] | None = None,
        font_probe=None,
        project_store: ProjectStore | None = None,
        pdf_exporter=None,
        pdf_comparer=None,
        live_verifier=None,
        run_l4_qa: bool = True,
    ) -> None:
        self.parser = parser or DocxParser()
        self.word_probe = word_probe
        self.font_probe = font_probe
        self.project_store = project_store
        self.pdf_exporter = pdf_exporter or export_docx_to_pdf_with_word
        self.pdf_comparer = pdf_comparer or compare_pdfs
        self.live_verifier = live_verifier or LiveVerifier(self.parser)
        self.run_l4_qa = bool(run_l4_qa)

    def _parse_model_for_options(self, source: Path, options: RebuildOptions):
        # Output metadata policy is deliberately not part of the canonical source
        # fingerprint or reconstruction blueprint. It is applied after compilation.
        return self.parser.parse(Path(source)) if hasattr(self.parser, "parse") else self.parser(Path(source))

    @staticmethod
    def _apply_metadata_policy(model, options: RebuildOptions) -> None:
        source_props = model.extras.get("source_properties", DocumentProperties())
        model.extras["metadata_policy"] = build_output_metadata(
            source_props, options.metadata, options.preserve_author_fields, options.custom_metadata_allowlist
        )

    @staticmethod
    def _compile_blueprint(model, options: RebuildOptions):
        return BlueprintCompiler(
            enable_table_fast_path=options.interactive.enable_table_fast_path,
        ).compile(model)

    def prepare(self, source: Path, options: RebuildOptions, paths, audit, observer=None) -> PreparedInteractiveRun:
        model = self._parse_model_for_options(Path(source), options)
        blueprint = self._compile_blueprint(model, options)
        self._apply_metadata_policy(model, options)
        kwargs = {"font_probe": self.font_probe}
        if self.word_probe is not None:
            kwargs["word_probe"] = self.word_probe
        preflight = analyze_preflight(model, blueprint, options.interactive, **kwargs)
        blueprint_payload = {
            "schema_version": blueprint.schema_version,
            "source_sha256": blueprint.source_sha256,
            "source_model_fingerprint": blueprint.source_model_fingerprint,
            "fingerprint": blueprint.fingerprint,
            "total_events": blueprint.total_events,
            "total_visible_characters": blueprint.total_visible_characters,
            "semantic_counts": blueprint.semantic_counts,
            "events": [asdict(event) for event in blueprint.events],
        }
        (paths.working_dir / "blueprint.json").write_text(
            json.dumps(blueprint_payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8"
        )
        (paths.logs_dir / "preflight.json").write_text(
            json.dumps(asdict(preflight), indent=2, ensure_ascii=False, default=str), encoding="utf-8"
        )
        audit.append("BLUEPRINT_COMPILED", {"fingerprint": blueprint.fingerprint, "events": blueprint.total_events})
        audit.append("PREFLIGHT_COMPLETED", {"can_proceed": preflight.can_proceed, "blocking_reasons": preflight.blocking_reasons})
        callback = getattr(observer, "preflight", None) if observer is not None else None
        if callback is not None:
            callback(preflight)
        return PreparedInteractiveRun(str(Path(source).resolve()), model, blueprint, preflight, paths, options)

    @staticmethod
    def _settings_dict(options: InteractiveOptions) -> dict:
        data = asdict(options)
        data["speed_mode"] = options.speed_mode.value
        data["fidelity"] = options.fidelity.value
        return data

    @classmethod
    def _run_settings_dict(cls, options: RebuildOptions) -> dict:
        data = cls._settings_dict(options.interactive)
        data["metadata_mode"] = options.metadata.value
        data["visibility"] = options.visibility.value
        data["preserve_author_fields"] = bool(options.preserve_author_fields)
        data["custom_metadata_allowlist"] = list(options.custom_metadata_allowlist)
        return data

    @staticmethod
    def _options_from_settings(settings: dict) -> InteractiveOptions:
        allowed = set(InteractiveOptions.__dataclass_fields__)
        values = {key: value for key, value in dict(settings).items() if key in allowed}
        values["speed_mode"] = InteractiveSpeedMode(values.get("speed_mode", InteractiveSpeedMode.FAST.value))
        values["fidelity"] = InteractiveFidelity(values.get("fidelity", InteractiveFidelity.MAXIMUM.value))
        return InteractiveOptions(**values)

    @classmethod
    def _rebuild_options_from_settings(cls, settings: dict) -> RebuildOptions:
        interactive = cls._options_from_settings(settings)
        return RebuildOptions(
            reconstruction_mode=ReconstructionMode.INTERACTIVE,
            visibility=VisibilityMode(settings.get("visibility", VisibilityMode.BACKGROUND.value)),
            interactive=interactive,
            metadata=MetadataMode(settings.get("metadata_mode", MetadataMode.FRESH.value)),
            preserve_author_fields=bool(settings.get("preserve_author_fields", False)),
            custom_metadata_allowlist=tuple(settings.get("custom_metadata_allowlist", ()) or ()),
        )

    @staticmethod
    def _materialize_assets(model, working_dir: Path) -> dict[str, Path]:
        assets_dir = working_dir / "interactive_assets"
        assets_dir.mkdir(parents=True, exist_ok=True)
        result: dict[str, Path] = {}
        for asset_id, asset in model.assets.items():
            suffix = detect_image_extension(asset.part_name, asset) or Path(asset.part_name).suffix or ".bin"
            target = assets_dir / f"{asset_id}{suffix}"
            if not target.exists() or target.read_bytes() != asset.bytes_data:
                target.write_bytes(asset.bytes_data)
            result[asset_id] = target
        return result

    @staticmethod
    def _remove_unexpected_bookmarks(output_path: Path, expected_names: set[str]) -> int:
        """Remove Word-generated bookmarks that were not present in the source model."""
        from lxml import etree

        output_path = Path(output_path)
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        names = {str(name) for name in expected_names}
        with ZipFile(output_path, "r") as archive:
            parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        document_xml = parts.get("word/document.xml")
        if document_xml is None:
            return 0

        root = etree.fromstring(document_xml)
        ns = {"w": namespace}
        unexpected_ids: set[str] = set()
        removed = 0
        for node in root.xpath("//w:bookmarkStart", namespaces=ns):
            if node.get(f"{{{namespace}}}name") in names:
                continue
            bookmark_id = node.get(f"{{{namespace}}}id")
            if bookmark_id is not None:
                unexpected_ids.add(bookmark_id)
            node.getparent().remove(node)
            removed += 1
        for node in root.xpath("//w:bookmarkEnd", namespaces=ns):
            if node.get(f"{{{namespace}}}id") in unexpected_ids:
                node.getparent().remove(node)

        if not removed:
            return 0
        parts["word/document.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return removed

    @staticmethod
    def _remove_unexpected_headers(output_path: Path, expected_targets: set[str]) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rel_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        parts: dict[str, bytes] = {}
        with ZipFile(output_path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}

        def normalized_target(target: str) -> str:
            return target.lstrip("/") if target.startswith("/") else f"word/{target}"

        rels_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
        relationships = rels_root.xpath(
            "//*[local-name()='Relationship' and contains(@Type, '/header')]"
        )
        target_by_id = {node.get("Id"): normalized_target(node.get("Target", "")) for node in relationships}
        document_root = etree.fromstring(parts["word/document.xml"])
        ns = {"w": namespace, "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        removed_ids: set[str] = set()
        removed = 0
        for node in document_root.xpath("//w:headerReference", namespaces=ns):
            rel_id = node.get("{%s}id" % ns["r"])
            if target_by_id.get(rel_id) in expected_targets:
                continue
            if rel_id:
                removed_ids.add(rel_id)
            node.getparent().remove(node)
            removed += 1

        if not removed and all(target in expected_targets for target in target_by_id.values()):
            return 0

        for node in relationships:
            rel_id = node.get("Id")
            target = target_by_id.get(rel_id)
            if rel_id in removed_ids or target not in expected_targets:
                rels_root.remove(node)
                parts.pop(target, None)

        if "[Content_Types].xml" in parts:
            types_root = etree.fromstring(parts["[Content_Types].xml"])
            for node in types_root.xpath(
                "//*[local-name()='Override' and contains(@ContentType, '.header+xml')]"
            ):
                if normalized_target(node.get("PartName", "")) not in expected_targets:
                    types_root.remove(node)
            parts["[Content_Types].xml"] = etree.tostring(
                types_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        parts["word/document.xml"] = etree.tostring(
            document_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        parts["word/_rels/document.xml.rels"] = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return removed

    @staticmethod
    def _remove_template_paragraph_spacing(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        removed = 0
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            if source_paragraph.find("w:pPr", namespaces=ns) is not None:
                continue
            output_ppr = output_paragraph.find("w:pPr", namespaces=ns)
            if output_ppr is None:
                continue
            children = [etree.QName(child).localname for child in output_ppr]
            if children and all(child == "spacing" for child in children):
                output_paragraph.remove(output_ppr)
                removed += 1
        if not removed:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return removed

    @staticmethod
    def _remove_unexpected_header_shape_defaults(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            if "word/settings.xml" in source_archive.namelist():
                source_root = etree.fromstring(source_archive.read("word/settings.xml"))
                if source_root.xpath("./w:hdrShapeDefaults", namespaces=ns):
                    return 0
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        if "word/settings.xml" not in parts:
            return 0
        output_root = etree.fromstring(parts["word/settings.xml"])
        unexpected = output_root.xpath("./w:hdrShapeDefaults", namespaces=ns)
        for node in unexpected:
            output_root.remove(node)
        if not unexpected:
            return 0
        parts["word/settings.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return len(unexpected)

    @staticmethod
    def _restore_explicit_paragraph_alignment(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        restored = 0
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            source_jc = source_paragraph.find("w:pPr/w:jc", namespaces=ns)
            if source_jc is None:
                continue
            output_ppr = output_paragraph.find("w:pPr", namespaces=ns)
            if output_ppr is None:
                output_ppr = etree.Element(f"{{{namespace}}}pPr")
                output_paragraph.insert(0, output_ppr)
            output_jc = output_ppr.find("w:jc", namespaces=ns)
            if output_jc is not None and output_jc.get(f"{{{namespace}}}val") == source_jc.get(f"{{{namespace}}}val"):
                continue
            copied = etree.fromstring(etree.tostring(source_jc))
            if output_jc is not None:
                output_ppr.replace(output_jc, copied)
            else:
                spacing = output_ppr.find("w:spacing", namespaces=ns)
                if spacing is None:
                    output_ppr.append(copied)
                else:
                    output_ppr.insert(list(output_ppr).index(spacing), copied)
            restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_explicit_run_character_spacing(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_runs = source_root.xpath("//w:body//w:r", namespaces=ns)
        output_runs = output_root.xpath("//w:body//w:r", namespaces=ns)
        restored = 0
        for source_run, output_run in zip(source_runs, output_runs):
            source_spacing = source_run.find("w:rPr/w:spacing", namespaces=ns)
            if source_spacing is None:
                continue
            output_rpr = output_run.find("w:rPr", namespaces=ns)
            if output_rpr is None:
                output_rpr = etree.Element(f"{{{namespace}}}rPr")
                output_run.insert(0, output_rpr)
            output_spacing = output_rpr.find("w:spacing", namespaces=ns)
            copied = etree.fromstring(etree.tostring(source_spacing))
            if output_spacing is None:
                output_rpr.append(copied)
            else:
                output_rpr.replace(output_spacing, copied)
            if output_spacing is None or output_spacing.get(f"{{{namespace}}}val") != source_spacing.get(f"{{{namespace}}}val"):
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_explicit_run_font_names(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        if len(source_paragraphs) != len(output_paragraphs):
            return 0
        source_paragraph_texts = [
            "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
            for paragraph in source_paragraphs
        ]
        output_paragraph_texts = [
            "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
            for paragraph in output_paragraphs
        ]
        if source_paragraph_texts != output_paragraph_texts:
            return 0
        restored = 0
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            source_runs = source_paragraph.findall("w:r", namespaces=ns)
            output_runs = output_paragraph.findall("w:r", namespaces=ns)
            source_texts = ["".join(run.xpath(".//w:t/text()", namespaces=ns)) for run in source_runs]
            output_texts = ["".join(run.xpath(".//w:t/text()", namespaces=ns)) for run in output_runs]
            font_pairs = []
            if source_texts == output_texts:
                font_pairs = [
                    (source_run.find("w:rPr/w:rFonts", namespaces=ns), output_run)
                    for source_run, output_run, source_text in zip(
                        source_runs, output_runs, source_texts
                    )
                    if source_text
                ]
            elif "".join(source_texts) == "".join(output_texts):
                source_text_runs = [
                    run for run, text in zip(source_runs, source_texts) if text
                ]
                output_text_runs = [
                    run for run, text in zip(output_runs, output_texts) if text
                ]
                source_fonts = [
                    run.find("w:rPr/w:rFonts", namespaces=ns)
                    for run in source_text_runs
                ]
                if (
                    source_fonts
                    and output_text_runs
                    and all(fonts is not None for fonts in source_fonts)
                    and all(
                        dict(fonts.attrib) == dict(source_fonts[0].attrib)
                        for fonts in source_fonts[1:]
                    )
                ):
                    font_pairs = [
                        (source_fonts[0], output_run)
                        for output_run in output_text_runs
                    ]
            for source_fonts, output_run in font_pairs:
                if source_fonts is None:
                    continue
                output_rpr = output_run.find("w:rPr", namespaces=ns)
                output_fonts = (
                    output_rpr.find("w:rFonts", namespaces=ns) if output_rpr is not None else None
                )
                if output_fonts is not None and dict(output_fonts.attrib) == dict(source_fonts.attrib):
                    continue
                copied = etree.fromstring(etree.tostring(source_fonts))
                if output_rpr is None:
                    output_rpr = etree.Element(f"{{{namespace}}}rPr")
                    output_run.insert(0, output_rpr)
                if output_fonts is not None:
                    output_rpr.replace(output_fonts, copied)
                else:
                    output_style = output_rpr.find("w:rStyle", namespaces=ns)
                    insertion_index = (
                        list(output_rpr).index(output_style) + 1
                        if output_style is not None
                        else 0
                    )
                    output_rpr.insert(insertion_index, copied)
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_explicit_drawing_effect_extents(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        drawing_namespace = (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        )
        drawingml_namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
        relationship_namespace = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        )
        ns = {
            "w": word_namespace,
            "wp": drawing_namespace,
            "a": drawingml_namespace,
            "r": relationship_namespace,
        }
        with ZipFile(source_path) as source_archive:
            source_parts = {
                name: source_archive.read(name) for name in source_archive.namelist()
            }
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        source_root = etree.fromstring(source_parts["word/document.xml"])
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        if len(source_paragraphs) != len(output_paragraphs):
            return 0

        def visible_paragraph_text(paragraph) -> str:
            return "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))

        source_texts = [visible_paragraph_text(paragraph) for paragraph in source_paragraphs]
        output_texts = [visible_paragraph_text(paragraph) for paragraph in output_paragraphs]
        if source_texts != output_texts:
            return 0
        source_drawings = source_root.xpath("//w:body//w:drawing", namespaces=ns)
        output_drawings = output_root.xpath("//w:body//w:drawing", namespaces=ns)
        if len(source_drawings) != len(output_drawings):
            return 0

        def normalized_part_name(target: str) -> str:
            target_path = PurePosixPath(target.lstrip("/"))
            combined = target_path if target.startswith("/") else PurePosixPath("word") / target_path
            normalized = []
            for part in combined.parts:
                if part in ("", "."):
                    continue
                if part == "..":
                    if not normalized:
                        return ""
                    normalized.pop()
                else:
                    normalized.append(part)
            return "/".join(normalized)

        def relationship_targets(package_parts: dict[str, bytes]) -> dict[str, str]:
            relationships_name = "word/_rels/document.xml.rels"
            if relationships_name not in package_parts:
                return {}
            relationships_root = etree.fromstring(package_parts[relationships_name])
            return {
                relationship.get("Id"): relationship.get("Target")
                for relationship in relationships_root
                if relationship.get("Id")
                and relationship.get("Target")
                and relationship.get("TargetMode") != "External"
            }

        def drawing_identities(root, package_parts: dict[str, bytes]):
            paragraphs = root.xpath("//w:body//w:p", namespaces=ns)
            paragraph_indices = {id(paragraph): index for index, paragraph in enumerate(paragraphs)}
            targets = relationship_targets(package_parts)
            identities = []
            for drawing in root.xpath("//w:body//w:drawing", namespaces=ns):
                paragraph = drawing.xpath("ancestor::w:p[1]", namespaces=ns)
                if not paragraph:
                    return None
                paragraph = paragraph[0]
                paragraph_drawings = paragraph.xpath(".//w:drawing", namespaces=ns)
                drawing_index = next(
                    (index for index, candidate in enumerate(paragraph_drawings) if candidate is drawing),
                    None,
                )
                if drawing_index is None:
                    return None
                host = drawing.find("wp:inline", namespaces=ns)
                if host is None:
                    host = drawing.find("wp:anchor", namespaces=ns)
                extent = host.find("wp:extent", namespaces=ns) if host is not None else None
                if host is None or extent is None:
                    return None
                media_hashes = []
                for blip in drawing.xpath(".//a:blip[@r:embed]", namespaces=ns):
                    relationship_id = blip.get(f"{{{relationship_namespace}}}embed")
                    target = targets.get(relationship_id)
                    part_name = normalized_part_name(target) if target else ""
                    if not part_name or part_name not in package_parts:
                        return None
                    media_hashes.append(hashlib.sha256(package_parts[part_name]).hexdigest())
                identities.append(
                    (
                        paragraph_indices[id(paragraph)],
                        drawing_index,
                        visible_paragraph_text(paragraph),
                        etree.QName(host).localname,
                        tuple(media_hashes),
                    )
                )
            return identities

        source_identities = drawing_identities(source_root, source_parts)
        output_identities = drawing_identities(output_root, parts)
        if source_identities is None or source_identities != output_identities:
            return 0
        restored = 0
        for source_drawing, output_drawing in zip(source_drawings, output_drawings):
            source_host = source_drawing.find("wp:inline", namespaces=ns)
            if source_host is None:
                source_host = source_drawing.find("wp:anchor", namespaces=ns)
            output_host = output_drawing.find("wp:inline", namespaces=ns)
            if output_host is None:
                output_host = output_drawing.find("wp:anchor", namespaces=ns)
            if (
                source_host is None
                or output_host is None
                or source_host.tag != output_host.tag
            ):
                continue
            source_extent = source_host.find("wp:extent", namespaces=ns)
            output_extent = output_host.find("wp:extent", namespaces=ns)
            if (
                source_extent is None
                or output_extent is None
                or dict(source_extent.attrib) != dict(output_extent.attrib)
            ):
                continue
            source_effect = source_host.find("wp:effectExtent", namespaces=ns)
            if source_effect is None:
                continue
            output_effect = output_host.find("wp:effectExtent", namespaces=ns)
            if output_effect is not None and dict(output_effect.attrib) == dict(source_effect.attrib):
                continue
            copied = etree.fromstring(etree.tostring(source_effect))
            if output_effect is not None:
                output_host.replace(output_effect, copied)
            else:
                output_host.insert(list(output_host).index(output_extent) + 1, copied)
            restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_run_segmentation(output_path: Path, source_path: Path) -> int:
        """Restore source run boundaries in text-only body paragraphs.

        Word can coalesce adjacent, equally formatted runs when it saves the
        interactively typed document. In justified paragraphs those otherwise
        semantic-neutral boundaries measurably alter glyph placement in Word's
        PDF renderer. Keep the fast grouped COM insertion, then restore only
        paragraphs whose direct content is plain runs and whose visible text is
        already identical. Their source paragraph shell and run segmentation are
        deterministic package fidelity; all non-text structures stay owned by the
        reconstructed output.
        """
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        paragraph_properties_tag = f"{{{namespace}}}pPr"
        run_tag = f"{{{namespace}}}r"
        run_properties_tag = f"{{{namespace}}}rPr"
        text_tag = f"{{{namespace}}}t"
        rendered_break_tag = f"{{{namespace}}}lastRenderedPageBreak"

        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])

        def text_only_runs(paragraph):
            if any(child.tag not in {paragraph_properties_tag, run_tag} for child in paragraph):
                return None
            runs = paragraph.findall("w:r", namespaces=ns)
            for run in runs:
                if any(
                    child.tag not in {run_properties_tag, text_tag, rendered_break_tag}
                    for child in run
                ):
                    return None
            return runs

        restored = 0
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            source_runs = text_only_runs(source_paragraph)
            output_runs = text_only_runs(output_paragraph)
            if source_runs is None or output_runs is None:
                continue
            source_text = "".join(source_paragraph.xpath(".//w:t/text()", namespaces=ns))
            output_text = "".join(output_paragraph.xpath(".//w:t/text()", namespaces=ns))
            if source_text != output_text:
                continue
            source_properties = source_paragraph.find("w:pPr", namespaces=ns)
            output_properties = output_paragraph.find("w:pPr", namespaces=ns)
            if (
                source_properties is not None
                and source_properties.find(".//w:sectPr", namespaces=ns) is not None
            ):
                continue
            source_run_xml = [etree.tostring(run) for run in source_runs]
            output_run_xml = [etree.tostring(run) for run in output_runs]
            source_properties_xml = (
                etree.tostring(source_properties) if source_properties is not None else None
            )
            output_properties_xml = (
                etree.tostring(output_properties) if output_properties is not None else None
            )
            if (
                source_run_xml == output_run_xml
                and source_properties_xml == output_properties_xml
                and dict(source_paragraph.attrib) == dict(output_paragraph.attrib)
            ):
                continue
            if output_properties is not None:
                output_paragraph.remove(output_properties)
            if source_properties is not None:
                output_paragraph.insert(0, etree.fromstring(source_properties_xml))
            for run in output_runs:
                output_paragraph.remove(run)
            for run in source_runs:
                output_paragraph.append(etree.fromstring(etree.tostring(run)))
            output_paragraph.attrib.clear()
            output_paragraph.attrib.update(source_paragraph.attrib)
            restored += 1

        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_empty_runs(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        restored = 0
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            source_empty = [run for run in source_paragraph.findall("w:r", namespaces=ns) if len(run) == 0]
            output_empty = [run for run in output_paragraph.findall("w:r", namespaces=ns) if len(run) == 0]
            for _ in range(max(0, len(source_empty) - len(output_empty))):
                output_paragraph.append(etree.Element(f"{{{namespace}}}r"))
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_explicit_column_space(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        attr = f"{{{namespace}}}space"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        restored = 0
        for source_section, output_section in zip(
            source_root.xpath("//w:sectPr", namespaces=ns), output_root.xpath("//w:sectPr", namespaces=ns)
        ):
            source_cols = source_section.find("w:cols", namespaces=ns)
            output_cols = output_section.find("w:cols", namespaces=ns)
            if source_cols is None or source_cols.get(attr) is None or output_cols is None:
                continue
            if output_cols.get(attr) != source_cols.get(attr):
                output_cols.set(attr, source_cols.get(attr))
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_explicit_page_number_start(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        attr = f"{{{namespace}}}start"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        restored = 0
        for source_section, output_section in zip(
            source_root.xpath("//w:sectPr", namespaces=ns), output_root.xpath("//w:sectPr", namespaces=ns)
        ):
            source_num = source_section.find("w:pgNumType", namespaces=ns)
            if source_num is None or source_num.get(attr) is None:
                continue
            output_num = output_section.find("w:pgNumType", namespaces=ns)
            if output_num is None:
                output_num = etree.Element(f"{{{namespace}}}pgNumType")
                output_section.insert(0, output_num)
            if output_num.get(attr) != source_num.get(attr):
                output_num.set(attr, source_num.get(attr))
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_document_defaults(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        with ZipFile(source_path) as source_archive:
            source_styles = source_archive.read("word/styles.xml")
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        if "word/styles.xml" not in parts:
            return 0
        source_root = etree.fromstring(source_styles)
        output_root = etree.fromstring(parts["word/styles.xml"])
        source_defaults = next((node for node in source_root if etree.QName(node).localname == "docDefaults"), None)
        output_defaults = next((node for node in output_root if etree.QName(node).localname == "docDefaults"), None)
        if source_defaults is None or output_defaults is None:
            return 0
        if etree.tostring(source_defaults) == etree.tostring(output_defaults):
            return 0
        index = list(output_root).index(output_defaults)
        output_root.remove(output_defaults)
        output_root.insert(index, etree.fromstring(etree.tostring(source_defaults)))
        parts["word/styles.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return 1

    @staticmethod
    def _restore_source_theme_style_latin_fonts(
        output_path: Path, source_path: Path
    ) -> int:
        from lxml import etree

        word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        drawing_namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"

        def word_attr(name: str) -> str:
            return f"{{{word_namespace}}}{name}"

        namespaces = {"w": word_namespace, "a": drawing_namespace}
        styles_path = "word/styles.xml"
        relationships_path = "word/_rels/document.xml.rels"
        relationship_namespace = (
            "http://schemas.openxmlformats.org/package/2006/relationships"
        )

        def active_theme_path(part_names: set[str], relationships: bytes) -> str | None:
            root = etree.fromstring(relationships)
            for relationship in root.findall(
                f"{{{relationship_namespace}}}Relationship"
            ):
                relationship_type = relationship.get("Type") or ""
                if not relationship_type.endswith("/theme"):
                    continue
                if relationship.get("TargetMode") == "External":
                    continue
                target = relationship.get("Target") or ""
                if not target:
                    return None
                if target.startswith("/"):
                    candidate = posixpath.normpath(target.lstrip("/"))
                else:
                    candidate = posixpath.normpath(posixpath.join("word", target))
                if candidate == ".." or candidate.startswith("../"):
                    return None
                return candidate if candidate in part_names else None
            return None

        with ZipFile(source_path) as source_archive:
            source_names = set(source_archive.namelist())
            if relationships_path not in source_names or styles_path not in source_names:
                return 0
            source_theme_path = active_theme_path(
                source_names, source_archive.read(relationships_path)
            )
            if source_theme_path is None:
                return 0
            source_theme = etree.fromstring(source_archive.read(source_theme_path))
            source_styles = etree.fromstring(source_archive.read(styles_path))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        if styles_path not in parts:
            return 0

        output_styles = etree.fromstring(parts[styles_path])
        source_latin_scheme = {}
        for prefix, font_group in (("major", "majorFont"), ("minor", "minorFont")):
            source_latin = source_theme.find(
                f".//a:{font_group}/a:latin", namespaces
            )
            typeface = source_latin.get("typeface") if source_latin is not None else None
            if typeface:
                source_latin_scheme[f"{prefix}Ascii"] = typeface
                source_latin_scheme[f"{prefix}HAnsi"] = typeface

        def style_key(style):
            return style.get(word_attr("type")), style.get(word_attr("styleId"))

        output_style_map = {
            style_key(style): style for style in output_styles.findall("w:style", namespaces)
        }
        restored_style_fonts = 0
        for source_style in source_styles.findall("w:style", namespaces):
            output_style = output_style_map.get(style_key(source_style))
            if output_style is None:
                continue
            source_fonts = source_style.find("w:rPr/w:rFonts", namespaces)
            if source_fonts is None:
                continue
            output_run_properties = output_style.find("w:rPr", namespaces)
            if output_run_properties is None:
                continue
            output_fonts = output_run_properties.find("w:rFonts", namespaces)
            if output_fonts is None:
                continue
            changed = False
            for direct_name, theme_name in (
                ("ascii", "asciiTheme"),
                ("hAnsi", "hAnsiTheme"),
            ):
                theme_reference = source_fonts.get(word_attr(theme_name))
                resolved_font = source_latin_scheme.get(theme_reference or "")
                if not resolved_font:
                    continue
                direct_attribute = word_attr(direct_name)
                theme_attribute = word_attr(theme_name)
                if (
                    output_fonts.get(direct_attribute) == resolved_font
                    and output_fonts.get(theme_attribute) is None
                ):
                    continue
                output_fonts.set(direct_attribute, resolved_font)
                output_fonts.attrib.pop(theme_attribute, None)
                changed = True
            if changed:
                restored_style_fonts += 1

        if not restored_style_fonts:
            return 0
        if restored_style_fonts:
            parts[styles_path] = etree.tostring(
                output_styles, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored_style_fonts

    @staticmethod
    def _restore_relationship_free_source_headers(
        output_path: Path,
        source_path: Path,
        expected_targets: set[str],
    ) -> int:
        with ZipFile(source_path) as source_archive:
            source_names = set(source_archive.namelist())
            source_parts = {
                name: source_archive.read(name)
                for name in expected_targets
                if name in source_names and name.startswith("word/header") and name.endswith(".xml")
            }
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        restored = 0
        for name, data in source_parts.items():
            part = PurePosixPath(name)
            relationship_part = str(part.parent / "_rels" / f"{part.name}.rels")
            if relationship_part in source_names or relationship_part in parts:
                continue
            if name in parts and parts[name] != data:
                parts[name] = data
                restored += 1
        if not restored:
            return 0
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_footer_stories(output_path: Path, source_path: Path, expected_types: list[set[str]]) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rel_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        ns = {"w": namespace, "r": rel_namespace}

        def footer_targets_by_section(archive_path: Path) -> tuple[dict[int, dict[str, str]], dict[str, bytes]]:
            with ZipFile(archive_path) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            document_root = etree.fromstring(parts["word/document.xml"])
            rels_bytes = parts.get("word/_rels/document.xml.rels")
            rel_map = {}
            if rels_bytes is not None:
                rels_root = etree.fromstring(rels_bytes)
                rel_map = {rel.get("Id"): rel.get("Target") for rel in rels_root}
            by_section: dict[int, dict[str, str]] = {}
            for index, section in enumerate(document_root.xpath("//w:sectPr", namespaces=ns)):
                types: dict[str, str] = {}
                for ref in section.xpath("./w:footerReference", namespaces=ns):
                    footer_type = ref.get(f"{{{namespace}}}type") or "default"
                    rid = ref.get(f"{{{rel_namespace}}}id")
                    target = rel_map.get(rid)
                    if target is not None:
                        types[footer_type] = f"word/{target}"
                by_section[index] = types
            return by_section, parts

        # Physical footer part names (footer1.xml, footer2.xml, ...) are assigned
        # independently by whichever tool last saved each package - Word routinely
        # renumbers them on rebuild (e.g. it materializes empty even/first-page
        # variants the source never declared). Matching parts by filename would
        # therefore silently pair up unrelated footers from different sections that
        # happen to share a number. Match by (section index, footer type) instead,
        # resolved through each package's own relationships.
        source_sections, source_parts = footer_targets_by_section(source_path)
        output_sections, output_parts = footer_targets_by_section(output_path)
        document_root = etree.fromstring(output_parts["word/document.xml"])
        rels_root = etree.fromstring(output_parts["word/_rels/document.xml.rels"])
        types_root = (
            etree.fromstring(output_parts["[Content_Types].xml"])
            if "[Content_Types].xml" in output_parts else None
        )

        restored = 0
        sections = document_root.xpath("//w:sectPr", namespaces=ns)
        missing: list[tuple[Any, str, str]] = []  # (sectPr element, footer_type, source_target)
        for index, section in enumerate(sections):
            allowed = expected_types[index] if index < len(expected_types) else set()
            source_types = source_sections.get(index, {})
            output_types = output_sections.get(index, {})
            for ref in section.xpath("./w:footerReference", namespaces=ns):
                footer_type = ref.get(f"{{{namespace}}}type") or "default"
                if footer_type not in allowed:
                    ref.getparent().remove(ref)
                    restored += 1
                    continue
                source_target = source_types.get(footer_type)
                output_target = output_types.get(footer_type)
                if source_target is None or output_target is None:
                    continue
                source_data = source_parts.get(source_target)
                if source_data is not None and output_parts.get(output_target) != source_data:
                    output_parts[output_target] = source_data
                    restored += 1
            # A footer type the source declares for this section but that has no
            # footerReference left in the output at all (e.g. Word's own automatic
            # repagination silently dropped it while inserting an absolutely
            # positioned image elsewhere in the document - confirmed live) can't be
            # repaired above, since there is no existing reference to redirect.
            # It must be rebuilt: a new relationship, a new physical part, and a
            # new <w:footerReference> inserted back into this section.
            present_types = {
                ref.get(f"{{{namespace}}}type") or "default"
                for ref in section.xpath("./w:footerReference", namespaces=ns)
            }
            for footer_type in sorted(allowed - present_types):
                source_target = source_types.get(footer_type)
                if source_target is None or source_target not in source_parts:
                    continue
                missing.append((section, footer_type, source_target))

        if missing:
            footer_rel_type = f"{rel_namespace}/footer"
            footer_content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
            existing_rel_type = rels_root.xpath("//*[local-name()='Relationship' and contains(@Type, '/footer')]/@Type")
            if existing_rel_type:
                footer_rel_type = str(existing_rel_type[0])
            if types_root is not None:
                existing_content_type = types_root.xpath(
                    "//*[local-name()='Override' and contains(@ContentType, '.footer+xml')]/@ContentType"
                )
                if existing_content_type:
                    footer_content_type = str(existing_content_type[0])

            def next_numbered_name(prefix: str, existing: set[str]) -> str:
                n = 1
                while f"{prefix}{n}.xml" in existing:
                    n += 1
                return f"{prefix}{n}.xml"

            existing_rids = {rel.get("Id") for rel in rels_root}
            existing_footer_files = {name.split("/", 1)[-1] for name in output_parts if name.startswith("word/footer")}

            for section, footer_type, source_target in missing:
                target_name = next_numbered_name("footer", existing_footer_files)
                existing_footer_files.add(target_name)
                new_rid = f"rId{len(existing_rids) + 1}"
                while new_rid in existing_rids:
                    new_rid = f"rId{int(new_rid[3:]) + 1}"
                existing_rids.add(new_rid)

                relationship = etree.SubElement(
                    rels_root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
                )
                relationship.set("Id", new_rid)
                relationship.set("Type", footer_rel_type)
                relationship.set("Target", target_name)

                if types_root is not None:
                    override = etree.SubElement(
                        types_root, "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
                    )
                    override.set("PartName", f"/word/{target_name}")
                    override.set("ContentType", footer_content_type)

                new_ref = etree.Element(f"{{{namespace}}}footerReference")
                new_ref.set(f"{{{namespace}}}type", footer_type)
                new_ref.set(f"{{{rel_namespace}}}id", new_rid)
                # footerReference elements must sort before pgSz/pgMar/etc per the
                # OOXML sectPr content model - insert alongside any that remain.
                existing_refs = section.xpath("./w:footerReference", namespaces=ns)
                if existing_refs:
                    existing_refs[-1].addnext(new_ref)
                else:
                    section.insert(0, new_ref)

                output_parts[f"word/{target_name}"] = source_parts[source_target]
                restored += 1

        if not restored:
            return 0
        output_parts["word/document.xml"] = etree.tostring(
            document_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        output_parts["word/_rels/document.xml.rels"] = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        if types_root is not None:
            output_parts["[Content_Types].xml"] = etree.tostring(
                types_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in output_parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_footer_topology(output_path: Path, source_path: Path) -> int:
        """Restore the source footer graph after Word materializes duplicates.

        Word may split one footer shared by several sections into separate
        physical parts and leave additional generated footer parts behind.
        Restore only the footer subgraph from the signed source package: all
        footer parts, footer relationships, content-type declarations, and the
        direct footerReference children of each matching section. If section
        topology or any referenced source part is incomplete, fail closed.
        """
        from lxml import etree

        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        pr_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        document_name = "word/document.xml"
        rels_name = "word/_rels/document.xml.rels"
        types_name = "[Content_Types].xml"
        footer_content_type_suffix = ".wordprocessingml.footer+xml"

        with ZipFile(source_path) as source_archive:
            source_parts = {
                name: source_archive.read(name) for name in source_archive.namelist()
            }
        with ZipFile(output_path) as output_archive:
            output_parts = {
                name: output_archive.read(name) for name in output_archive.namelist()
            }
        required = {document_name, rels_name, types_name}
        if not required <= source_parts.keys() or not required <= output_parts.keys():
            return 0

        source_document = etree.fromstring(source_parts[document_name])
        output_document = etree.fromstring(output_parts[document_name])
        source_rels = etree.fromstring(source_parts[rels_name])
        output_rels = etree.fromstring(output_parts[rels_name])
        source_types = etree.fromstring(source_parts[types_name])
        output_types = etree.fromstring(output_parts[types_name])
        source_sections = source_document.xpath("//w:sectPr", namespaces={"w": w_ns})
        output_sections = output_document.xpath("//w:sectPr", namespaces={"w": w_ns})
        if len(source_sections) != len(output_sections):
            return 0

        def footer_relationships(root):
            return [
                node
                for node in root.findall(f"{{{pr_ns}}}Relationship")
                if (node.get("Type") or "").endswith("/footer")
                and node.get("TargetMode") != "External"
            ]

        def footer_part_names(types_root):
            return {
                (node.get("PartName") or "").lstrip("/")
                for node in types_root.findall(f"{{{ct_ns}}}Override")
                if (node.get("ContentType") or "").endswith(footer_content_type_suffix)
            }

        source_footer_rels = footer_relationships(source_rels)
        output_footer_rels = footer_relationships(output_rels)
        source_footer_parts = footer_part_names(source_types)
        output_footer_parts = footer_part_names(output_types)
        if any(name not in source_parts for name in source_footer_parts):
            return 0

        source_rids = {node.get("Id") for node in source_footer_rels}
        if any(
            ref.get(f"{{{r_ns}}}id") not in source_rids
            for section in source_sections
            for ref in section.findall(f"{{{w_ns}}}footerReference")
        ):
            return 0
        if not source_footer_rels and not output_footer_rels and not source_footer_parts and not output_footer_parts:
            return 0

        for node in output_footer_rels:
            output_rels.remove(node)
        for name in output_footer_parts:
            output_parts.pop(name, None)
            path = PurePosixPath(name)
            output_parts.pop(str(path.parent / "_rels" / f"{path.name}.rels"), None)
        for node in list(output_types.findall(f"{{{ct_ns}}}Override")):
            if (node.get("PartName") or "").lstrip("/") in output_footer_parts:
                output_types.remove(node)

        used_ids = {
            node.get("Id") for node in output_rels.findall(f"{{{pr_ns}}}Relationship")
        }
        rid_map: dict[str, str] = {}
        next_id = 1
        for source_rel in source_footer_rels:
            source_id = source_rel.get("Id") or ""
            new_id = source_id
            if not new_id or new_id in used_ids:
                while f"rId{next_id}" in used_ids:
                    next_id += 1
                new_id = f"rId{next_id}"
                next_id += 1
            used_ids.add(new_id)
            rid_map[source_id] = new_id
            restored_rel = deepcopy(source_rel)
            restored_rel.set("Id", new_id)
            output_rels.append(restored_rel)

        source_overrides = {
            (node.get("PartName") or "").lstrip("/"): node
            for node in source_types.findall(f"{{{ct_ns}}}Override")
        }
        for name in source_footer_parts:
            output_parts[name] = source_parts[name]
            path = PurePosixPath(name)
            sidecar = str(path.parent / "_rels" / f"{path.name}.rels")
            if sidecar in source_parts:
                output_parts[sidecar] = source_parts[sidecar]
            declaration = source_overrides.get(name)
            if declaration is not None:
                output_types.append(deepcopy(declaration))

        for source_section, output_section in zip(source_sections, output_sections):
            for output_ref in list(output_section.findall(f"{{{w_ns}}}footerReference")):
                output_section.remove(output_ref)
            insertion_index = 0
            for child in output_section:
                if etree.QName(child).localname != "headerReference":
                    break
                insertion_index += 1
            for source_ref in source_section.findall(f"{{{w_ns}}}footerReference"):
                restored_ref = deepcopy(source_ref)
                restored_ref.set(
                    f"{{{r_ns}}}id", rid_map[source_ref.get(f"{{{r_ns}}}id")]
                )
                output_section.insert(insertion_index, restored_ref)
                insertion_index += 1

        output_parts[document_name] = etree.tostring(
            output_document, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        output_parts[rels_name] = etree.tostring(
            output_rels, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        output_parts[types_name] = etree.tostring(
            output_types, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{Path(output_path).name}.", suffix=".tmp", dir=Path(output_path).parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in output_parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), Path(output_path))
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return len(output_footer_rels) + len(source_footer_rels)

    @staticmethod
    def _restore_source_cross_paragraph_field_shells(output_path: Path, source_path: Path) -> int:
        """Restore a missing outer TOC field without replacing its result.

        WordReplica reconstructs every visible TOC paragraph and its nested
        PAGEREF fields, but Word's Fields.Add API cannot create one field whose
        cached result spans several paragraphs.  The saved document therefore
        lacks only the outer begin/instruction/separate/end shell.  Copy that
        shell from the source only when paragraph topology and every visible
        result string in the covered range are identical.  Any less certain
        shape fails closed.
        """
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        field_type_attr = f"{{{namespace}}}fldCharType"
        paragraph_tag = f"{{{namespace}}}p"

        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])

        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        if len(source_paragraphs) != len(output_paragraphs):
            return 0
        source_paragraph_index = {paragraph: index for index, paragraph in enumerate(source_paragraphs)}

        def owner_paragraph(node):
            return next((ancestor for ancestor in node.iterancestors() if ancestor.tag == paragraph_tag), None)

        records: list[dict] = []
        stack: list[dict] = []
        malformed = False
        for node in source_root.xpath("//w:body//*[self::w:fldChar or self::w:instrText]", namespaces=ns):
            local = etree.QName(node).localname
            if local == "instrText":
                if stack:
                    stack[-1]["instruction_nodes"].append(node)
                continue
            field_type = node.get(field_type_attr)
            if field_type == "begin":
                stack.append({"begin": node, "instruction_nodes": [], "separate": None, "end": None})
            elif field_type == "separate":
                if not stack:
                    malformed = True
                    break
                stack[-1]["separate"] = node
            elif field_type == "end":
                if not stack:
                    malformed = True
                    break
                record = stack.pop()
                record["end"] = node
                records.append(record)
        if malformed or stack:
            return 0

        output_instructions = {
            " ".join((node.text or "").split()).casefold()
            for node in output_root.xpath("//w:body//w:instrText", namespaces=ns)
        }

        restored = 0
        for record in records:
            begin_paragraph = owner_paragraph(record["begin"])
            separate_paragraph = owner_paragraph(record["separate"]) if record["separate"] is not None else None
            end_paragraph = owner_paragraph(record["end"])
            if begin_paragraph is None or separate_paragraph is not begin_paragraph or end_paragraph is None:
                continue
            begin_index = source_paragraph_index[begin_paragraph]
            end_index = source_paragraph_index[end_paragraph]
            if begin_index >= end_index:
                continue
            instruction = "".join(node.text or "" for node in record["instruction_nodes"])
            normalized_instruction = " ".join(instruction.split()).casefold()
            if not (normalized_instruction == "toc" or normalized_instruction.startswith("toc ")):
                continue
            if normalized_instruction in output_instructions:
                continue

            begin_run = record["begin"].getparent()
            separate_run = record["separate"].getparent()
            end_run = record["end"].getparent()
            if (
                begin_run is None
                or separate_run is None
                or end_run is None
                or begin_run.getparent() is not begin_paragraph
                or separate_run.getparent() is not begin_paragraph
                or end_run.getparent() is not end_paragraph
            ):
                continue
            begin_children = list(begin_paragraph)
            start = begin_children.index(begin_run)
            separate = begin_children.index(separate_run)
            if start > separate:
                continue
            end_children = list(end_paragraph)
            end_position = end_children.index(end_run)
            if any(child.xpath(".//w:t", namespaces=ns) for child in begin_children[:start]):
                continue
            if any(child.xpath(".//w:t", namespaces=ns) for child in end_children[end_position + 1 :]):
                continue

            source_visible = [
                [node.text or "" for node in paragraph.xpath(".//w:t", namespaces=ns)]
                for paragraph in source_paragraphs[begin_index : end_index + 1]
            ]
            output_visible = [
                [node.text or "" for node in paragraph.xpath(".//w:t", namespaces=ns)]
                for paragraph in output_paragraphs[begin_index : end_index + 1]
            ]
            if source_visible != output_visible:
                continue

            output_begin = output_paragraphs[begin_index]
            output_end = output_paragraphs[end_index]
            insert_at = 1 if output_begin.find("w:pPr", namespaces=ns) is not None else 0
            for offset, shell_node in enumerate(begin_children[start : separate + 1]):
                output_begin.insert(insert_at + offset, deepcopy(shell_node))
            output_end.append(deepcopy(end_run))
            output_instructions.add(normalized_instruction)
            restored += 1

        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_field_instructions(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        field_type_attr = f"{{{namespace}}}fldCharType"
        instruction_attr = f"{{{namespace}}}instr"
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])

        def parse_fields(root):
            records: list[dict] = []
            context: list[int] = []
            loose_instructions: list[object] = []
            malformed = False

            def add_record(kind: str, node) -> int:
                index = len(records)
                records.append(
                    {
                        "kind": kind,
                        "node": node,
                        "parent": context[-1] if context else None,
                        "depth": len(context),
                        "instruction_nodes": [],
                        "separate": None,
                        "end": None,
                    }
                )
                return index

            def walk(node) -> None:
                nonlocal malformed
                local_name = etree.QName(node).localname
                if local_name == "fldSimple":
                    index = add_record("simple", node)
                    context.append(index)
                    for child in node:
                        walk(child)
                    context.pop()
                    return
                if local_name == "fldChar":
                    field_type = node.get(field_type_attr)
                    if field_type == "begin":
                        context.append(add_record("complex", node))
                    elif field_type == "separate":
                        if not context or records[context[-1]]["kind"] != "complex":
                            malformed = True
                        else:
                            records[context[-1]]["separate"] = node
                    elif field_type == "end":
                        if not context or records[context[-1]]["kind"] != "complex":
                            malformed = True
                        else:
                            records[context[-1]]["end"] = node
                            context.pop()
                elif local_name == "instrText":
                    if context and records[context[-1]]["kind"] == "complex":
                        records[context[-1]]["instruction_nodes"].append(node)
                    else:
                        loose_instructions.append(node)
                for child in node:
                    walk(child)

            walk(root)
            if context:
                malformed = True
            for record in records:
                if record["kind"] == "simple":
                    record["instruction"] = record["node"].get(instruction_attr) or ""
                else:
                    record["instruction"] = "".join(
                        node.text or "" for node in record["instruction_nodes"]
                    )
                    if record["separate"] is None or record["end"] is None:
                        malformed = True
            return records, loose_instructions, malformed

        def normalized_instruction(value: str) -> str:
            tokens = value.split()
            normalized = []
            index = 0
            while index < len(tokens):
                if (
                    tokens[index] == "\\*"
                    and index + 1 < len(tokens)
                    and tokens[index + 1].casefold() == "mergeformat"
                ):
                    index += 2
                    continue
                normalized.append(tokens[index].casefold())
                index += 1
            return " ".join(normalized)

        def complex_span(record):
            begin_run = record["node"].getparent()
            separate_run = record["separate"].getparent()
            end_run = record["end"].getparent()
            if begin_run is None or separate_run is None or end_run is None:
                return None
            container = begin_run.getparent()
            if container is None or separate_run.getparent() is not container or end_run.getparent() is not container:
                return None
            start = container.index(begin_run)
            separate = container.index(separate_run)
            end = container.index(end_run)
            if not start <= separate < end:
                return None
            return container, start, separate, end

        def output_result(record):
            if record["kind"] == "simple":
                return list(record["node"]), record["node"].tail
            span = complex_span(record)
            if span is None:
                return None
            container, _, separate, end = span
            return list(container)[separate + 1 : end], list(container)[end].tail

        def source_replacement(record, result_nodes, output_tail):
            if record["kind"] == "simple":
                wrapper = deepcopy(record["node"])
                for child in list(wrapper):
                    wrapper.remove(child)
                for child in result_nodes:
                    wrapper.append(deepcopy(child))
                wrapper.tail = output_tail
                return [wrapper]
            span = complex_span(record)
            if span is None:
                return None
            container, start, separate, end = span
            children = list(container)
            replacement = [deepcopy(node) for node in children[start : separate + 1]]
            replacement.extend(deepcopy(node) for node in result_nodes)
            replacement.append(deepcopy(children[end]))
            replacement[-1].tail = output_tail
            return replacement

        def replace_output_field(record, replacement) -> bool:
            if record["kind"] == "simple":
                node = record["node"]
                parent = node.getparent()
                if parent is None:
                    return False
                index = parent.index(node)
                parent.remove(node)
            else:
                span = complex_span(record)
                if span is None:
                    return False
                parent, index, _, end = span
                for child in list(parent)[index : end + 1]:
                    parent.remove(child)
            for offset, node in enumerate(replacement):
                parent.insert(index + offset, node)
            return True

        source_fields, source_loose, source_malformed = parse_fields(source_root)
        output_fields, output_loose, output_malformed = parse_fields(output_root)
        if source_malformed or output_malformed or len(source_fields) != len(output_fields):
            return 0
        if [field["parent"] for field in source_fields] != [field["parent"] for field in output_fields]:
            return 0
        if any(
            normalized_instruction(source_field["instruction"])
            != normalized_instruction(output_field["instruction"])
            for source_field, output_field in zip(source_fields, output_fields)
        ):
            return 0
        if len(source_loose) != len(output_loose):
            return 0

        source_order = {node: index for index, node in enumerate(source_root.iter())}
        output_order = {node: index for index, node in enumerate(output_root.iter())}

        def has_nested_instruction_field(fields, field_index: int, order: dict) -> bool:
            field = fields[field_index]
            separate = field.get("separate")
            if separate is None:
                return False
            separate_position = order[separate]
            return any(
                child["parent"] == field_index and order[child["node"]] < separate_position
                for child in fields
            )

        def restore_instruction_only(source_field, output_field) -> int:
            if source_field["kind"] == "simple" and output_field["kind"] == "simple":
                source_instruction = source_field["node"].get(instruction_attr) or ""
                if output_field["node"].get(instruction_attr) != source_instruction:
                    output_field["node"].set(instruction_attr, source_instruction)
                    return 1
                return 0
            if source_field["kind"] != "complex" or output_field["kind"] != "complex":
                return 0
            source_nodes = source_field["instruction_nodes"]
            output_nodes = output_field["instruction_nodes"]
            if len(source_nodes) != len(output_nodes):
                return 0
            changed = False
            for source_node, output_node in zip(source_nodes, output_nodes):
                if output_node.text != source_node.text:
                    output_node.text = source_node.text
                    changed = True
            return int(changed)

        restored = 0
        paired_fields = list(enumerate(zip(source_fields, output_fields)))
        for field_index, (source_field, output_field) in sorted(
            paired_fields,
            key=lambda item: (item[1][1]["depth"], item[0]),
            reverse=True,
        ):
            if (
                has_nested_instruction_field(source_fields, field_index, source_order)
                or has_nested_instruction_field(output_fields, field_index, output_order)
            ):
                restored += restore_instruction_only(source_field, output_field)
                continue
            result = output_result(output_field)
            if result is None:
                restored += restore_instruction_only(source_field, output_field)
                continue
            result_nodes, output_tail = result
            replacement = source_replacement(source_field, result_nodes, output_tail)
            if replacement is not None and replace_output_field(output_field, replacement):
                restored += 1

        for source_node, output_node in zip(source_loose, output_loose):
            if output_node.text != source_node.text:
                output_node.text = source_node.text
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_table_layout(output_path: Path, source_path: Path) -> int:
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_tables = source_root.xpath("//w:body//w:tbl", namespaces=ns)
        output_tables = output_root.xpath("//w:body//w:tbl", namespaces=ns)

        table_tag = f"{{{namespace}}}tbl"

        def direct_nested_tables(cell, owner_table) -> list:
            nested = []
            for candidate in cell.xpath(".//w:tbl", namespaces=ns):
                nearest_table = next(
                    (ancestor for ancestor in candidate.iterancestors() if ancestor.tag == table_tag),
                    None,
                )
                if nearest_table is owner_table:
                    nested.append(candidate)
            return nested

        def table_topology(table) -> tuple:
            return tuple(
                tuple(
                    tuple(table_topology(nested) for nested in direct_nested_tables(cell, table))
                    for cell in row.xpath("./w:tc", namespaces=ns)
                )
                for row in table.xpath("./w:tr", namespaces=ns)
            )

        source_topology = tuple(
            table_topology(table)
            for table in source_tables
            if not any(ancestor.tag == table_tag for ancestor in table.iterancestors())
        )
        output_topology = tuple(
            table_topology(table)
            for table in output_tables
            if not any(ancestor.tag == table_tag for ancestor in table.iterancestors())
        )
        if source_topology != output_topology:
            return 0

        def synchronize_child(source_parent, output_parent, path: str, index: int) -> int:
            source_child = source_parent.find(path, namespaces=ns)
            output_child = output_parent.find(path, namespaces=ns)
            if source_child is None and output_child is None:
                return 0
            if source_child is not None and output_child is not None:
                if etree.tostring(source_child) == etree.tostring(output_child):
                    return 0
                output_parent.replace(output_child, deepcopy(source_child))
                return 1
            if output_child is not None:
                output_parent.remove(output_child)
                return 1
            output_parent.insert(min(index, len(output_parent)), deepcopy(source_child))
            return 1

        restored = 0
        for source_table, output_table in zip(source_tables, output_tables):
            restored += synchronize_child(source_table, output_table, "w:tblPr", 0)
            grid_index = 1 if output_table.find("w:tblPr", namespaces=ns) is not None else 0
            restored += synchronize_child(source_table, output_table, "w:tblGrid", grid_index)
            source_rows = source_table.xpath("./w:tr", namespaces=ns)
            output_rows = output_table.xpath("./w:tr", namespaces=ns)
            for source_row, output_row in zip(source_rows, output_rows):
                row_properties_index = 1 if output_row.find("w:tblPrEx", namespaces=ns) is not None else 0
                restored += synchronize_child(source_row, output_row, "w:trPr", row_properties_index)
                source_cells = source_row.xpath("./w:tc", namespaces=ns)
                output_cells = output_row.xpath("./w:tc", namespaces=ns)
                for source_cell, output_cell in zip(source_cells, output_cells):
                    restored += synchronize_child(source_cell, output_cell, "w:tcPr", 0)
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_numbering_definitions(output_path: Path, source_path: Path) -> int:
        """The interactive renderer cannot ask Word to reuse a source numId - it
        can only call ListFormat.ApplyBulletDefault()/ApplyNumberDefault(), which
        makes Word mint its own brand-new list definition (a different numId,
        and Word's own default indent/format for that level, discarding whatever
        indent the source's numbering level actually specified, and resetting
        any explicit <w:ind> override the paragraph had of its own). Since
        every list paragraph the renderer touches originates from a source
        paragraph that already carries its own explicit numId (see
        blueprint.py's CreateListBinding guard), it is always safe to redirect
        each such paragraph's numId/ilvl/explicit indent back to its source
        value and then replace Word's generated word/numbering.xml outright
        with the source's - no output paragraph ends up referencing a numId
        that source doesn't define.
        """
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            if "word/numbering.xml" not in source_archive.namelist():
                return 0
            source_numbering = source_archive.read("word/numbering.xml")
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        if "word/document.xml" not in parts:
            return 0
        output_root = etree.fromstring(parts["word/document.xml"])
        restored = 0
        for source_p, output_p in zip(
            source_root.xpath("//w:p", namespaces=ns), output_root.xpath("//w:p", namespaces=ns)
        ):
            source_num_id = source_p.find("w:pPr/w:numPr/w:numId", namespaces=ns)
            if source_num_id is None or source_num_id.get(f"{{{namespace}}}val") is None:
                continue
            output_num_pr = output_p.find("w:pPr/w:numPr", namespaces=ns)
            if output_num_pr is None:
                continue
            output_num_id = output_num_pr.find("w:numId", namespaces=ns)
            if output_num_id is None:
                continue
            attr = f"{{{namespace}}}val"
            changed = False
            if output_num_id.get(attr) != source_num_id.get(attr):
                output_num_id.set(attr, source_num_id.get(attr))
                changed = True
            source_ilvl = source_p.find("w:pPr/w:numPr/w:ilvl", namespaces=ns)
            output_ilvl = output_num_pr.find("w:ilvl", namespaces=ns)
            source_ilvl_val = source_ilvl.get(attr) if source_ilvl is not None else None
            if output_ilvl is not None and source_ilvl_val is not None and output_ilvl.get(attr) != source_ilvl_val:
                output_ilvl.set(attr, source_ilvl_val)
                changed = True
            # A source paragraph can also carry its own explicit <w:ind>
            # overriding the numbering level's indent (e.g. a tighter hanging
            # indent for a bibliography-style list). ApplyNumberDefault()
            # resets the paragraph's indent to the list's own default when it
            # runs (confirmed live: it fires after ApplyParagraphProperties
            # already set the explicit indent, silently clobbering it), so
            # this override needs the same restoration as numId/ilvl above.
            source_ind = source_p.find("w:pPr/w:ind", namespaces=ns)
            output_p_pr = output_p.find("w:pPr", namespaces=ns)
            if source_ind is not None and output_p_pr is not None:
                output_ind = output_p_pr.find("w:ind", namespaces=ns)
                source_ind_xml = etree.tostring(source_ind)
                if output_ind is None:
                    output_p_pr.append(etree.fromstring(source_ind_xml))
                    changed = True
                elif etree.tostring(output_ind) != source_ind_xml:
                    index = list(output_p_pr).index(output_ind)
                    output_p_pr.remove(output_ind)
                    output_p_pr.insert(index, etree.fromstring(source_ind_xml))
                    changed = True
            if changed:
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        parts["word/numbering.xml"] = source_numbering
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_invisible_field_marker_run_formatting(output_path: Path, source_path: Path) -> int:
        """A run whose only content is a field-control marker (fldChar/
        instrText, no <w:t>) is invisible but can still carry its own
        formatting (e.g. bold) - this happens at the tail of a multi-
        paragraph field the renderer can never fully replay (see
        BlueprintCompiler._paragraph_closing_field_begins), where the
        closing fldChar end lands alone in its own paragraph. Since the
        renderer never creates that field at all, the paragraph ends up with
        zero runs instead of the one empty, formatted run source has. This
        restores that run's formatting shell only - never the fldChar/
        instrText content itself, which would misrepresent a field the
        renderer did not actually create.
        """
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(source_path) as source_archive:
            source_root = etree.fromstring(source_archive.read("word/document.xml"))
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}
        output_root = etree.fromstring(parts["word/document.xml"])
        source_paragraphs = source_root.xpath("//w:body//w:p", namespaces=ns)
        output_paragraphs = output_root.xpath("//w:body//w:p", namespaces=ns)
        restored = 0
        for source_paragraph, output_paragraph in zip(source_paragraphs, output_paragraphs):
            source_runs = source_paragraph.findall("w:r", namespaces=ns)
            if not source_runs or output_paragraph.findall("w:r", namespaces=ns):
                continue
            has_visible_text = any(
                (node.text or "").strip()
                for run in source_runs
                for node in run.findall("w:t", namespaces=ns)
            )
            if has_visible_text:
                continue
            for source_run in source_runs:
                r_pr = source_run.find("w:rPr", namespaces=ns)
                if r_pr is None:
                    continue
                new_run = etree.SubElement(output_paragraph, f"{{{namespace}}}r")
                new_run.append(deepcopy(r_pr))
                restored += 1
        if not restored:
            return 0
        parts["word/document.xml"] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _remove_word_added_note_reference_space(output_path: Path) -> int:
        """Word unconditionally inserts one leading space before the first
        text of every footnote/endnote it saves, regardless of what was
        actually typed there - confirmed live: typing text with no leading
        space at all still comes back from SaveAs2 with one added. When the
        source text already starts with its own space (common - authors often
        type a literal space after the reference mark), the result is two
        spaces. There is nothing to compare against the source for: Word's
        extra space is unconditional, so it is always safe to strip exactly
        one leading space from the first text immediately after a
        footnoteRef/endnoteRef marker.
        """
        from lxml import etree

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns = {"w": namespace}
        with ZipFile(output_path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        restored = 0
        roots: dict[str, object] = {}
        for part_name in ("word/footnotes.xml", "word/endnotes.xml"):
            if part_name not in parts:
                continue
            root = etree.fromstring(parts[part_name])
            changed = False
            for note in root.xpath("./w:footnote | ./w:endnote", namespaces=ns):
                for paragraph in note.xpath("./w:p", namespaces=ns):
                    runs = paragraph.xpath("./w:r", namespaces=ns)
                    def has_marker(run) -> bool:
                        return (
                            run.find(f"{{{namespace}}}footnoteRef") is not None
                            or run.find(f"{{{namespace}}}endnoteRef") is not None
                        )

                    marker_index = next(
                        (i for i, run in enumerate(runs) if has_marker(run)),
                        None,
                    )
                    if marker_index is None:
                        continue
                    for run in runs[marker_index + 1:]:
                        t = run.find(f"w:t", namespaces=ns)
                        if t is None:
                            continue
                        if t.text and t.text.startswith(" "):
                            t.text = t.text[1:]
                            restored += 1
                            changed = True
                        break
            if changed:
                roots[part_name] = root
        if not restored:
            return 0
        for part_name, root in roots.items():
            parts[part_name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return restored

    @staticmethod
    def _restore_source_compatibility_settings(output_path: Path, source_path: Path) -> int:
        """Replace only Word's normalized w:compat subtree with the source one."""
        from lxml import etree

        settings_name = "word/settings.xml"
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with ZipFile(source_path) as source_archive:
            if settings_name not in source_archive.namelist():
                return 0
            source_settings = source_archive.read(settings_name)
        with ZipFile(output_path) as output_archive:
            parts = {
                name: output_archive.read(name) for name in output_archive.namelist()
            }
        output_settings = parts.get(settings_name)
        if output_settings is None:
            return 0

        source_root = etree.fromstring(source_settings)
        output_root = etree.fromstring(output_settings)
        source_compat = source_root.find(f"{{{w_ns}}}compat")
        output_compat = output_root.find(f"{{{w_ns}}}compat")
        source_xml = etree.tostring(source_compat) if source_compat is not None else None
        output_xml = etree.tostring(output_compat) if output_compat is not None else None
        if source_xml == output_xml:
            return 0

        if output_compat is not None:
            insertion_index = output_root.index(output_compat)
            output_root.remove(output_compat)
        elif source_compat is not None:
            insertion_index = min(source_root.index(source_compat), len(output_root))
        else:
            insertion_index = len(output_root)
        if source_compat is not None:
            output_root.insert(insertion_index, deepcopy(source_compat))
        parts[settings_name] = etree.tostring(
            output_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{Path(output_path).name}.", suffix=".tmp", dir=Path(output_path).parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), Path(output_path))
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return 1

    @staticmethod
    def _restore_source_package_parts(output_path: Path, source_path: Path, model) -> int:
        """Restore source parts the parser explicitly classified for transfer.

        Interactive Word authoring starts with a new document, so Word cannot
        recreate opaque package attachments, thumbnails, or unreferenced media.
        The parser already separates those byte-preserved parts from authored
        content. Restore only missing, declared parts plus an otherwise-unused
        source numbering part, and recreate only their source relationships and
        content-type declarations. Existing output parts are never overwritten.
        """
        from lxml import etree

        content_type_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        output_path = Path(output_path)
        source_path = Path(source_path)
        with ZipFile(source_path) as source_archive:
            source_parts = {
                name: source_archive.read(name) for name in source_archive.namelist()
            }
        with ZipFile(output_path) as output_archive:
            parts = {name: output_archive.read(name) for name in output_archive.namelist()}

        desired = {
            str(name): part.data
            for name, part in getattr(model, "preserved_parts", {}).items()
            if str(name) in source_parts
        }
        if (
            getattr(model, "numbering_xml", None) is not None
            and "word/numbering.xml" in source_parts
        ):
            desired.setdefault("word/numbering.xml", source_parts["word/numbering.xml"])
        missing = {name: data for name, data in desired.items() if name not in parts}
        if not missing:
            return 0
        parts.update(missing)

        source_types = etree.fromstring(source_parts["[Content_Types].xml"])
        output_types = etree.fromstring(parts["[Content_Types].xml"])
        output_overrides = {
            node.get("PartName")
            for node in output_types.findall(f"{{{content_type_ns}}}Override")
        }
        output_defaults = {
            (node.get("Extension") or "").lower()
            for node in output_types.findall(f"{{{content_type_ns}}}Default")
        }
        source_overrides = {
            node.get("PartName"): node
            for node in source_types.findall(f"{{{content_type_ns}}}Override")
        }
        source_defaults = {
            (node.get("Extension") or "").lower(): node
            for node in source_types.findall(f"{{{content_type_ns}}}Default")
        }
        for name in missing:
            part_name = f"/{name}"
            override = source_overrides.get(part_name)
            if override is not None:
                if part_name not in output_overrides:
                    output_types.append(deepcopy(override))
                    output_overrides.add(part_name)
                continue
            extension = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            default = source_defaults.get(extension)
            if default is not None and extension not in output_defaults:
                output_types.append(deepcopy(default))
                output_defaults.add(extension)
        parts["[Content_Types].xml"] = etree.tostring(
            output_types, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        def owner_part(rels_name: str) -> str:
            if rels_name == "_rels/.rels":
                return ""
            rels_path = PurePosixPath(rels_name)
            return posixpath.join(
                str(rels_path.parent.parent), rels_path.name.removesuffix(".rels")
            )

        def restore_relationship(owner_rels: str, target_part: str, rel_type: str | None) -> None:
            source_rels_xml = source_parts.get(owner_rels)
            if source_rels_xml is None:
                return
            source_root = etree.fromstring(source_rels_xml)
            owner = owner_part(owner_rels)
            source_node = next(
                (
                    node
                    for node in source_root.findall(f"{{{relationship_ns}}}Relationship")
                    if (rel_type is None or node.get("Type") == rel_type)
                    and node.get("TargetMode") != "External"
                    and posixpath.normpath(
                        posixpath.join(posixpath.dirname(owner), node.get("Target") or "")
                    ) == target_part
                ),
                None,
            )
            if source_node is None:
                return
            if owner_rels in parts:
                output_root = etree.fromstring(parts[owner_rels])
            else:
                output_root = etree.Element(
                    f"{{{relationship_ns}}}Relationships", nsmap={None: relationship_ns}
                )
            if any(
                node.get("Type") == source_node.get("Type")
                and node.get("TargetMode") != "External"
                and posixpath.normpath(
                    posixpath.join(posixpath.dirname(owner), node.get("Target") or "")
                ) == target_part
                for node in output_root.findall(f"{{{relationship_ns}}}Relationship")
            ):
                return
            used_ids = {
                node.get("Id")
                for node in output_root.findall(f"{{{relationship_ns}}}Relationship")
            }
            index = 1
            while f"rId{index}" in used_ids:
                index += 1
            restored_node = deepcopy(source_node)
            restored_node.set("Id", f"rId{index}")
            output_root.append(restored_node)
            parts[owner_rels] = etree.tostring(
                output_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        for name, part in getattr(model, "preserved_parts", {}).items():
            if str(name) not in missing or not getattr(part, "relationship_type", None):
                continue
            restore_relationship(part.owner_rels, str(name), part.relationship_type)
        if "word/numbering.xml" in missing:
            restore_relationship(
                "word/_rels/document.xml.rels", "word/numbering.xml", None
            )

        fd, temporary_name = tempfile.mkstemp(
            prefix=f"{output_path.name}.", suffix=".tmp", dir=output_path.parent
        )
        os.close(fd)
        try:
            with ZipFile(temporary_name, "w", ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            _replace_with_retry(Path(temporary_name), output_path)
        finally:
            Path(temporary_name).unlink(missing_ok=True)
        return len(missing)

    @staticmethod
    def _load_save_history(path: Path) -> list[dict]:
        if not path.exists():
            return []
        rows: list[dict] = []
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip():
                rows.append(json.loads(line))
        return rows

    def _finalize_qa(self, prepared, output_path: Path, save_manager, audit, renderer) -> RunResult:
        paths = prepared.paths
        # The custom property states the number of real saves after this final save completes.
        renderer.set_custom_property("WordReplicaActualSaveCount", save_manager.next_sequence)
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        coordinator = InteractiveCheckpointCoordinator(
            project_id=paths.project_id, source_path=Path(prepared.source_path), blueprint=prepared.blueprint,
            output_path=output_path, settings=self._run_settings_dict(prepared.options),
            renderer=renderer, save_manager=save_manager, checkpoint_store=checkpoint_store,
        )
        final_checkpoint = coordinator.milestone(prepared.blueprint.total_events - 1, "final_output")
        # Word may add system/TOC bookmarks while saving. Close the owned Word
        # session before the deterministic package cleanup can rewrite output.
        close = getattr(renderer, "close", None)
        if callable(close):
            close()
        removed_header_shape_defaults = self._remove_unexpected_header_shape_defaults(
            output_path, Path(prepared.source_path)
        )
        removed_bookmarks = self._remove_unexpected_bookmarks(
            output_path,
            {str(bookmark.name) for bookmark in prepared.model.bookmarks},
        )
        expected_header_targets = {
            rel.target
            for section in prepared.model.sections
            for ref in (section.properties.get("header_refs") or [])
            for rel in [prepared.model.relationships.get(f"word/document.xml:{ref.get('rel_id')}")]
            if rel is not None and rel.target in prepared.model.headers
        }
        removed_headers = self._remove_unexpected_headers(output_path, expected_header_targets)
        restored_header_stories = self._restore_relationship_free_source_headers(
            output_path, Path(prepared.source_path), expected_header_targets
        )
        removed_template_spacing = self._remove_template_paragraph_spacing(
            output_path, Path(prepared.source_path)
        )
        restored_alignment = self._restore_explicit_paragraph_alignment(
            output_path, Path(prepared.source_path)
        )
        restored_run_character_spacing = self._restore_explicit_run_character_spacing(
            output_path, Path(prepared.source_path)
        )
        restored_run_font_names = self._restore_explicit_run_font_names(
            output_path, Path(prepared.source_path)
        )
        restored_drawing_effect_extents = self._restore_explicit_drawing_effect_extents(
            output_path, Path(prepared.source_path)
        )
        restored_run_segmentation = self._restore_source_run_segmentation(
            output_path, Path(prepared.source_path)
        )
        restored_empty_runs = self._restore_empty_runs(output_path, Path(prepared.source_path))
        restored_column_space = self._restore_explicit_column_space(
            output_path, Path(prepared.source_path)
        )
        restored_page_number_start = self._restore_explicit_page_number_start(
            output_path, Path(prepared.source_path)
        )
        restored_defaults = self._restore_source_document_defaults(
            output_path, Path(prepared.source_path)
        )
        restored_theme_style_latin_fonts = (
            self._restore_source_theme_style_latin_fonts(
                output_path, Path(prepared.source_path)
            )
        )
        expected_footer_types = [
            {str(ref.get("type", "default")) for ref in (section.properties.get("footer_refs") or [])}
            for section in prepared.model.sections
        ]
        restored_footer_stories = self._restore_source_footer_stories(
            output_path, Path(prepared.source_path), expected_footer_types
        )
        restored_footer_topology = self._restore_source_footer_topology(
            output_path, Path(prepared.source_path)
        )
        restored_cross_paragraph_field_shells = self._restore_source_cross_paragraph_field_shells(
            output_path, Path(prepared.source_path)
        )
        restored_field_instructions = self._restore_source_field_instructions(
            output_path, Path(prepared.source_path)
        )
        restored_table_layout = self._restore_source_table_layout(
            output_path, Path(prepared.source_path)
        )
        restored_numbering_definitions = self._restore_source_numbering_definitions(
            output_path, Path(prepared.source_path)
        )
        restored_invisible_field_marker_runs = self._restore_invisible_field_marker_run_formatting(
            output_path, Path(prepared.source_path)
        )
        removed_note_reference_spaces = self._remove_word_added_note_reference_space(output_path)
        restored_compatibility_settings = self._restore_source_compatibility_settings(
            output_path, Path(prepared.source_path)
        )
        restored_package_parts = self._restore_source_package_parts(
            output_path, Path(prepared.source_path), prepared.model
        )
        if removed_header_shape_defaults or removed_bookmarks or removed_headers or restored_header_stories or removed_template_spacing or restored_alignment or restored_run_character_spacing or restored_run_font_names or restored_drawing_effect_extents or restored_run_segmentation or restored_empty_runs or restored_column_space or restored_page_number_start or restored_defaults or restored_theme_style_latin_fonts or restored_footer_stories or restored_footer_topology or restored_cross_paragraph_field_shells or restored_field_instructions or restored_table_layout or restored_numbering_definitions or restored_invisible_field_marker_runs or removed_note_reference_spaces or restored_compatibility_settings or restored_package_parts:
            final_checkpoint = replace(
                final_checkpoint,
                output_sha256=sha256_file(output_path),
            )
            checkpoint_store.write(final_checkpoint)
        if removed_header_shape_defaults:
            audit.append(
                "UNEXPECTED_HEADER_SHAPE_DEFAULTS_REMOVED",
                {"count": removed_header_shape_defaults},
            )
        if removed_bookmarks:
            audit.append("UNEXPECTED_BOOKMARKS_REMOVED", {"count": removed_bookmarks})
        if removed_headers:
            audit.append("UNEXPECTED_HEADERS_REMOVED", {"count": removed_headers})
        if restored_header_stories:
            audit.append("SOURCE_HEADER_STORIES_RESTORED", {"count": restored_header_stories})
        if removed_template_spacing:
            audit.append("TEMPLATE_PARAGRAPH_SPACING_REMOVED", {"count": removed_template_spacing})
        if restored_alignment:
            audit.append("EXPLICIT_PARAGRAPH_ALIGNMENT_RESTORED", {"count": restored_alignment})
        if restored_run_character_spacing:
            audit.append("EXPLICIT_RUN_CHARACTER_SPACING_RESTORED", {"count": restored_run_character_spacing})
        if restored_run_font_names:
            audit.append("EXPLICIT_RUN_FONT_NAMES_RESTORED", {"count": restored_run_font_names})
        if restored_drawing_effect_extents:
            audit.append(
                "EXPLICIT_DRAWING_EFFECT_EXTENTS_RESTORED",
                {"count": restored_drawing_effect_extents},
            )
        if restored_run_segmentation:
            audit.append(
                "SOURCE_RUN_SEGMENTATION_RESTORED", {"count": restored_run_segmentation}
            )
        if restored_empty_runs:
            audit.append("EMPTY_RUNS_RESTORED", {"count": restored_empty_runs})
        if restored_column_space:
            audit.append("EXPLICIT_COLUMN_SPACE_RESTORED", {"count": restored_column_space})
        if restored_page_number_start:
            audit.append("EXPLICIT_PAGE_NUMBER_START_RESTORED", {"count": restored_page_number_start})
        if restored_defaults:
            audit.append("SOURCE_DOCUMENT_DEFAULTS_RESTORED", {"count": restored_defaults})
        if restored_theme_style_latin_fonts:
            audit.append(
                "SOURCE_THEME_STYLE_LATIN_FONTS_RESTORED",
                {"count": restored_theme_style_latin_fonts},
            )
        if restored_footer_stories:
            audit.append("SOURCE_FOOTER_STORIES_RESTORED", {"count": restored_footer_stories})
        if restored_footer_topology:
            audit.append(
                "SOURCE_FOOTER_TOPOLOGY_RESTORED", {"count": restored_footer_topology}
            )
        if restored_cross_paragraph_field_shells:
            audit.append(
                "SOURCE_CROSS_PARAGRAPH_FIELD_SHELLS_RESTORED",
                {"count": restored_cross_paragraph_field_shells},
            )
        if restored_field_instructions:
            audit.append("SOURCE_FIELD_INSTRUCTIONS_RESTORED", {"count": restored_field_instructions})
        if restored_table_layout:
            audit.append("SOURCE_TABLE_LAYOUT_RESTORED", {"count": restored_table_layout})
        if restored_numbering_definitions:
            audit.append("SOURCE_NUMBERING_DEFINITIONS_RESTORED", {"count": restored_numbering_definitions})
        if restored_invisible_field_marker_runs:
            audit.append(
                "INVISIBLE_FIELD_MARKER_RUN_FORMATTING_RESTORED",
                {"count": restored_invisible_field_marker_runs},
            )
        if removed_note_reference_spaces:
            audit.append("WORD_ADDED_NOTE_REFERENCE_SPACE_REMOVED", {"count": removed_note_reference_spaces})
        if restored_compatibility_settings:
            audit.append(
                "SOURCE_COMPATIBILITY_SETTINGS_RESTORED", {"count": restored_compatibility_settings}
            )
        if restored_package_parts:
            audit.append(
                "SOURCE_PACKAGE_PARTS_RESTORED", {"count": restored_package_parts}
            )
        warnings: list[WarningItem] = []
        for item in prepared.preflight.capability_items:
            if item.classification.value != "RECONSTRUCTED":
                warnings.append(WarningItem(
                    code=f"CAPABILITY_{item.classification.value}", message=item.reason,
                    element_id=item.source_element_id, affects_status=True,
                ))
        try:
            rebuilt_model = self.parser.parse(output_path) if hasattr(self.parser, "parse") else self.parser(output_path)
            bundle = run_l0_l3(prepared.model, rebuilt_model)
        except Exception as exc:
            if self.project_store is not None:
                self.project_store.set_status(paths.project_id, "FAILED")
            audit.append("INTERACTIVE_FINAL_QA_FAILED", {"reason": f"output parse failed: {exc}"})
            return RunResult(RunStatus.FAIL, output_path, None, project_id=paths.project_id,
                save_count=save_manager.sequence, warnings=warnings, reasons=[f"output parse failed: {exc}"])

        if self.run_l4_qa:
            source_pdf = paths.qa_dir / "l4_source.pdf"
            rebuilt_pdf = paths.qa_dir / "l4_rebuilt.pdf"
            try:
                self.pdf_exporter(Path(prepared.source_path), source_pdf, visible=False)
                self.pdf_exporter(output_path, rebuilt_pdf, visible=False)
                bundle.render = self.pdf_comparer(source_pdf, rebuilt_pdf, paths.qa_dir)
            except Exception as exc:
                warnings.append(WarningItem("L4_UNAVAILABLE", f"Controlled Word PDF comparison failed: {exc}", affects_status=True))
                bundle.render = None
        else:
            warnings.append(WarningItem(
                "L4_DEFERRED",
                "Controlled Word PDF comparison was deferred to the Golden audit",
                affects_status=False,
            ))
            bundle.render = None

        status = classify_run(bundle, warnings)
        qa_report = paths.qa_dir / "qa_report.html"
        saves = self._load_save_history(paths.logs_dir / "save_history.jsonl")
        write_qa_report(
            qa_report, status=status.value, source_sha256=prepared.blueprint.source_sha256,
            output_sha256=final_checkpoint.output_sha256, renderer="interactive-word",
            fidelity=prepared.options.interactive.fidelity.value, metadata=prepared.options.metadata.value,
            saves=saves, levels=bundle.levels, warnings=warnings, render_result=bundle.render,
            # The owned interactive Word session has already been closed above.
            # Starting a second DispatchEx instance only to read Version/Build
            # can hang under Word contention and adds no document-fidelity data.
            environment=capture_environment_fingerprint(include_word=False),
        )
        reasons: list[str] = []
        for level_name in ("L0", "L1"):
            level = bundle.levels.get(level_name)
            if level is not None and not level.passed:
                for finding in level.findings[:10]:
                    reasons.append(f"{getattr(finding, 'code', level_name)} at {getattr(finding, 'path', '/')}")
        project_status = "FAILED" if status is RunStatus.FAIL else "COMPLETED"
        if self.project_store is not None:
            self.project_store.set_status(paths.project_id, project_status)
        audit.append("INTERACTIVE_FINAL_QA_COMPLETED", {
            "status": status.value, "qa_report": str(qa_report),
            "l4_available": bundle.render is not None,
            "l4_deferred": not self.run_l4_qa,
            "l4_within_tolerance": getattr(bundle.render, "within_tolerance", None),
        })
        return RunResult(status, output_path, qa_report, project_id=paths.project_id,
            save_count=save_manager.sequence, warnings=warnings, reasons=reasons)

    def _execute(
        self,
        prepared: PreparedInteractiveRun,
        *,
        controller_factory,
        control,
        observer=None,
        resume_checkpoint=None,
    ) -> RunResult:
        paths = prepared.paths
        audit = AuditLog(paths.logs_dir / "audit.jsonl")
        settings = self._run_settings_dict(prepared.options)
        output_path = Path(resume_checkpoint.output_path) if resume_checkpoint is not None else (
            paths.output_dir / f"{Path(prepared.source_path).stem}_reconstructed.docx"
        )
        renderer = InteractiveWordRenderer(
            controller=controller_factory(),
            options=prepared.options.interactive,
        )
        assets = self._materialize_assets(prepared.model, paths.working_dir)
        renderer.controller.set_asset_resolver(lambda asset_id: assets[asset_id])
        save_manager = CheckpointManager(paths.logs_dir / "save_history.jsonl", audit)
        if resume_checkpoint is not None:
            save_manager.sequence = int(resume_checkpoint.save_sequence)
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        coordinator = InteractiveCheckpointCoordinator(
            project_id=paths.project_id,
            source_path=Path(prepared.source_path),
            blueprint=prepared.blueprint,
            output_path=output_path,
            settings=settings,
            renderer=renderer,
            save_manager=save_manager,
            checkpoint_store=checkpoint_store,
        )
        service_observer = _InteractiveServiceObserver(
            prepared.blueprint, control, audit, observer,
            live_verifier=self.live_verifier, source_model=prepared.model,
            output_path=output_path, coordinator=coordinator, settings=settings,
        )
        composite = _CompositeObserver(coordinator, service_observer)
        if self.project_store is not None:
            self.project_store.set_status(paths.project_id, "RUNNING")
        try:
            if resume_checkpoint is None:
                renderer.open_blank()
                start_index = 0
            else:
                renderer.open_existing(output_path)
                renderer.restore_checkpoint_state(resume_checkpoint)
                start_index = resume_checkpoint.last_completed_event_index + 1
            renderer.apply_metadata(prepared.model.extras.get("metadata_policy", {}))
            if resume_checkpoint is None:
                initial_checkpoint = coordinator.milestone(-1, "initial")
                service_observer.tracker.checkpoint_saved(
                    -1, initial_checkpoint.timestamp_utc
                )
            if start_index > 0:
                service_observer.prime_through(start_index - 1)
            audit.append("INTERACTIVE_EXECUTION_STARTED", {"start_index": start_index})
            outcome = renderer.execute_blueprint(
                prepared.blueprint, control, composite, start_index=start_index
            )
            audit.append("INTERACTIVE_METRICS", {
                "completed_events": service_observer.tracker.completed_events,
                "completed_characters": service_observer.tracker.completed_characters,
                "completed_tables": service_observer.tracker.completed_tables,
                "completed_images": service_observer.tracker.completed_images,
                "completed_sections": service_observer.tracker.completed_sections,
                "total_events": prepared.blueprint.total_events,
                "total_characters": prepared.blueprint.total_visible_characters,
                "outcome": outcome.status,
            })
            if outcome.status == "PAUSED_FIDELITY_MISMATCH":
                if coordinator.last_checkpoint_event_index != outcome.last_completed_index:
                    coordinator.milestone(outcome.last_completed_index, "fidelity_mismatch")
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "PAUSED")
                reasons = list(service_observer.halt_reasons) or ["Live fidelity verification mismatch"]
                audit.append("INTERACTIVE_EXECUTION_PAUSED_FIDELITY", {
                    "last_completed_event_index": outcome.last_completed_index, "reasons": reasons,
                })
                return RunResult(
                    status=RunStatus.WARN, output_path=output_path, qa_report_path=None,
                    project_id=paths.project_id, save_count=save_manager.sequence, reasons=reasons,
                )
            if outcome.status == "PAUSED_STATE_MISMATCH":
                coordinator.milestone(outcome.last_completed_index, "word_state_mismatch")
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "PAUSED")
                return RunResult(
                    status=RunStatus.WARN, output_path=output_path, qa_report_path=None,
                    project_id=paths.project_id, save_count=save_manager.sequence,
                    reasons=["Word state mismatch; reconstruction paused before the next event"],
                )
            if outcome.status == "STOPPED":
                coordinator.stop(outcome.last_completed_index)
                if self.project_store is not None:
                    self.project_store.set_status(paths.project_id, "STOPPED")
                audit.append("INTERACTIVE_EXECUTION_STOPPED", {"last_completed_event_index": outcome.last_completed_index})
                return RunResult(
                    status=RunStatus.WARN,
                    output_path=output_path,
                    qa_report_path=None,
                    project_id=paths.project_id,
                    save_count=save_manager.sequence,
                    reasons=["Interactive reconstruction stopped safely; resume checkpoint saved"],
                )
            if self.project_store is not None:
                self.project_store.set_status(paths.project_id, "VERIFYING")
            audit.append("INTERACTIVE_EXECUTION_COMPLETED", {"last_completed_event_index": outcome.last_completed_index})
            return self._finalize_qa(prepared, output_path, save_manager, audit, renderer)
        finally:
            renderer.close()

    def start(self, prepared: PreparedInteractiveRun, controller_factory=None, *, control=None, observer=None) -> RunResult:
        if not prepared.preflight.can_proceed:
            return RunResult(
                status=RunStatus.FAIL,
                output_path=None,
                qa_report_path=None,
                project_id=prepared.paths.project_id,
                reasons=list(prepared.preflight.blocking_reasons),
            )
        controller_factory = controller_factory or (
            lambda: InteractiveWordController(
                visible=prepared.options.visibility is VisibilityMode.VISIBLE
            )
        )
        control = control or InteractiveRunControl()
        if control.state.value == "CREATED":
            control.start()
        return self._execute(
            prepared,
            controller_factory=controller_factory,
            control=control,
            observer=observer,
        )

    def resume(self, project_id: str, control, observer=None, *, controller_factory=None) -> RunResult:
        if self.project_store is None:
            raise RuntimeError("resume requires a ProjectStore")
        paths = self.project_store.paths_for_project(project_id)
        project = self.project_store.get_project(project_id)
        source = Path(project["source_path"])
        checkpoint_store = InteractiveCheckpointStore(paths.logs_dir / "interactive_checkpoint.json")
        try:
            checkpoint = checkpoint_store.load_latest()
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=[f"checkpoint unreadable: {exc}"])
        if not source.exists() or sha256_file(source) != checkpoint.source_sha256:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=["source hash mismatch"])
        options = self._rebuild_options_from_settings(checkpoint.settings)
        interactive = options.interactive
        model = self._parse_model_for_options(source, options)
        blueprint = self._compile_blueprint(model, options)
        validation = checkpoint_store.validate_resume(checkpoint, source_path=source, blueprint=blueprint)
        if not validation.valid:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=[validation.reason])
        self._apply_metadata_policy(model, options)
        # Resume has already validated the saved preflight contract; re-run environment checks before Word opens.
        kwargs = {"font_probe": self.font_probe}
        if self.word_probe is not None:
            kwargs["word_probe"] = self.word_probe
        preflight = analyze_preflight(model, blueprint, interactive, **kwargs)
        if not preflight.can_proceed:
            return RunResult(RunStatus.FAIL, None, None, project_id=project_id, reasons=list(preflight.blocking_reasons))
        prepared = PreparedInteractiveRun(str(source.resolve()), model, blueprint, preflight, paths, options)
        controller_factory = controller_factory or (
            lambda: InteractiveWordController(
                visible=prepared.options.visibility is VisibilityMode.VISIBLE
            )
        )
        return self._execute(
            prepared,
            controller_factory=controller_factory,
            control=control,
            observer=observer,
            resume_checkpoint=checkpoint,
        )
