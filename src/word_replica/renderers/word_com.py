import sys
import tempfile
from pathlib import Path

from word_replica.domain.enums import RendererChoice
from word_replica.domain.errors import RendererUnavailableError
from word_replica.domain.model import DocumentModel, Paragraph, PreservedPart, Table
from word_replica.domain.results import WarningItem
from word_replica.renderers.base import RenderResult
from word_replica.renderers.pure_docx import PureDocxRenderer


def choose_renderer_name(choice: RendererChoice, word_is_available: bool) -> str:
    if choice is RendererChoice.DOCX:
        return "docx"
    if choice is RendererChoice.WORD and not word_is_available:
        raise RendererUnavailableError("Microsoft Word desktop automation is unavailable")
    if choice is RendererChoice.WORD:
        return "word"
    return "word" if word_is_available else "docx"


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    pythoncom = None
    app = None
    try:
        import pythoncom as _pythoncom
        import win32com.client

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Word.Application")
        return True
    except Exception:
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class WordSession:
    def __init__(self, visible: bool) -> None:
        self.visible = visible
        self.app = None
        self._owned_docs: list[object] = []
        self._pythoncom = None

    def __enter__(self) -> "WordSession":
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        try:
            self.app = win32com.client.DispatchEx("Word.Application")
            self.app.Visible = self.visible
            self.app.DisplayAlerts = 0
        except Exception:
            pythoncom.CoUninitialize()
            self._pythoncom = None
            raise
        return self

    def new_document(self):
        if self.app is None:
            raise RuntimeError("WordSession is not active")
        doc = self.app.Documents.Add()
        self._owned_docs.append(doc)
        return doc

    def open_document(self, path: Path, *, read_only: bool = True):
        if self.app is None:
            raise RuntimeError("WordSession is not active")
        doc = self.app.Documents.Open(
            str(path),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
        )
        self._owned_docs.append(doc)
        return doc

    def release_document(self, doc) -> None:
        if doc in self._owned_docs:
            self._owned_docs.remove(doc)

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            for doc in reversed(self._owned_docs):
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if self.app is not None:
                try:
                    self.app.Quit()
                except Exception:
                    pass
        finally:
            self._owned_docs.clear()
            self.app = None
            if self._pythoncom is not None:
                self._pythoncom.CoUninitialize()
                self._pythoncom = None


SAFE_PACKAGE_TRANSFER_REL_TYPES: frozenset[str] = frozenset()


class WordComRenderer:
    """Microsoft Word renderer with one code path for visible/background use.

    The renderer deliberately avoids keyboard simulation, sleeps, random timing,
    and any attempt to fabricate a human editing history. Every save is a real
    Word SaveAs2 call and shared save-count sealing is delegated to RenderContext.
    """

    def __init__(self, visible: bool):
        self.visible = visible
        self._warnings: list[WarningItem] = []
        self._stages: list[str] = []
        self._document = None
        self._post_save_transfers: list[PreservedPart] = []
        self._inserted_blocks = 0
        self._section_boundaries: dict[int, object] = {}

    @staticmethod
    def _twips_to_points(value) -> float | None:
        if value is None:
            return None
        try:
            return float(value) / 20.0
        except (TypeError, ValueError):
            return None

    def _apply_section_setup(self, setup, section) -> None:
        props = section.properties
        orientation = props.get("orientation")
        if orientation == "landscape":
            setup.Orientation = 1
        elif orientation == "portrait":
            setup.Orientation = 0
        mapping = {
            "width": "PageWidth",
            "height": "PageHeight",
            "margin_top": "TopMargin",
            "margin_right": "RightMargin",
            "margin_bottom": "BottomMargin",
            "margin_left": "LeftMargin",
            "header_distance": "HeaderDistance",
            "footer_distance": "FooterDistance",
            "gutter": "Gutter",
        }
        for key, attr in mapping.items():
            value = self._twips_to_points(props.get(key))
            if value is not None:
                try:
                    setattr(setup, attr, value)
                except Exception:
                    self._warnings.append(
                        WarningItem(
                            "WORD_SECTION_PROPERTY_UNSUPPORTED",
                            f"Could not apply section property {key}",
                            section.element_id,
                        )
                    )

    def _apply_sections(self, doc, model: DocumentModel) -> None:
        """Apply only the section Word already owns before body insertion.

        Creating every section up-front mutates the main story by inserting
        section-break paragraphs before body blocks are reconstructed. Section
        boundaries are therefore materialized only after their canonical
        boundary paragraph has been inserted.
        """
        if not model.sections:
            return
        try:
            self._apply_section_setup(doc.Sections(1).PageSetup, model.sections[0])
        finally:
            self._stages.append("sections")

    @staticmethod
    def _is_structural_section_boundary(block) -> bool:
        return (
            isinstance(block, Paragraph)
            and block.text() == ""
            and "section_index" in block.properties
            and set(block.properties).issubset({"section_index"})
        )

    def _materialize_section_boundary(self, doc, model: DocumentModel, section_index: int, boundary=None) -> None:
        target_number = section_index + 2
        if target_number > len(model.sections):
            self._warnings.append(WarningItem(
                "WORD_SECTION_BOUNDARY_INVALID",
                f"Section boundary {section_index} has no following section",
            ))
            return
        if doc.Sections.Count >= target_number:
            self._apply_section_setup(doc.Sections(target_number).PageSetup, model.sections[target_number - 1])
            return
        if boundary is None:
            position = max(int(doc.Content.Start), int(doc.Content.End) - 1)
            range_obj = doc.Range(position, position)
        else:
            range_obj = boundary.Range
        new_section = doc.Sections.Add(Range=range_obj, Start=2)
        self._apply_section_setup(new_section.PageSetup, model.sections[target_number - 1])

    def _materialize_sections(self, doc, model: DocumentModel) -> None:
        if len(model.sections) <= 1:
            return
        for section_index in range(len(model.sections) - 1):
            target_number = section_index + 2
            if doc.Sections.Count >= target_number:
                self._apply_section_setup(doc.Sections(target_number).PageSetup, model.sections[target_number - 1])
                continue
            boundary = self._section_boundaries.get(section_index)
            if boundary is None:
                self._warnings.append(WarningItem(
                    "WORD_SECTION_BOUNDARY_APPROXIMATION",
                    f"Section boundary {section_index} was not modeled; inserting the section at document end",
                    model.sections[section_index].element_id,
                ))
            self._materialize_section_boundary(doc, model, section_index, boundary=boundary)

    def _apply_metadata(self, doc, policy: dict[str, str]) -> None:
        built_in = {
            "title": "Title",
            "subject": "Subject",
            "keywords": "Keywords",
            "creator": "Author",
            "category": "Category",
        }
        for key, word_name in built_in.items():
            if key not in policy:
                continue
            try:
                doc.BuiltInDocumentProperties(word_name).Value = policy[key]
            except Exception:
                self._warnings.append(
                    WarningItem("WORD_METADATA_PROPERTY_UNSUPPORTED", f"Could not set {key}")
                )
        for key, value in policy.items():
            if key.startswith("custom:"):
                self.set_custom_property(key.split(":", 1)[1], value)
        self._stages.append("metadata")

    def _new_paragraph_after(self, doc, paragraph):
        paragraph.Range.InsertParagraphAfter()
        return doc.Paragraphs(doc.Paragraphs.Count)

    def _paragraph_target(self, doc):
        if self._inserted_blocks == 0 and doc.Paragraphs.Count >= 1:
            return doc.Paragraphs(1)
        last = doc.Paragraphs(doc.Paragraphs.Count)
        if hasattr(last.Range, "InsertParagraphAfter"):
            return self._new_paragraph_after(doc, last)
        return doc.Paragraphs.Add()

    @staticmethod
    def _word_break_type(break_type: str) -> int:
        # WdBreakType values from the Word object model.
        return {
            "line": 6,
            "textWrapping": 6,
            "page": 7,
            "column": 8,
            "clearLeft": 9,
            "clearRight": 10,
        }.get(break_type, 6)

    @staticmethod
    def _iter_run_segments(run):
        break_types = iter(run.properties.get("break_types", []))
        buffer: list[str] = []

        def flush_buffer():
            if buffer:
                text = "".join(buffer)
                buffer.clear()
                return text
            return None

        for char in run.text:
            if char == "\n":
                text = flush_buffer()
                if text is not None:
                    yield "text", text
                yield "break", next(break_types, "line")
            else:
                buffer.append(char)
        text = flush_buffer()
        if text is not None:
            yield "text", text

    def _apply_formatting_to_range(self, target, run) -> None:
        try:
            if "bold" in run.properties:
                target.Bold = -1 if run.properties["bold"] else 0
            if "italic" in run.properties:
                target.Italic = -1 if run.properties["italic"] else 0
            if "underline" in run.properties:
                target.Underline = 1 if run.properties["underline"] else 0
            if "hidden" in run.properties or run.hidden:
                target.Font.Hidden = -1 if (run.hidden or run.properties.get("hidden")) else 0
        except Exception:
            self._warnings.append(WarningItem(
                "WORD_RUN_FORMATTING_UNSUPPORTED",
                "Could not reproduce run-level formatting",
                run.element_id,
            ))

    def _clear_paragraph_content(self, doc, paragraph) -> None:
        start = int(paragraph.Range.Start)
        end = max(start, int(paragraph.Range.End) - 1)
        if end > start:
            doc.Range(start, end).Text = ""

    def _append_run_to_paragraph(self, doc, paragraph, run):
        created = []
        current = paragraph
        for segment_type, value in self._iter_run_segments(run):
            if segment_type == "break" and value == "page":
                try:
                    current = self._new_paragraph_after(doc, current)
                    current.Format.PageBreakBefore = -1
                    created.append(current)
                except Exception:
                    self._warnings.append(WarningItem(
                        "WORD_BREAK_RESTORE_FAILED",
                        "Could not reproduce page break as a continuation paragraph",
                        run.element_id,
                    ))
                continue

            insertion_point = max(int(current.Range.Start), int(current.Range.End) - 1)
            target = doc.Range(insertion_point, insertion_point)
            if segment_type == "break":
                try:
                    target.InsertBreak(Type=self._word_break_type(value))
                except Exception:
                    self._warnings.append(WarningItem(
                        "WORD_BREAK_RESTORE_FAILED",
                        f"Could not reproduce {value} break",
                        run.element_id,
                    ))
                continue
            target.InsertAfter(value)
            self._apply_formatting_to_range(target, run)
        return current, created

    def _apply_paragraph_semantics(self, paragraph, block: Paragraph, *, force_page_break: bool = False) -> None:
        if block.style_id:
            try:
                paragraph.Range.Style = block.style_id
            except Exception:
                self._warnings.append(
                    WarningItem(
                        "WORD_STYLE_UNAVAILABLE",
                        f"Could not apply style {block.style_id}",
                        block.element_id,
                    )
                )
        fmt = paragraph.Format
        props = block.properties
        if "keepNext" in props:
            fmt.KeepWithNext = -1 if props["keepNext"] else 0
        if force_page_break:
            fmt.PageBreakBefore = -1
        elif "pageBreakBefore" in props:
            fmt.PageBreakBefore = -1 if props["pageBreakBefore"] else 0
        if "keepLines" in props:
            fmt.KeepTogether = -1 if props["keepLines"] else 0

    def _insert_paragraph(self, doc, block: Paragraph) -> None:
        paragraph = self._paragraph_target(doc)
        self._clear_paragraph_content(doc, paragraph)
        touched = [(paragraph, False)]
        current = paragraph
        for run in block.runs:
            current, created = self._append_run_to_paragraph(doc, current, run)
            touched.extend((item, True) for item in created)
        for item, is_page_continuation in touched:
            self._apply_paragraph_semantics(item, block, force_page_break=is_page_continuation)
        if "section_index" in block.properties:
            try:
                self._section_boundaries[int(block.properties["section_index"])] = current
            except (TypeError, ValueError):
                self._warnings.append(WarningItem(
                    "WORD_SECTION_BOUNDARY_INVALID",
                    "Canonical section boundary index is not an integer",
                    block.element_id,
                ))
        self._inserted_blocks += 1

    def _insert_table(self, doc, block: Table) -> None:
        rows = max(1, len(block.rows))
        logical_cols = 1
        for row in block.rows:
            logical_cols = max(
                logical_cols,
                sum(max(1, int(cell.properties.get("grid_span", 1))) for cell in row.cells),
            )
        insert_at = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        table = doc.Tables.Add(insert_at, rows, logical_cols)

        # Capture every Word Cell reference and populate text before performing
        # any merge. Word mutates row/cell collections immediately after a
        # merge, so later Table.Cell(row, column) lookups can become invalid.
        horizontal_merges: list[tuple[object, object, str]] = []
        vertical_open: dict[int, tuple[object, object, str]] = {}
        vertical_merges: list[tuple[object, object, str]] = []

        for r_index, row in enumerate(block.rows, start=1):
            c_index = 1
            continued_columns: set[int] = set()
            for cell in row.cells:
                span = max(1, int(cell.properties.get("grid_span", 1)))
                start_cell = table.Cell(r_index, c_index)
                text_parts = []
                for child in cell.blocks:
                    if isinstance(child, Paragraph):
                        text_parts.append(child.text())
                start_cell.Range.Text = "\r".join(text_parts)

                if span > 1:
                    end_cell = table.Cell(r_index, c_index + span - 1)
                    horizontal_merges.append((start_cell, end_cell, cell.element_id))

                merge_kind = cell.properties.get("v_merge")
                if merge_kind == "restart":
                    previous = vertical_open.pop(c_index, None)
                    if previous is not None and previous[0] is not previous[1]:
                        vertical_merges.append(previous)
                    vertical_open[c_index] = (start_cell, start_cell, cell.element_id)
                    continued_columns.add(c_index)
                elif merge_kind == "continue":
                    current = vertical_open.get(c_index)
                    if current is None:
                        self._warnings.append(WarningItem(
                            "WORD_VERTICAL_MERGE_INVALID",
                            "Vertical merge continuation has no restart cell",
                            cell.element_id,
                        ))
                    else:
                        vertical_open[c_index] = (current[0], start_cell, current[2])
                    continued_columns.add(c_index)
                else:
                    current = vertical_open.pop(c_index, None)
                    if current is not None and current[0] is not current[1]:
                        vertical_merges.append(current)
                c_index += span

            # A vertical run ends when a following row does not continue it.
            for column in list(vertical_open):
                if column not in continued_columns:
                    current = vertical_open.pop(column)
                    if current[0] is not current[1]:
                        vertical_merges.append(current)

        for current in vertical_open.values():
            if current[0] is not current[1]:
                vertical_merges.append(current)

        # Use the captured COM objects only; do not query Table.Cell after the
        # first merge mutates Word's collection.
        for start_cell, end_cell, element_id in vertical_merges + horizontal_merges:
            try:
                start_cell.Merge(end_cell)
            except Exception:
                self._warnings.append(WarningItem(
                    "WORD_CELL_MERGE_UNSUPPORTED",
                    "Could not reproduce a merged table cell",
                    element_id,
                ))
        self._inserted_blocks += 1

    def _insert_block(self, doc, block) -> None:
        if isinstance(block, Paragraph):
            self._insert_paragraph(doc, block)
        elif isinstance(block, Table):
            self._insert_table(doc, block)
        else:
            self._warnings.append(
                WarningItem("UNSUPPORTED_BLOCK", f"Unsupported block type: {type(block).__name__}")
            )

    def _insert_assets(self, doc, model: DocumentModel) -> None:
        inserted = 0
        for asset in model.assets.values():
            suffix = Path(asset.part_name).suffix or ".bin"
            temp_path: Path | None = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as handle:
                    handle.write(asset.bytes_data)
                    temp_path = Path(handle.name)
                target = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                doc.InlineShapes.AddPicture(str(temp_path), False, True, target)
                inserted += 1
            except Exception:
                self._warnings.append(WarningItem(
                    "WORD_ASSET_RESTORE_FAILED",
                    f"Could not embed media asset {asset.part_name}",
                ))
            finally:
                if temp_path is not None:
                    temp_path.unlink(missing_ok=True)
        if inserted:
            self._warnings.append(WarningItem(
                "WORD_ASSET_POSITION_APPROXIMATION",
                "Media bytes were embedded, but exact original inline/anchor positions are not yet represented in the canonical model",
            ))
        self._stages.append("assets")

    @staticmethod
    def _blocks_text(blocks) -> str:
        parts: list[str] = []
        for block in blocks:
            if isinstance(block, Paragraph):
                parts.append(block.text())
        return "\r".join(parts)

    def _restore_headers_footers(self, doc, model: DocumentModel) -> None:
        header_text = self._blocks_text(next(iter(model.headers.values()), []))
        footer_text = self._blocks_text(next(iter(model.footers.values()), []))
        if header_text or footer_text:
            for index in range(1, doc.Sections.Count + 1):
                section = doc.Sections(index)
                if header_text:
                    try:
                        section.Headers(1).Range.Text = header_text
                    except Exception:
                        self._warnings.append(WarningItem("WORD_HEADER_RESTORE_FAILED", "Could not restore header"))
                if footer_text:
                    try:
                        section.Footers(1).Range.Text = footer_text
                    except Exception:
                        self._warnings.append(WarningItem("WORD_FOOTER_RESTORE_FAILED", "Could not restore footer"))
        self._stages.append("headers_footers")

    def _restore_notes(self, doc, model: DocumentModel) -> None:
        restored = 0
        for note_id, blocks in model.footnotes.items():
            if str(note_id).startswith("-"):
                continue
            text = self._blocks_text(blocks)
            try:
                anchor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                doc.Footnotes.Add(Range=anchor, Text=text)
                restored += 1
            except Exception:
                self._warnings.append(WarningItem("WORD_FOOTNOTE_RESTORE_FAILED", f"Could not restore footnote {note_id}"))
        for note_id, blocks in model.endnotes.items():
            if str(note_id).startswith("-"):
                continue
            text = self._blocks_text(blocks)
            try:
                anchor = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                doc.Endnotes.Add(Range=anchor, Text=text)
                restored += 1
            except Exception:
                self._warnings.append(WarningItem("WORD_ENDNOTE_RESTORE_FAILED", f"Could not restore endnote {note_id}"))
        if restored:
            self._warnings.append(WarningItem(
                "WORD_NOTE_REFERENCE_POSITION_APPROXIMATION",
                "Footnote/endnote content was recreated, but exact source reference positions are not yet modeled",
            ))
        self._stages.append("notes")

    def _restore_fields_bookmarks(self, doc, model: DocumentModel) -> None:
        bookmark_restored = 0
        for bookmark in model.bookmarks:
            try:
                anchor = doc.Range(doc.Content.Start, doc.Content.Start)
                doc.Bookmarks.Add(Name=bookmark.name, Range=anchor)
                bookmark_restored += 1
            except Exception:
                self._warnings.append(WarningItem("WORD_BOOKMARK_RESTORE_FAILED", f"Could not restore bookmark {bookmark.name}", bookmark.bookmark_id))
        field_restored = 0
        for field in model.fields:
            try:
                target = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                doc.Fields.Add(Range=target, Type=-1, Text=field.instruction, PreserveFormatting=True)
                field_restored += 1
            except Exception:
                self._warnings.append(WarningItem("WORD_FIELD_RESTORE_FAILED", f"Could not restore field {field.instruction}", field.field_id))
        if field_restored:
            self._warnings.append(WarningItem(
                "WORD_FIELD_POSITION_APPROXIMATION",
                "Field instructions were recreated, but exact source field ranges are not yet modeled",
            ))
        if bookmark_restored:
            self._warnings.append(WarningItem(
                "WORD_BOOKMARK_POSITION_APPROXIMATION",
                "Bookmark definitions were recreated, but exact source ranges are not yet modeled",
            ))
        self._stages.append("fields_bookmarks")

    def _warn_review_structures(self, model: DocumentModel) -> None:
        if model.comments or model.revisions or model.extras.get("tracked_changes_enabled"):
            self._warnings.append(WarningItem(
                "WORD_REVIEW_STRUCTURES_APPROXIMATION",
                "Comments/tracked-revision provenance remains canonical evidence but exact Word review ranges are not yet reconstructed",
            ))

    def _handle_preserved_part(self, part: PreservedPart) -> None:
        if part.relationship_type in SAFE_PACKAGE_TRANSFER_REL_TYPES:
            self._post_save_transfers.append(part)
            return
        code = (
            "OLE_TRANSFER_UNSUPPORTED"
            if "oleObject" in (part.relationship_type or "")
            else "EMBEDDED_PART_UNSUPPORTED"
        )
        self._warnings.append(WarningItem(code=code, message=f"Could not safely transfer {part.part_name}"))

    def save(self, output_path: Path) -> None:
        if self._document is None:
            raise RuntimeError("No active Word document")
        self._document.SaveAs2(str(output_path))

    def set_custom_property(self, name: str, value: str) -> None:
        if self._document is None:
            raise RuntimeError("No active Word document")
        props = self._document.CustomDocumentProperties
        try:
            props(name).Value = value
        except Exception:
            # msoPropertyTypeString = 4
            props.Add(name, False, 4, str(value))

    def _validate_package_with_word(self, output_path: Path) -> None:
        """Open the reconstructed package in real Word without mutating it.

        Canonical OOXML reconstruction is intentionally performed by the
        deterministic package engine; Word is then used as the compatibility
        gate. Keeping the validation session read-only avoids Word normalizing
        away unreferenced-but-preserved package parts while still proving that
        desktop Word can open and paginate the result.
        """
        with WordSession(visible=self.visible) as session:
            doc = session.open_document(Path(output_path), read_only=True)
            try:
                try:
                    doc.Repaginate()
                except Exception:
                    # Opening successfully is the compatibility gate. Some Word
                    # builds reject Repaginate on a read-only/background doc.
                    pass
            finally:
                try:
                    doc.Close(False)
                finally:
                    session.release_document(doc)

    def render(self, model: DocumentModel, output_path: Path, context) -> RenderResult:
        """Hybrid renderer: deterministic canonical rebuild + real Word gate.

        Earlier v1 builds reconstructed every block through COM. That exposed
        Word collection/range mutation quirks (merged cells, section breaks,
        page-break paragraph marks) and made otherwise-valid documents fail.
        The release path now uses the same canonical package builder that is
        covered by the full 13-document corpus, then opens the exact resulting
        bytes in Microsoft Word read-only. No fake typing, delays, or fabricated
        metadata are introduced; all recorded saves remain real package writes.
        """
        output_path = Path(output_path)
        canonical = PureDocxRenderer()
        result = canonical.render(model, output_path, context)
        self._warnings = list(result.warnings)
        self._stages = list(result.stages_completed)
        self._validate_package_with_word(output_path)
        self._stages.append("word_validation")
        return RenderResult(
            output_path=output_path,
            warnings=list(self._warnings),
            stages_completed=list(self._stages),
        )
