from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import queue
import threading
from typing import Any

from word_replica.config import InteractiveOptions, RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    InteractiveFidelity,
    InteractiveRunState,
    InteractiveSpeedMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
    RunStatus,
    VisibilityMode,
)
from word_replica.services.project_store import ProjectStore
from word_replica.services.rebuild import RebuildService


AUTOMATION_DISCLOSURE = (
    "Interactive Reconstruction je automatizirani unos znak-po-znak i objekt-po-objekt; "
    "proces je transparentno automatiziran i nije dokaz za ručno autorstvo niti izmišljenu povijest uređivanja."
)


@dataclass(slots=True)
class DesktopState:
    source_path: Path | None = None
    reconstruction_mode: str = "instant"
    renderer: str = "auto"
    visibility: str = "background"
    fidelity: str = "clean"
    metadata: str = "fresh"
    allow_source_overwrite: bool = False
    preserve_author_fields: bool = False
    custom_metadata_allowlist: tuple[str, ...] = ()
    interactive_speed_mode: str = "fast"
    interactive_characters_per_second: float = 25.0
    interactive_object_step_delay_ms: int = 150
    interactive_fidelity: str = "maximum"
    checkpoint_after_tables: bool = True
    checkpoint_after_images: bool = True
    checkpoint_after_sections: bool = True
    checkpoint_event_interval: int = 500
    verify_during_run: bool = True
    block_on_unsupported: bool = True
    allow_preserved_objects: bool = False


def options_from_desktop_state(state: DesktopState) -> RebuildOptions:
    mode = ReconstructionMode(state.reconstruction_mode)
    renderer = RendererChoice.WORD if mode is ReconstructionMode.INTERACTIVE else RendererChoice(state.renderer)
    visibility = VisibilityMode.VISIBLE if mode is ReconstructionMode.INTERACTIVE else VisibilityMode(state.visibility)
    interactive = InteractiveOptions(
        speed_mode=InteractiveSpeedMode(state.interactive_speed_mode),
        characters_per_second=float(state.interactive_characters_per_second),
        object_step_delay_ms=int(state.interactive_object_step_delay_ms),
        fidelity=InteractiveFidelity(state.interactive_fidelity),
        checkpoint_after_tables=bool(state.checkpoint_after_tables),
        checkpoint_after_images=bool(state.checkpoint_after_images),
        checkpoint_after_sections=bool(state.checkpoint_after_sections),
        checkpoint_event_interval=int(state.checkpoint_event_interval),
        verify_during_run=bool(state.verify_during_run),
        block_on_unsupported=bool(state.block_on_unsupported),
        allow_preserved_objects=bool(state.allow_preserved_objects),
    )
    return RebuildOptions(
        renderer=renderer,
        visibility=visibility,
        fidelity=FidelityMode(state.fidelity),
        metadata=MetadataMode(state.metadata),
        allow_source_overwrite=state.allow_source_overwrite,
        preserve_author_fields=state.preserve_author_fields,
        custom_metadata_allowlist=state.custom_metadata_allowlist,
        reconstruction_mode=mode,
        interactive=interactive,
    )


def progress_view_model(progress) -> dict[str, object]:
    percent = 0.0 if not progress.total_events else (progress.completed_events / progress.total_events) * 100.0
    location_parts = [f"story {progress.location.story}", f"sekcija {progress.location.section_index + 1}"]
    if progress.location.row_index is not None:
        location_parts.append(f"red {progress.location.row_index}")
    if progress.location.cell_index is not None:
        location_parts.append(f"ćelija {progress.location.cell_index}")
    return {
        "percent": percent,
        "status": (
            f"Znakovi {progress.completed_characters}/{progress.total_characters} · "
            f"tablice {progress.completed_tables} · slike {progress.completed_images} · sekcije {progress.completed_sections}"
        ),
        "location": " · ".join(location_parts),
        "verification": progress.verification_status,
    }


def update_interactive_control_speed(control, mode: str, characters_per_second: float) -> None:
    speed_mode = InteractiveSpeedMode(mode)
    cps = float(characters_per_second) if speed_mode is InteractiveSpeedMode.CUSTOM else None
    control.set_speed(speed_mode, cps)


class DesktopInteractiveObserver:
    """Worker-thread observer. It only writes plain payloads to a Queue; never touches Tk."""

    def __init__(self, event_queue) -> None:
        self.event_queue = event_queue

    def preflight(self, report) -> None:
        self.event_queue.put(("preflight", report))

    def progress(self, progress) -> None:
        self.event_queue.put(("progress", progress))

    def state_mismatch(self, index, event, expected, actual) -> None:
        self.event_queue.put(("state", {
            "status": "WORD_STATE_MISMATCH",
            "event_index": index,
            "source_element_id": event.source_element_id,
        }))


def validate_source(state: DesktopState) -> tuple[bool, str]:
    source = state.source_path
    if source is None:
        return False, "Odaberi .docx dokument."
    if source.suffix.lower() != ".docx":
        return False, "Word Replica v1 podržava samo .docx dokumente."
    if not source.exists():
        return False, "Odabrani .docx dokument više ne postoji."
    return True, ""


def _open_path(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    raise RuntimeError("Otvaranje putanje iz GUI-ja podržano je u Windows izdanju.")


class WordReplicaDesktop:
    POLL_MS = 125

    def __init__(self, root, service: RebuildService | None = None) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.ttk = ttk
        self.root = root
        self.service = service or RebuildService.default()
        self.store = ProjectStore()
        self.state = DesktopState()
        self._events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._running = False
        self._last_result = None
        self._last_project_root: Path | None = None
        self._interactive_control = None
        self._interactive_observer = DesktopInteractiveObserver(self._events)
        self._resume_project_id: str | None = None

        root.title("Word Replica")
        root.geometry("980x860")
        root.minsize(840, 700)

        self.source_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="Instant Reconstruction")
        self.renderer_var = tk.StringVar(value="Automatski")
        self.visibility_var = tk.StringVar(value="U pozadini")
        self.fidelity_var = tk.StringVar(value="Čista replika")
        self.metadata_var = tk.StringVar(value="Novi metapodaci")
        self.preserve_author_var = tk.BooleanVar(value=False)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.custom_properties_var = tk.StringVar()

        self.speed_var = tk.StringVar(value="Fast")
        self.custom_speed_var = tk.DoubleVar(value=25.0)
        self.object_delay_var = tk.IntVar(value=150)
        self.interactive_fidelity_var = tk.StringVar(value="Maximum Fidelity")
        self.cp_tables_var = tk.BooleanVar(value=True)
        self.cp_images_var = tk.BooleanVar(value=True)
        self.cp_sections_var = tk.BooleanVar(value=True)
        self.cp_interval_var = tk.IntVar(value=500)
        self.verify_live_var = tk.BooleanVar(value=True)
        self.allow_preserved_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Spremno")
        self.saves_var = tk.StringVar(value="Stvarna spremanja: 0")
        self.semantic_var = tk.StringVar(value="")
        self.verification_var = tk.StringVar(value="Verifikacija: nije pokrenuta")
        self.preflight_var = tk.StringVar(value="Preflight: nije pokrenut")
        self.resume_var = tk.StringVar(value="Nema nedovršene rekonstrukcije.")

        outer = ttk.Frame(root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Word Replica", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Lokalna DOCX rekonstrukcija: stabilni Instant način ili vidljivi Interactive znak-po-znak način.",
        ).pack(anchor="w", pady=(2, 14))

        source_box = ttk.LabelFrame(outer, text="1. Dokument", padding=12)
        source_box.pack(fill="x")
        source_box.columnconfigure(0, weight=1)
        self.source_entry = ttk.Entry(source_box, textvariable=self.source_var, state="readonly")
        self.source_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.browse_button = ttk.Button(source_box, text="Odaberi DOCX", command=self.choose_source)
        self.browse_button.grid(row=0, column=1)

        self.resume_box = ttk.LabelFrame(outer, text="Nedovršena Interactive rekonstrukcija", padding=10)
        self.resume_box.pack(fill="x", pady=(8, 0))
        ttk.Label(self.resume_box, textvariable=self.resume_var).pack(side="left", fill="x", expand=True)
        self.resume_button = ttk.Button(self.resume_box, text="Nastavi", command=self.resume_latest, state="disabled")
        self.resume_button.pack(side="right")

        mode_box = ttk.LabelFrame(outer, text="2. Način rekonstrukcije", padding=12)
        mode_box.pack(fill="x", pady=(12, 0))
        mode_box.columnconfigure(0, weight=1)
        ttk.Label(mode_box, text="Mode").grid(row=0, column=0, sticky="w")
        self.mode_combo = ttk.Combobox(
            mode_box,
            textvariable=self.mode_var,
            values=("Instant Reconstruction", "Interactive Reconstruction"),
            state="readonly",
        )
        self.mode_combo.grid(row=1, column=0, sticky="ew")
        ttk.Label(
            mode_box,
            text="Interactive otvara novi Microsoft Word i stvarno gradi sadržaj znak-po-znak i objekt-po-objekt.",
        ).grid(row=2, column=0, sticky="w", pady=(5, 0))

        self.interactive_box = ttk.LabelFrame(outer, text="3. Interactive opcije", padding=12)
        for c in range(4):
            self.interactive_box.columnconfigure(c, weight=1)
        ttk.Label(self.interactive_box, text="Brzina").grid(row=0, column=0, sticky="w")
        self.speed_combo = ttk.Combobox(
            self.interactive_box,
            textvariable=self.speed_var,
            values=("Slow", "Fast", "Custom", "Maximum"),
            state="readonly",
        )
        self.speed_combo.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Label(self.interactive_box, text="Custom znakova/s").grid(row=0, column=1, sticky="w")
        self.custom_speed_spin = ttk.Spinbox(self.interactive_box, from_=1, to=1000, increment=1, textvariable=self.custom_speed_var)
        self.custom_speed_spin.grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Label(self.interactive_box, text="Delay objekta (ms)").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(self.interactive_box, from_=0, to=5000, increment=25, textvariable=self.object_delay_var).grid(row=1, column=2, sticky="ew", padx=6)
        ttk.Label(self.interactive_box, text="Interactive fidelity").grid(row=0, column=3, sticky="w")
        ttk.Combobox(
            self.interactive_box,
            textvariable=self.interactive_fidelity_var,
            values=("Standard", "Maximum Fidelity"),
            state="readonly",
        ).grid(row=1, column=3, sticky="ew", padx=(6, 0))

        cp_frame = ttk.Frame(self.interactive_box)
        cp_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(cp_frame, text="Checkpoint nakon tablice", variable=self.cp_tables_var).pack(side="left")
        ttk.Checkbutton(cp_frame, text="nakon slike", variable=self.cp_images_var).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(cp_frame, text="nakon sekcije", variable=self.cp_sections_var).pack(side="left", padx=(10, 0))
        ttk.Label(cp_frame, text="svakih N eventa:").pack(side="left", padx=(16, 4))
        ttk.Spinbox(cp_frame, from_=25, to=10000, increment=25, textvariable=self.cp_interval_var, width=8).pack(side="left")
        ttk.Checkbutton(cp_frame, text="Live verify", variable=self.verify_live_var).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(cp_frame, text="Dopusti PRESERVED uz upozorenje", variable=self.allow_preserved_var).pack(side="left", padx=(14, 0))

        technical = ttk.LabelFrame(outer, text="4. Fidelity, metapodaci i Instant Advanced", padding=12)
        technical.pack(fill="x", pady=(12, 0))
        for c in range(4):
            technical.columnconfigure(c, weight=1)
        ttk.Label(technical, text="Instant renderer").grid(row=0, column=0, sticky="w")
        ttk.Label(technical, text="Instant Word prikaz").grid(row=0, column=1, sticky="w")
        ttk.Label(technical, text="Instant fidelity").grid(row=0, column=2, sticky="w")
        ttk.Label(technical, text="Metapodaci").grid(row=0, column=3, sticky="w")
        self.renderer_combo = ttk.Combobox(technical, textvariable=self.renderer_var, values=("Automatski", "Microsoft Word", "Pure DOCX"), state="readonly")
        self.visibility_combo = ttk.Combobox(technical, textvariable=self.visibility_var, values=("U pozadini", "Vidljivo"), state="readonly")
        self.fidelity_combo = ttk.Combobox(technical, textvariable=self.fidelity_var, values=("Čista replika", "Puna vjernost"), state="readonly")
        self.metadata_combo = ttk.Combobox(technical, textvariable=self.metadata_var, values=("Novi metapodaci", "Sačuvaj legitimne"), state="readonly")
        for c, widget in enumerate((self.renderer_combo, self.visibility_combo, self.fidelity_combo, self.metadata_combo)):
            widget.grid(row=1, column=c, sticky="ew", padx=(0 if c == 0 else 5, 5 if c < 3 else 0))
        self.preserve_author_check = ttk.Checkbutton(technical, text="Sačuvaj legitimna opisna polja autora/tvrtke", variable=self.preserve_author_var)
        self.preserve_author_check.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.overwrite_check = ttk.Checkbutton(technical, text="Instant: dopusti zamjenu izvornika nakon verificiranog backupa", variable=self.overwrite_var)
        self.overwrite_check.grid(row=2, column=2, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(technical, text="Custom properties za očuvanje (odvojene zarezom)").grid(row=3, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.custom_entry = ttk.Entry(technical, textvariable=self.custom_properties_var)
        self.custom_entry.grid(row=4, column=0, columnspan=4, sticky="ew")

        run_box = ttk.LabelFrame(outer, text="5. Rekonstrukcija", padding=12)
        run_box.pack(fill="both", expand=True, pady=(12, 0))
        run_box.columnconfigure(0, weight=1)
        run_box.rowconfigure(6, weight=1)
        controls = ttk.Frame(run_box)
        controls.grid(row=0, column=0, sticky="ew")
        self.start_button = ttk.Button(controls, text="Pokreni rekonstrukciju", command=self.start_rebuild)
        self.start_button.pack(side="left")
        self.pause_button = ttk.Button(controls, text="Pause", command=self.toggle_pause, state="disabled")
        self.pause_button.pack(side="left", padx=(8, 0))
        self.stop_button = ttk.Button(controls, text="Stop", command=self.stop_interactive, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(run_box, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(10, 4))
        status_row = ttk.Frame(run_box)
        status_row.grid(row=2, column=0, sticky="ew")
        ttk.Label(status_row, textvariable=self.status_var).pack(side="left")
        ttk.Label(status_row, textvariable=self.saves_var).pack(side="right")
        ttk.Label(run_box, textvariable=self.preflight_var).grid(row=3, column=0, sticky="w")
        ttk.Label(run_box, textvariable=self.semantic_var).grid(row=4, column=0, sticky="w")
        ttk.Label(run_box, textvariable=self.verification_var).grid(row=5, column=0, sticky="w")
        self.log = tk.Text(run_box, height=9, wrap="word", state="disabled", font=("Consolas", 9))
        self.log.grid(row=6, column=0, sticky="nsew", pady=(8, 8))

        actions = ttk.Frame(run_box)
        actions.grid(row=7, column=0, sticky="w")
        self.open_doc_button = ttk.Button(actions, text="Otvori gotovi DOCX", command=self.open_output, state="disabled")
        self.open_doc_button.pack(side="left")
        self.open_report_button = ttk.Button(actions, text="Otvori QA izvještaj", command=self.open_report, state="disabled")
        self.open_report_button.pack(side="left", padx=(8, 0))
        self.open_project_button = ttk.Button(actions, text="Otvori projektni folder", command=self.open_project, state="disabled")
        self.open_project_button.pack(side="left", padx=(8, 0))

        ttk.Label(outer, text=AUTOMATION_DISCLOSURE, font=("Segoe UI", 8), wraplength=930).pack(anchor="w", pady=(10, 0))

        self.mode_combo.bind("<<ComboboxSelected>>", lambda _event: self._sync_mode_state())
        self.renderer_combo.bind("<<ComboboxSelected>>", lambda _event: self._sync_visibility_state())
        self.metadata_combo.bind("<<ComboboxSelected>>", lambda _event: self._sync_metadata_state())
        self.speed_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_speed_changed())
        self.custom_speed_spin.bind("<Return>", lambda _event: self._on_speed_changed())
        self._sync_mode_state()
        self._sync_metadata_state()
        self._refresh_resume_card()
        root.after(self.POLL_MS, self._poll_events)

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def choose_source(self) -> None:
        from tkinter import filedialog
        selected = filedialog.askopenfilename(title="Odaberi Word dokument", filetypes=(("Word dokument", "*.docx"),))
        if selected:
            self.state.source_path = Path(selected)
            self.source_var.set(selected)
            self._append_log(f"Odabran izvor: {selected}")

    def _sync_mode_state(self) -> None:
        interactive = self.mode_var.get() == "Interactive Reconstruction"
        if interactive:
            if not self.interactive_box.winfo_ismapped():
                self.interactive_box.pack(fill="x", pady=(12, 0), before=self.renderer_combo.master)
            self.renderer_var.set("Microsoft Word")
            self.visibility_var.set("Vidljivo")
            self.renderer_combo.configure(state="disabled")
            self.visibility_combo.configure(state="disabled")
            self.overwrite_var.set(False)
            self.overwrite_check.configure(state="disabled")
        else:
            if self.interactive_box.winfo_ismapped():
                self.interactive_box.pack_forget()
            self.renderer_combo.configure(state="readonly")
            self.overwrite_check.configure(state="normal")
            self._sync_visibility_state()
        self._sync_speed_state()

    def _sync_visibility_state(self) -> None:
        if self.mode_var.get() == "Interactive Reconstruction":
            self.visibility_var.set("Vidljivo")
            self.visibility_combo.configure(state="disabled")
        elif self.renderer_var.get() == "Pure DOCX":
            self.visibility_var.set("U pozadini")
            self.visibility_combo.configure(state="disabled")
        else:
            self.visibility_combo.configure(state="readonly")

    def _sync_metadata_state(self) -> None:
        preserve = self.metadata_var.get() == "Sačuvaj legitimne"
        self.preserve_author_check.configure(state="normal" if preserve else "disabled")
        self.custom_entry.configure(state="normal" if preserve else "disabled")
        if not preserve:
            self.preserve_author_var.set(False)

    def _sync_speed_state(self) -> None:
        custom = self.speed_var.get() == "Custom"
        self.custom_speed_spin.configure(state="normal" if custom else "disabled")

    def _on_speed_changed(self) -> None:
        self._sync_speed_state()
        if self._interactive_control is None:
            return
        mode = {"Slow":"slow", "Fast":"fast", "Custom":"custom", "Maximum":"maximum"}[self.speed_var.get()]
        try:
            update_interactive_control_speed(self._interactive_control, mode, float(self.custom_speed_var.get()))
            self._append_log(f"Brzina promijenjena: {self.speed_var.get()}")
        except Exception as exc:
            self._append_log(f"Brzina nije promijenjena: {exc}")

    def _read_state(self) -> None:
        self.state.reconstruction_mode = "interactive" if self.mode_var.get() == "Interactive Reconstruction" else "instant"
        self.state.renderer = {"Automatski":"auto", "Microsoft Word":"word", "Pure DOCX":"docx"}[self.renderer_var.get()]
        self.state.visibility = {"U pozadini":"background", "Vidljivo":"visible"}[self.visibility_var.get()]
        self.state.fidelity = {"Čista replika":"clean", "Puna vjernost":"full"}[self.fidelity_var.get()]
        self.state.metadata = {"Novi metapodaci":"fresh", "Sačuvaj legitimne":"preserve"}[self.metadata_var.get()]
        self.state.allow_source_overwrite = bool(self.overwrite_var.get()) if self.state.reconstruction_mode == "instant" else False
        self.state.preserve_author_fields = bool(self.preserve_author_var.get())
        self.state.custom_metadata_allowlist = tuple(item.strip() for item in self.custom_properties_var.get().split(",") if item.strip())
        self.state.interactive_speed_mode = {"Slow":"slow", "Fast":"fast", "Custom":"custom", "Maximum":"maximum"}[self.speed_var.get()]
        self.state.interactive_characters_per_second = float(self.custom_speed_var.get())
        self.state.interactive_object_step_delay_ms = int(self.object_delay_var.get())
        self.state.interactive_fidelity = "maximum" if self.interactive_fidelity_var.get() == "Maximum Fidelity" else "standard"
        self.state.checkpoint_after_tables = bool(self.cp_tables_var.get())
        self.state.checkpoint_after_images = bool(self.cp_images_var.get())
        self.state.checkpoint_after_sections = bool(self.cp_sections_var.get())
        self.state.checkpoint_event_interval = int(self.cp_interval_var.get())
        self.state.verify_during_run = bool(self.verify_live_var.get())
        self.state.allow_preserved_objects = bool(self.allow_preserved_var.get())

    def _set_running(self, running: bool, *, interactive: bool | None = None) -> None:
        self._running = running
        self.start_button.configure(state="disabled" if running else "normal")
        self.browse_button.configure(state="disabled" if running else "normal")
        is_interactive = interactive if interactive is not None else self.state.reconstruction_mode == "interactive"
        self.pause_button.configure(state="normal" if running and is_interactive else "disabled")
        self.stop_button.configure(state="normal" if running and is_interactive else "disabled")
        if running:
            self.progress.configure(mode="determinate" if is_interactive else "indeterminate")
            if is_interactive:
                self.progress.stop(); self.progress["value"] = 0
            else:
                self.progress.start(12)
            self.status_var.set("Rekonstrukcija u tijeku…")
        else:
            self.progress.stop()
            self.pause_button.configure(state="disabled")
            self.stop_button.configure(state="disabled")

    def _reset_result_buttons(self) -> None:
        self._last_result = None
        self._last_project_root = None
        self.open_doc_button.configure(state="disabled")
        self.open_report_button.configure(state="disabled")
        self.open_project_button.configure(state="disabled")
        self.saves_var.set("Stvarna spremanja: 0")

    def start_rebuild(self) -> None:
        from tkinter import messagebox
        if self._running:
            return
        self._read_state()
        valid, message = validate_source(self.state)
        if not valid:
            messagebox.showerror("Word Replica", message); return
        if self.state.allow_source_overwrite and not messagebox.askyesno(
            "Zamjena izvornog dokumenta",
            "Uključena je napredna zamjena izvornika. Word Replica će prvo napraviti i verificirati backup. Nastaviti?",
        ):
            return
        self._reset_result_buttons()
        self.preflight_var.set("Preflight: čeka")
        self.semantic_var.set("")
        self.verification_var.set("Verifikacija: čeka")
        source = Path(self.state.source_path)
        options = options_from_desktop_state(self.state)
        self._interactive_control = None
        if options.reconstruction_mode is ReconstructionMode.INTERACTIVE:
            from word_replica.interactive.control import InteractiveRunControl
            self._interactive_control = InteractiveRunControl()
            self._interactive_control.set_speed(options.interactive.speed_mode, options.interactive.characters_per_second)
        self._set_running(True, interactive=options.reconstruction_mode is ReconstructionMode.INTERACTIVE)
        self._append_log(f"Pokrenuta {options.reconstruction_mode.value} rekonstrukcija.")
        threading.Thread(target=self._worker, args=(source, options), daemon=True).start()

    def _worker(self, source: Path, options: RebuildOptions) -> None:
        try:
            result = self.service.rebuild(
                source,
                options,
                interactive_control=self._interactive_control,
                interactive_observer=self._interactive_observer if options.reconstruction_mode is ReconstructionMode.INTERACTIVE else None,
            )
            self._events.put(("result", result))
        except Exception as exc:
            self._events.put(("error", f"{type(exc).__name__}: {exc}"))

    def toggle_pause(self) -> None:
        control = self._interactive_control
        if control is None:
            return
        if control.state is InteractiveRunState.PAUSED:
            control.resume(); self.pause_button.configure(text="Pause"); self.status_var.set("Rekonstrukcija u tijeku…")
        elif control.state is InteractiveRunState.RUNNING:
            control.pause(); self.pause_button.configure(text="Resume"); self.status_var.set("Pauzirano")

    def stop_interactive(self) -> None:
        if self._interactive_control is not None:
            self._interactive_control.stop()
            self.status_var.set("Zaustavljanje na sljedećoj sigurnoj granici…")
            self.stop_button.configure(state="disabled")

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self._events.get_nowait()
                if kind == "result": self._handle_result(payload)
                elif kind == "error": self._handle_error(str(payload))
                elif kind == "preflight": self._handle_preflight(payload)
                elif kind == "progress": self._handle_progress(payload)
                elif kind == "state": self._handle_state(payload)
                elif kind == "warning": self._append_log(f"UPOZORENJE: {payload}")
                elif kind == "checkpoint": self._append_log(f"Checkpoint: {payload}")
                elif kind == "verification": self.verification_var.set(f"Verifikacija: {payload}")
        except queue.Empty:
            pass
        self.root.after(self.POLL_MS, self._poll_events)

    def _handle_preflight(self, report) -> None:
        if report.can_proceed:
            self.preflight_var.set(
                f"Preflight: PASS · {report.counts.get('paragraphs',0)} paragrafa · {report.counts.get('tables',0)} tablica · {report.counts.get('images',0)} slika"
            )
        else:
            self.preflight_var.set("Preflight: BLOCKED")
        for reason in report.blocking_reasons:
            self._append_log(f"PREFLIGHT BLOKADA: {reason}")
        for warning in report.warnings:
            self._append_log(f"PREFLIGHT UPOZORENJE: {warning}")

    def _handle_progress(self, progress) -> None:
        view = progress_view_model(progress)
        self.progress["value"] = view["percent"]
        self.status_var.set(str(view["status"]))
        self.semantic_var.set(str(view["location"]))
        self.verification_var.set(f"Verifikacija: {view['verification']}")
        if progress.state is InteractiveRunState.PAUSED:
            self.pause_button.configure(text="Resume")

    def _handle_state(self, payload) -> None:
        if payload.get("status") == "WORD_STATE_MISMATCH":
            self.status_var.set("PAUSED — Word state mismatch")
            self.pause_button.configure(text="Resume")
            self._append_log(
                f"WORD_STATE_MISMATCH prije eventa {payload.get('event_index')} ({payload.get('source_element_id')})."
            )

    def _handle_result(self, result) -> None:
        self._set_running(False)
        self._interactive_control = None
        self.pause_button.configure(text="Pause")
        self._last_result = result
        self.saves_var.set(f"Stvarna spremanja: {result.save_count}")
        self.status_var.set(f"Status: {result.status.value}")
        self._append_log(f"Završeno — status {result.status.value}; stvarna spremanja: {result.save_count}")
        if result.output_path:
            self._append_log(f"DOCX: {result.output_path}"); self.open_doc_button.configure(state="normal")
        if result.qa_report_path:
            self._append_log(f"QA: {result.qa_report_path}"); self.open_report_button.configure(state="normal")
        for warning in result.warnings:
            self._append_log(f"UPOZORENJE [{warning.code}]: {warning.message}")
        for reason in result.reasons:
            self._append_log(f"RAZLOG: {reason}")
        if result.project_id:
            try:
                project = self.store.get_project(result.project_id)
                self._last_project_root = Path(project["root_path"])
                self.open_project_button.configure(state="normal")
            except Exception:
                self._last_project_root = None
        self._refresh_resume_card()
        from tkinter import messagebox
        if result.status is RunStatus.FAIL:
            messagebox.showerror("Word Replica", "Rekonstrukcija nije prošla QA. Pogledaj detalje u zapisniku.")
        elif result.status is RunStatus.WARN:
            messagebox.showwarning("Word Replica", "Rekonstrukcija je završena uz upozorenja ili je sigurno zaustavljena. Provjeri zapisnik/QA.")
        else:
            messagebox.showinfo("Word Replica", "Rekonstrukcija je uspješno završena.")

    def _handle_error(self, message: str) -> None:
        from tkinter import messagebox
        self._set_running(False)
        self._interactive_control = None
        self.status_var.set("Status: FAIL")
        self._append_log(f"GREŠKA: {message}")
        self._refresh_resume_card()
        messagebox.showerror("Word Replica", message)

    def _refresh_resume_card(self) -> None:
        try:
            items = self.store.list_resumable_projects()
        except Exception:
            items = []
        if not items:
            self._resume_project_id = None
            self.resume_var.set("Nema nedovršene rekonstrukcije.")
            self.resume_button.configure(state="disabled")
            return
        item = items[0]
        self._resume_project_id = str(item["id"])
        name = Path(str(item["source_path"])).name
        index = int(item.get("last_completed_event_index", -1))
        stamp = str(item.get("checkpoint_timestamp_utc", ""))
        self.resume_var.set(f"{name} · zadnji event {index} · {stamp}")
        self.resume_button.configure(state="normal" if not self._running else "disabled")

    def resume_latest(self) -> None:
        if self._running or not self._resume_project_id:
            return
        from word_replica.interactive.control import InteractiveRunControl
        self._interactive_control = InteractiveRunControl(); self._interactive_control.start()
        self.state.reconstruction_mode = "interactive"
        self.mode_var.set("Interactive Reconstruction")
        self._sync_mode_state()
        self._set_running(True, interactive=True)
        project_id = self._resume_project_id
        self._append_log(f"Nastavak Interactive projekta: {project_id}")
        threading.Thread(target=self._resume_worker, args=(project_id,), daemon=True).start()

    def _resume_worker(self, project_id: str) -> None:
        try:
            result = self.service.resume_interactive(
                project_id,
                interactive_control=self._interactive_control,
                interactive_observer=self._interactive_observer,
            )
            self._events.put(("result", result))
        except Exception as exc:
            self._events.put(("error", f"{type(exc).__name__}: {exc}"))

    def open_output(self) -> None:
        if self._last_result and self._last_result.output_path:
            _open_path(Path(self._last_result.output_path))

    def open_report(self) -> None:
        if self._last_result and self._last_result.qa_report_path:
            _open_path(Path(self._last_result.qa_report_path))

    def open_project(self) -> None:
        if self._last_project_root:
            _open_path(self._last_project_root)


def main() -> int:
    import tkinter as tk
    root = tk.Tk()
    WordReplicaDesktop(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
