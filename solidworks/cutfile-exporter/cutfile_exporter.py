from __future__ import annotations

import argparse
import csv
from dataclasses import asdict, dataclass
import json
import os
from pathlib import Path
import tempfile
import threading
import traceback
from typing import Callable, Iterable

from cutfile_core import (
    CutfileValidationError,
    LINE_MARKING_LAYER,
    TEXT_MARKING_LAYER,
    add_marking_paths,
    assign_cut_layers,
    infer_model_to_dxf_scale,
    load_dxf,
    save_dxf,
)
from solidworks_adapter import (
    SolidWorksExportError,
    SolidWorksSession,
    project_marking_paths,
)


@dataclass(frozen=True)
class ExportRecord:
    source: str
    output: str
    status: str
    outside_loops: int = 0
    inside_loops: int = 0
    cut_entities: int = 0
    marking_paths: int = 0
    units: str = ""
    message: str = ""


def export_part(
    session: SolidWorksSession,
    source: Path,
    output: Path,
    *,
    sketch_name: str,
    overwrite: bool,
) -> ExportRecord:
    if output.exists() and not overwrite:
        return ExportRecord(
            str(source), str(output), "SKIPPED", message="Output exists; overwrite is off."
        )

    opened = None
    try:
        opened = session.open_part(source)
        face_info = session.find_export_face(opened.model)
        with tempfile.TemporaryDirectory(prefix="solidworks-cutfile-") as temp_dir:
            temporary_dxf = Path(temp_dir) / f"{source.stem}.dxf"
            session.export_face_dxf(opened, face_info, temporary_dxf)
            document = load_dxf(temporary_dxf)
            layering = assign_cut_layers(
                document,
                expected_outer_loops=face_info.outer_loops,
                expected_inner_loops=face_info.inner_loops,
            )
            scale, inferred_units = infer_model_to_dxf_scale(
                document, face_info.projected_points_m
            )
            marking_model_paths = session.marking_paths_model_space(
                opened.model, sketch_name
            )
            line_marking_paths = project_marking_paths(
                marking_model_paths.line_paths, face_info.frame, scale
            )
            text_marking_paths = project_marking_paths(
                marking_model_paths.text_paths, face_info.frame, scale
            )
            line_marking = add_marking_paths(
                document,
                line_marking_paths,
                layer=LINE_MARKING_LAYER,
            )
            text_marking = add_marking_paths(
                document,
                text_marking_paths,
                layer=TEXT_MARKING_LAYER,
            )
            marking_path_count = (
                line_marking.paths_added + text_marking.paths_added
            )
            save_dxf(document, output)

        units = layering.drawing_units
        if units == "unitless":
            units = f"{inferred_units} (inferred)"
        return ExportRecord(
            source=str(source),
            output=str(output),
            status="OK",
            outside_loops=layering.outside_loops,
            inside_loops=layering.inside_loops,
            cut_entities=layering.cut_entities,
            marking_paths=marking_path_count,
            units=units,
            message=(
                f"Created layered DXF; marking sketch '{sketch_name}' "
                f"produced {line_marking.paths_added} line-marking path(s) and "
                f"{text_marking.paths_added} text path(s)."
            ),
        )
    except (SolidWorksExportError, CutfileValidationError) as exc:
        return ExportRecord(str(source), str(output), "FAILED", message=str(exc))
    except Exception as exc:
        return ExportRecord(
            str(source),
            str(output),
            "FAILED",
            message=f"Unexpected error: {exc}\n{traceback.format_exc(limit=3)}",
        )
    finally:
        if opened is not None:
            session.close_part(opened)


def export_many(
    sources: Iterable[Path],
    output_dir: Path,
    *,
    sketch_name: str = "CUTFILE MARKING",
    overwrite: bool = False,
    progress: Callable[[ExportRecord], None] | None = None,
) -> list[ExportRecord]:
    source_list = [Path(source) for source in sources]
    _reject_duplicate_output_names(source_list)

    com_runtime = _initialize_com_thread()
    session = None
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        session = SolidWorksSession.connect(visible=True)
        records: list[ExportRecord] = []
        for source in source_list:
            record = export_part(
                session,
                source,
                output_dir / f"{source.stem}.dxf",
                sketch_name=sketch_name,
                overwrite=overwrite,
            )
            records.append(record)
            if progress:
                progress(record)
        write_report(output_dir / "cutfile_export_report.csv", records)
        return records
    finally:
        session = None
        if com_runtime is not None:
            com_runtime.CoUninitialize()


def write_report(path: Path, records: Iterable[ExportRecord]) -> None:
    rows = [asdict(record) for record in records]
    fieldnames = list(ExportRecord.__dataclass_fields__)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def discover_parts(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        return (
            [input_path]
            if input_path.suffix.casefold() == ".sldprt"
            and not input_path.name.startswith("~$")
            else []
        )
    candidates = input_path.rglob("*") if recursive else input_path.iterdir()
    return sorted(
        (
            path
            for path in candidates
            if path.is_file()
            and path.suffix.casefold() == ".sldprt"
            and not path.name.startswith("~$")
        ),
        key=lambda path: str(path).casefold(),
    )


def _reject_duplicate_output_names(sources: Iterable[Path]) -> None:
    grouped: dict[str, list[Path]] = {}
    for source in sources:
        grouped.setdefault(f"{source.stem}.dxf".casefold(), []).append(source)
    duplicates = [paths for paths in grouped.values() if len(paths) > 1]
    if not duplicates:
        return
    examples = "; ".join(
        ", ".join(str(path) for path in paths[:3]) for paths in duplicates[:3]
    )
    raise SolidWorksExportError(
        "Multiple input parts would create the same DXF filename. Rename the parts "
        f"or export their folders separately. Conflicts: {examples}"
    )


def _initialize_com_thread():
    """Initialize COM on whichever thread is performing SolidWorks automation."""

    if os.name != "nt":
        return None
    try:
        import pythoncom
    except ImportError as exc:
        raise SolidWorksExportError(
            "pywin32 is not installed. Run: py -m pip install -r requirements.txt"
        ) from exc
    pythoncom.CoInitialize()
    return pythoncom


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create layered DXF cut files from flat SolidWorks parts."
    )
    parser.add_argument("--input", type=Path, help="SLDPRT file or folder")
    parser.add_argument("--output", type=Path, help="Output DXF folder")
    parser.add_argument("--sketch-name", default="CUTFILE MARKING")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--json", action="store_true", help="Print JSON result")
    return parser


def run_cli(args: argparse.Namespace) -> int:
    if args.input is None or args.output is None:
        raise SystemExit("--input and --output are required in command-line mode")
    parts = discover_parts(args.input.resolve(), args.recursive)
    if not parts:
        raise SystemExit("No .SLDPRT files were found")
    records = export_many(
        parts,
        args.output.resolve(),
        sketch_name=args.sketch_name,
        overwrite=args.overwrite,
        progress=lambda record: print(
            f"[{record.status}] {Path(record.source).name}: {record.message}"
        ),
    )
    if args.json:
        print(json.dumps([asdict(record) for record in records], indent=2))
    return 1 if any(record.status == "FAILED" for record in records) else 0


def run_gui() -> None:
    import queue
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("SolidWorks Layered Cut-File Exporter")
    root.geometry("820x560")
    root.minsize(720, 480)
    steel_blue = "#57a0d3"
    dark_bg = "#101820"
    dark_surface = "#18232d"
    dark_input = "#111a22"
    text_primary = "#eef5f9"
    root.configure(background=dark_bg)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("TFrame", background=dark_bg)
    style.configure("TLabel", background=dark_bg, foreground=text_primary)
    style.configure(
        "TEntry",
        fieldbackground=dark_input,
        foreground=text_primary,
        insertcolor=text_primary,
        bordercolor="#40515f",
    )
    style.configure(
        "TButton",
        background=dark_surface,
        foreground=text_primary,
        bordercolor="#40515f",
        padding=(10, 6),
    )
    style.map("TButton", background=[("active", "#2b3c49")])
    style.configure(
        "Accent.TButton",
        background=steel_blue,
        foreground="#ffffff",
        bordercolor=steel_blue,
        padding=(12, 7),
    )
    style.map("Accent.TButton", background=[("active", "#6eb1df")])
    style.configure("TCheckbutton", background=dark_bg, foreground=text_primary)
    style.map("TCheckbutton", background=[("active", dark_bg)])

    input_var = tk.StringVar()
    output_var = tk.StringVar()
    sketch_var = tk.StringVar(value="CUTFILE MARKING")
    recursive_var = tk.BooleanVar(value=False)
    overwrite_var = tk.BooleanVar(value=False)
    status_var = tk.StringVar(value="Select an input folder and output folder.")

    frame = ttk.Frame(root, padding=14)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(6, weight=1)

    ttk.Label(frame, text="SolidWorks Layered Cut-File Exporter", font=("Segoe UI", 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 12)
    )
    ttk.Label(frame, text="Input SLDPRT folder").grid(row=1, column=0, sticky="w", pady=4)
    ttk.Entry(frame, textvariable=input_var).grid(row=1, column=1, sticky="ew", padx=8)
    ttk.Button(
        frame,
        text="Browse",
        command=lambda: input_var.set(filedialog.askdirectory() or input_var.get()),
    ).grid(row=1, column=2)

    ttk.Label(frame, text="Output DXF folder").grid(row=2, column=0, sticky="w", pady=4)
    ttk.Entry(frame, textvariable=output_var).grid(row=2, column=1, sticky="ew", padx=8)
    ttk.Button(
        frame,
        text="Browse",
        command=lambda: output_var.set(filedialog.askdirectory() or output_var.get()),
    ).grid(row=2, column=2)

    ttk.Label(frame, text="Marking sketch name").grid(row=3, column=0, sticky="w", pady=4)
    ttk.Entry(frame, textvariable=sketch_var).grid(row=3, column=1, sticky="ew", padx=8)
    ttk.Label(
        frame,
        text="Lines → PIN STAMP LINE MARKING; text → PIN STAMP TEXT",
    ).grid(
        row=3, column=2, sticky="w"
    )

    options = ttk.Frame(frame)
    options.grid(row=4, column=0, columnspan=3, sticky="w", pady=8)
    ttk.Checkbutton(options, text="Include subfolders", variable=recursive_var).pack(
        side="left"
    )
    ttk.Checkbutton(options, text="Overwrite existing DXFs", variable=overwrite_var).pack(
        side="left", padx=18
    )

    log = tk.Text(
        frame,
        height=18,
        wrap="word",
        state="disabled",
        font=("Consolas", 9),
        background=dark_input,
        foreground=text_primary,
        insertbackground=text_primary,
        selectbackground="#315f7d",
        relief="flat",
    )
    log.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(8, 8))
    ttk.Label(frame, textvariable=status_var).grid(row=7, column=0, columnspan=2, sticky="w")

    def append(message: str) -> None:
        log.configure(state="normal")
        log.insert("end", message.rstrip() + "\n")
        log.see("end")
        log.configure(state="disabled")

    def set_running(running: bool) -> None:
        run_button.configure(state="disabled" if running else "normal")

    def start_export() -> None:
        input_text = input_var.get().strip()
        output_text = output_var.get().strip()
        sketch_name = sketch_var.get().strip()
        recursive = recursive_var.get()
        overwrite = overwrite_var.get()
        if not input_text:
            messagebox.showerror("Missing input", "Select an existing SLDPRT folder.")
            return
        input_path = Path(input_text)
        output_path = Path(output_text)
        if not input_path.exists():
            messagebox.showerror("Missing input", "Select an existing SLDPRT folder.")
            return
        if not output_text:
            messagebox.showerror("Missing output", "Select an output folder.")
            return
        parts = discover_parts(input_path, recursive)
        if not parts:
            messagebox.showerror("No parts", "No .SLDPRT files were found.")
            return
        if not sketch_name:
            messagebox.showerror("Sketch name", "Enter the SolidWorks marking-sketch name.")
            return

        set_running(True)
        status_var.set(f"Running {len(parts)} part(s)...")
        append(f"Starting {len(parts)} part(s). SolidWorks may open or become visible.")
        events: queue.Queue[tuple[str, object]] = queue.Queue()

        def worker() -> None:
            try:
                records = export_many(
                    parts,
                    output_path,
                    sketch_name=sketch_name,
                    overwrite=overwrite,
                    progress=lambda record: events.put(("record", record)),
                )
                failures = sum(record.status == "FAILED" for record in records)
                skipped = sum(record.status == "SKIPPED" for record in records)
                successes = sum(record.status == "OK" for record in records)
                events.put(
                    (
                        "finish",
                        (successes, failures, skipped, output_path),
                    )
                )
            except Exception as exc:
                events.put(("fail", str(exc)))

        def poll_worker() -> None:
            done = False
            try:
                while True:
                    kind, payload = events.get_nowait()
                    if kind == "record":
                        record = payload
                        append(
                            f"[{record.status}] {Path(record.source).name}: "
                            f"{record.message}"
                        )
                    elif kind == "finish":
                        finish(*payload)
                        done = True
                    elif kind == "fail":
                        fail(str(payload))
                        done = True
            except queue.Empty:
                pass
            if not done:
                root.after(50, poll_worker)

        threading.Thread(target=worker, daemon=True).start()
        root.after(50, poll_worker)

    def finish(successes: int, failures: int, skipped: int, output_path: Path) -> None:
        set_running(False)
        status_var.set(
            f"Complete: {successes} created, {failures} failed, {skipped} skipped."
        )
        messagebox.showinfo(
            "Export complete",
            f"Created: {successes}\nFailed: {failures}\nSkipped: {skipped}\n\n"
            f"Review cutfile_export_report.csv in:\n{output_path}",
        )

    def fail(message: str) -> None:
        set_running(False)
        status_var.set("Export stopped.")
        append(f"[FATAL] {message}")
        messagebox.showerror("Export stopped", message)

    run_button = ttk.Button(
        frame,
        text="Create Layered DXFs",
        command=start_export,
        style="Accent.TButton",
    )
    run_button.grid(row=7, column=2, sticky="e")
    root.mainloop()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.input is None and args.output is None:
        run_gui()
        return 0
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
