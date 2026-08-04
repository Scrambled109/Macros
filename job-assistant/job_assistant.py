"""Windows-oriented dashboard and guided stages for engineering jobs."""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from job_core import (
    MANIFEST_NAME,
    STAGES,
    STAGE_GUIDANCE,
    JobError,
    PromotionItem,
    change_revision,
    command_bom,
    command_comparison,
    command_dxf,
    complete_stage,
    copy_source,
    dashboard_warnings,
    discover_drawings,
    existing_working_directory,
    export_parts_list_csv,
    load_manifest,
    load_settings,
    mark_needs_review,
    inferred_plate_thickness,
    parse_comparison_summary,
    plate_macro_instructions,
    plate_run_folders,
    plan_promotions,
    prepare_dxf_workspace,
    promote_files,
    recommended_next_action,
    record_artifact,
    record_event,
    reopen_stage,
    save_manifest,
    save_settings,
    safe_name,
    set_optional_path,
    setup_job,
    stage_checks,
    start_stage,
    suggest_job_number,
)

HERE = Path(__file__).resolve().parent
DEFAULT_REPO = HERE.parent


def open_path(path: Path) -> None:
    if sys.platform != "win32":
        raise JobError("Open Folder is available in the packaged Windows application.")
    os.startfile(str(path))


class JobAssistant(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Engineering Job Assistant — Beta")
        self.geometry("1280x860")
        self.minsize(1050, 700)
        self.settings = load_settings(repo_root=DEFAULT_REPO)
        self.manifest: dict | None = None
        self.manifest_path: Path | None = None
        self.active_stage: str | None = None
        self.candidates = []
        self.running_processes: dict[int, dict] = {}
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.close_application)

    @property
    def repo(self) -> Path:
        return Path(self.settings.get("macros_repo") or DEFAULT_REPO)

    def bundled_tool(self, name: str) -> Path | None:
        """Return a packaged companion executable, or use source tools."""
        if not getattr(sys, "frozen", False):
            return None
        path = Path(sys.executable).parent / name
        if not path.is_file():
            raise JobError(
                f"The packaged companion tool is missing: {path}. "
                "Reinstall or rebuild the complete Job Assistant distribution."
            )
        return path

    def _build(self) -> None:
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill="x")
        for text, command in (
            ("Set Up / Attach Job", self.setup_job),
            ("Open Job", self.open_job),
            ("Settings", self.edit_settings),
            ("Change Revision", self.change_job_revision),
            ("Refresh", self.refresh),
        ):
            ttk.Button(toolbar, text=text, command=command).pack(side="left", padx=3)
        self.running_summary = ttk.Label(
            toolbar,
            text="No external processes running",
            foreground="#176b32",
        )
        self.running_summary.pack(side="right", padx=8)
        self.heading = ttk.Label(
            self,
            text="Set up a new job or attach to an existing Engineering Process folder.",
            font=("Segoe UI", 14, "bold"),
            padding=(12, 3),
        )
        self.heading.pack(fill="x")
        self.next_action = ttk.Label(
            self,
            text="The assistant works on controlled copies and keeps generated files in staging.",
            padding=(12, 3),
        )
        self.next_action.pack(fill="x")
        self.warning_summary = ttk.Label(
            self,
            text="No outstanding warnings.",
            padding=(12, 5),
            foreground="#7a2e00",
            wraplength=1220,
        )
        self.warning_summary.pack(fill="x")
        self.paths_label = ttk.Label(self, text="", padding=(12, 3), wraplength=1140)
        self.paths_label.pack(fill="x")

        pane = ttk.Panedwindow(self, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=12, pady=8)
        left = ttk.Frame(pane)
        right = ttk.Frame(pane, padding=(12, 0))
        pane.add(left, weight=3)
        pane.add(right, weight=4)
        self.tree = ttk.Treeview(
            left, columns=("stage", "status", "files"), show="headings", height=18
        )
        for name, width in (("stage", 285), ("status", 120), ("files", 55)):
            self.tree.heading(name, text=name.title())
            self.tree.column(name, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._stage_selected)
        self.tree.tag_configure("not_started", foreground="#666666")
        self.tree.tag_configure("ready", foreground="#1f5f99")
        self.tree.tag_configure("in_progress", foreground="#7a4b00")
        self.tree.tag_configure("needs_review", foreground="#8a3f00")
        self.tree.tag_configure("complete", foreground="#176b32")
        self.tree.tag_configure("warning", foreground="#a31621")

        recent_frame = ttk.LabelFrame(left, text="Recently recorded files", padding=5)
        recent_frame.pack(fill="x", pady=(8, 0))
        self.recent_tree = ttk.Treeview(
            recent_frame,
            columns=("name", "revision"),
            show="headings",
            height=5,
        )
        self.recent_tree.heading("name", text="File")
        self.recent_tree.heading("revision", text="Rev")
        self.recent_tree.column("name", width=330, anchor="w")
        self.recent_tree.column("revision", width=50, anchor="center")
        self.recent_tree.pack(fill="x")
        self.recent_tree.bind("<Double-1>", lambda _event: self.open_recent_file())

        ttk.Label(right, text="Stage guide", font=("Segoe UI", 12, "bold")).pack(
            anchor="w"
        )
        self.guide = tk.Text(
            right, wrap="word", height=22, state="disabled", padx=8, pady=8
        )
        self.guide.pack(fill="both", expand=True, pady=(5, 8))
        actions = ttk.Frame(right)
        actions.pack(fill="x")
        for text, command in (
            ("Check This Step", self.run_checks),
            ("Start This Step", self.run_stage),
            ("Open This Step's Folder", self.open_stage_folder),
            ("Record File", self.record_file),
            ("Complete After Review", self.finish_stage),
            ("Reopen", self.reopen),
        ):
            ttk.Button(actions, text=text, command=command).pack(
                side="left", padx=2, pady=2
            )

        bottom = ttk.Frame(self, padding=(12, 0, 12, 10))
        bottom.pack(fill="x")
        ttk.Button(
            bottom, text="Set Optional Folder", command=self.set_optional_folder
        ).pack(side="left", padx=3)
        ttk.Button(
            bottom, text="Copy Approved Files to Production", command=self.promote
        ).pack(side="left", padx=3)
        ttk.Button(
            bottom, text="Open Assistant Workspace", command=self.open_workspace
        ).pack(side="left", padx=3)
        ttk.Button(bottom, text="Open Staging", command=self.open_staging).pack(
            side="left", padx=3
        )
        ttk.Button(bottom, text="Open Logs", command=self.open_logs).pack(
            side="left", padx=3
        )
        ttk.Button(
            bottom, text="Open Comparison Report", command=self.open_comparison_report
        ).pack(side="left", padx=3)
        self.progress = ttk.Progressbar(bottom, length=220)
        self.progress.pack(side="right", padx=5)
        self.progress_text = ttk.Label(bottom, text="0 of 0 complete")
        self.progress_text.pack(side="right", padx=5)

    def handle(self, operation) -> None:
        try:
            operation()
        except (JobError, OSError, subprocess.SubprocessError) as exc:
            save_error = ""
            if self.manifest is not None and self.manifest_path is not None:
                try:
                    save_manifest(self.manifest, self.manifest_path)
                except OSError as manifest_exc:
                    save_error = (
                        "\n\nThe updated audit history could not be saved: "
                        f"{manifest_exc}"
                    )
            messagebox.showerror(
                "Engineering Job Assistant", f"{exc}{save_error}", parent=self
            )

    def update_running_summary(self) -> None:
        count = len(self.running_processes)
        if not count:
            text, color = "No external processes running", "#176b32"
        else:
            jobs = sorted(
                {item["job_number"] for item in self.running_processes.values()}
            )
            text = f"Running: {count} process(es) for job(s) {', '.join(jobs)}"
            color = "#7a4b00"
        self.running_summary.configure(text=text, foreground=color)

    def close_application(self) -> None:
        if self.running_processes and not messagebox.askyesno(
            "Processes are still running",
            "External automation is still running. Closing the assistant will "
            "stop status monitoring and completion notifications, but will not "
            "stop AutoCAD.\n\nClose the assistant anyway?",
            icon="warning",
            parent=self,
        ):
            return
        self.destroy()

    def require_job(self) -> None:
        if not self.manifest or not self.manifest_path:
            raise JobError("Set up or open a job first.")

    def selected_stage(self) -> str:
        if not self.active_stage:
            raise JobError("Select a workflow stage first.")
        return self.active_stage

    def _stage_selected(self, _event=None) -> None:
        selected = self.tree.selection()
        if selected:
            self.active_stage = selected[0]
            self.show_stage()

    def setup_job(self) -> None:
        def operation() -> None:
            initial = self.settings.get("default_jobs_parent") or None
            root_value = filedialog.askdirectory(
                title="Select Engineering Process / root folder",
                initialdir=initial,
                parent=self,
            )
            if not root_value:
                return
            root = Path(root_value)
            model = filedialog.askdirectory(
                title="Select the 3D Model folder", initialdir=root, parent=self
            )
            if not model:
                return
            cut = filedialog.askdirectory(
                title="Select the Cut Files folder", initialdir=root, parent=self
            )
            if not cut:
                return
            number = simpledialog.askstring(
                "Job setup",
                "Confirm or enter the authoritative job number:",
                initialvalue=suggest_job_number(root),
                parent=self,
            )
            if not number:
                return
            name = (
                simpledialog.askstring(
                    "Job setup",
                    "Job name (optional):",
                    initialvalue=root.parent.name,
                    parent=self,
                )
                or number
            )
            revision = simpledialog.askstring(
                "Job setup", "Initial revision:", initialvalue="A", parent=self
            )
            if not revision:
                return
            existing = root / "_JOB_ASSISTANT" / MANIFEST_NAME
            if existing.exists():
                messagebox.showinfo(
                    "Existing assistant job",
                    "This folder already has an assistant manifest. It will be "
                    "opened; existing audit history will not be replaced.",
                    parent=self,
                )
                self._load(existing)
                return
            path = setup_job(root, Path(model), Path(cut), number, name, revision)
            self.settings["default_jobs_parent"] = str(root.parent)
            save_settings(self.settings)
            self._load(path)

        self.handle(operation)

    def open_job(self) -> None:
        selected = filedialog.askopenfilename(
            title="Open job manifest",
            filetypes=[("Job manifest", "*.json")],
            parent=self,
        )
        if selected:
            self.handle(lambda: self._load(Path(selected)))

    def _load(self, path: Path) -> None:
        self.manifest, self.manifest_path = load_manifest(path), path
        save_manifest(self.manifest, path)  # persist any safe migration
        self.refresh()

    def refresh(self) -> None:
        if not self.manifest:
            return
        self.manifest = load_manifest(self.manifest_path)
        job = self.manifest["job"]
        self.heading.configure(
            text=f"{job['number']} — {job['name']}  |  Revision {job['revision']}"
        )
        self.next_action.configure(text=recommended_next_action(self.manifest))
        warnings = dashboard_warnings(self.manifest)
        self.warning_summary.configure(
            text=(
                "ACTION / WARNING: " + "  •  ".join(warnings)
                if warnings
                else "No outstanding dashboard warnings."
            )
        )
        p = self.manifest["paths"]
        self.paths_label.configure(
            text=f"Engineering Process: {p['engineering_root']}\n3D Model: {p['model_3d']}    |    Cut Files: {p['cut_files']}\nOptional — Part Checking: {p['part_checking'] or 'not selected'}    |    Nesting: {p['nesting'] or 'not selected'}"
        )
        self.tree.delete(*self.tree.get_children())
        complete = 0
        for key, label in STAGES:
            item = self.manifest["stages"][key]
            complete += item["status"] == "complete"
            self.tree.insert(
                "",
                "end",
                iid=key,
                values=(
                    label,
                    item["status"].replace("_", " ").title(),
                    len(item["artifacts"]),
                ),
                tags=(item["status"],),
            )
        stage_keys = [key for key, _label in STAGES]
        if self.active_stage not in stage_keys:
            self.active_stage = next(
                (
                    key
                    for key in stage_keys
                    if self.manifest["stages"][key]["status"] != "complete"
                ),
                stage_keys[0],
            )
        self.tree.selection_set(self.active_stage)
        self.tree.focus(self.active_stage)
        self.tree.see(self.active_stage)
        self.progress.configure(maximum=len(STAGES), value=complete)
        self.progress_text.configure(text=f"{complete} of {len(STAGES)} complete")
        self.recent_tree.delete(*self.recent_tree.get_children())
        for index, artifact in enumerate(self.manifest.get("recent_files", [])[:5]):
            self.recent_tree.insert(
                "",
                "end",
                iid=f"recent-{index}",
                values=(artifact.get("name", ""), artifact.get("revision", "")),
            )
        self.show_stage()

    def show_stage(self) -> None:
        if not self.manifest or not self.active_stage:
            return
        stage = self.selected_stage()
        guide = STAGE_GUIDANCE[stage]
        item = self.manifest["stages"][stage]
        checks = stage_checks(self.manifest, stage)
        text = (
            f"STATUS: {item['status'].replace('_', ' ').upper()}\n\nUse Check This Step to review readiness, Start This Step to begin, and Open This Step's Folder to inspect its files.\n\n1. WHAT IS NEEDED\n{guide['need']}\n\n2. WHAT TO SELECT\n{guide['action']}\n\n3. WHAT WILL CHANGE\n{guide['changes']}\n\n4. TOOL / MACRO\n{guide['tool']}\n\n5. REVIEW AFTERWARD\n{guide['review']}\n\n6. OUTPUTS AND LOGS\nStaging: {self.manifest['workspace']['staging']}\nLogs: {self.manifest['workspace']['logs']}\n\n7. WARNINGS AND OVERRIDES\nOrdinary sequence warnings can be continued after confirmation and are logged with your Windows username. Missing files or executables cannot be launched.\n\nCURRENT CHECKS\n"
            + "\n".join(f"• {c.level.upper()}: {c.message}" for c in checks)
        )
        self.guide.configure(state="normal")
        self.guide.delete("1.0", "end")
        self.guide.insert("end", text)
        self.guide.configure(state="disabled")

    def run_checks(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            checks = stage_checks(self.manifest, stage)
            results = checks or [
                type("ReadinessPass", (), {
                    "level": "pass",
                    "message": "This step is ready to start.",
                    "correction": "No resolution is required.",
                })()
            ]
            messagebox.showinfo(
                "Readiness results",
                "\n\n".join(
                    f"{check.level.upper()}\nMessage: {check.message}\n"
                    f"Suggested resolution: {check.correction or 'No action required.'}"
                    for check in results
                ),
                parent=self,
            )

        self.handle(operation)

    def _accept_warnings(self, stage: str) -> bool:
        checks = stage_checks(self.manifest, stage)
        blocks = [c for c in checks if c.level == "block"]
        if blocks:
            raise JobError("Cannot continue:\n" + "\n".join(c.message for c in blocks))
        warnings = [c for c in checks if c.level == "warning"]
        if warnings and not messagebox.askyesno(
            "Continue Anyway?",
            "Warnings:\n\n"
            + "\n".join(c.message for c in warnings)
            + "\n\nContinuing may produce incomplete or mismatched results. Continue and record this override?",
            parent=self,
        ):
            return False
        start_stage(self.manifest, stage, warnings)
        return True

    def run_stage(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            if stage == "bom":
                self.run_bom()
            elif stage == "dxf":
                self.run_dxf()
            elif stage == "comparison":
                self.run_comparison()
            elif stage in {"plate_model", "autobom"}:
                self.launch_solidworks_macro(stage)
            else:
                if not self._accept_warnings(stage):
                    return
                open_path(
                    Path(
                        self.manifest["paths"]["model_3d"]
                        if stage in {"manual_model", "assembly"}
                        else self.manifest["paths"].get("nesting")
                        or self.manifest["root"]
                    )
                )
                mark_needs_review(
                    self.manifest,
                    stage,
                    "Manual workflow opened. Complete the listed engineering review.",
                )
            save_manifest(self.manifest, self.manifest_path)
            self.refresh()
            self.show_stage()

        self.handle(operation)

    def run_bom(self) -> None:
        source_value = filedialog.askopenfilename(
            title="Select the received BOP/BOM",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm")],
            parent=self,
        )
        if not source_value:
            return
        template_value = self.settings.get("parts_list_template", "")
        if not template_value or not Path(template_value).is_file():
            template_value = filedialog.askopenfilename(
                title="Select the standard Parts List template (remembered for next time)",
                filetypes=[("Excel workbook", "*.xlsx")],
                parent=self,
            )
            if not template_value:
                return
            self.settings["parts_list_template"] = template_value
            save_settings(self.settings)
        default = (
            Path(self.manifest["paths"]["engineering_root"])
            / f"{self.manifest['job']['number']}_PARTS_LIST.xlsx"
        )
        output = filedialog.asksaveasfilename(
            title="Confirm Parts List destination",
            initialdir=default.parent,
            initialfile=default.name,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            parent=self,
        )
        if not output or not self._accept_warnings("bom"):
            return
        source_copy = copy_source(
            Path(source_value),
            Path(self.manifest["workspace"]["source_copies"]) / "BOP-BOM",
        )
        command = command_bom(
            sys.executable,
            self.repo,
            source_copy,
            Path(output),
            Path(template_value),
            self.bundled_tool("Engineering BOM Converter.exe"),
        )
        # A remembered repository can be unavailable (especially a disconnected
        # network drive).  Passing that stale path as cwd raises WinError 267
        # before the packaged converter can even start.  The command already
        # contains absolute executable/script paths, so it is safe to inherit the
        # assistant's working directory when the preferred repository is absent.
        subprocess.Popen(command, cwd=existing_working_directory(self.repo))
        record_event(
            self.manifest,
            "external_process_started",
            stage="bom",
            command=command,
        )
        mark_needs_review(
            self.manifest,
            "bom",
            "Converter setup opened. Review mappings, convert, then positively review the Parts List.",
            [source_copy, Path(output)],
        )
        # The converter owns the output process, so it may not exist yet and
        # cannot be recorded as an artifact above.  Retain its intended path so
        # completion can create the orchestrator-specific, plate-only CSV.
        self.manifest["stages"]["bom"]["parts_list_workbook"] = str(Path(output))

    def run_dxf(self) -> None:
        incoming = filedialog.askdirectory(
            title="Select the folder containing received DXFs and shape sketches",
            parent=self,
        )
        if not incoming:
            return
        self.candidates = discover_drawings(Path(incoming))
        dxfs = sum(c.path.suffix.lower() == ".dxf" for c in self.candidates)
        dwgs = sum(c.path.suffix.lower() == ".dwg" for c in self.candidates)
        if not self.candidates:
            raise JobError("No DXF or DWG files were found directly in that folder.")
        review = messagebox.askyesnocancel(
            "Review proposed plate inputs",
            f"Found {dxfs} DXF files and {dwgs} DWG files.\n\nThe assistant plans to copy the DXFs into the plate automation workspace and leave the DWGs untouched.\n\nYes = confirm/prepare\nNo = review individual files\nCancel = choose nothing",
            parent=self,
        )
        if review is None:
            return
        selected = (
            [c.path for c in self.candidates if c.selected]
            if review
            else self.review_drawings()
        )
        if not selected or not self._accept_warnings("dxf"):
            return
        parts_list_csv = (
            Path(self.manifest["workspace"]["source_copies"]) / "Parts List.csv"
        )
        if not parts_list_csv.is_file():
            selected_csv = filedialog.askopenfilename(
                title="Select the Parts List CSV required by the DXF orchestrator",
                filetypes=[("CSV Parts List", "*.csv")],
                parent=self,
            )
            if not selected_csv:
                raise JobError(
                    "The DXF orchestrator requires Parts List.csv. Run the Parts "
                    "List stage first or select an existing CSV."
                )
            parts_list_csv = Path(selected_csv)
        run = prepare_dxf_workspace(self.manifest, selected, parts_list_csv)
        if not messagebox.askyesno(
            "Workspace prepared",
            f"Copied {len(selected)} file(s) to:\n{run / '001'}\n\nOriginals were not changed. Launch the orchestrator now?",
            parent=self,
        ):
            mark_needs_review(
                self.manifest,
                "dxf",
                "Working copies prepared; orchestrator not yet launched.",
            )
            return
        log = Path(self.manifest["workspace"]["logs"]) / f"dxf-{run.name}.log"
        command = command_dxf(
            sys.executable,
            self.repo,
            run,
            run / "Parts List.csv",
            self.settings.get("autocad_console") or None,
            self.settings.get("autocad_executable") or None,
            self.bundled_tool("Engineering DXF Orchestrator.exe"),
        )
        record_event(
            self.manifest,
            "external_process_requested",
            stage="dxf",
            command=command,
            workspace=str(run),
            log=str(log),
        )
        save_manifest(self.manifest, self.manifest_path)
        handle = log.open("w", encoding="utf-8")
        handle.write(f"Windows command: {subprocess.list2cmdline(command)}\n")
        handle.write(f"Arguments: {command!r}\nWorkspace: {run}\nAssistant log: {log}\nStage: dxf\n\n")
        handle.flush()
        try:
            process = subprocess.Popen(
                command,
                cwd=run,
                stdout=handle,
                stderr=subprocess.STDOUT,
            )
        except Exception:
            handle.close()
            raise
        record_event(
            self.manifest,
            "external_process_launched",
            stage="dxf",
            pid=process.pid,
            workspace=str(run),
            log=str(log),
        )
        self.running_processes[process.pid] = {
            "stage": "dxf",
            "job_number": self.manifest["job"]["number"],
            "manifest_path": str(self.manifest_path),
            "log": str(log),
        }
        self.update_running_summary()
        save_manifest(self.manifest, self.manifest_path)
        # Keep the launched job's state with the process. The operator can open a
        # different job while AutoCAD is still running; a later callback must not
        # write that process result into whichever job happens to be displayed.
        self._poll_dxf_process(
            process,
            handle,
            log,
            run,
            self.manifest,
            self.manifest_path,
        )

    def _poll_dxf_process(
        self,
        process,
        handle,
        log: Path,
        run: Path,
        process_manifest: dict,
        process_manifest_path: Path,
    ) -> None:
        exit_code = process.poll()
        if exit_code is None:
            self.after(
                500,
                self._poll_dxf_process,
                process,
                handle,
                log,
                run,
                process_manifest,
                process_manifest_path,
            )
            return
        handle.write(f"\nExit code: {exit_code}\n")
        handle.close()
        running = self.__dict__.get("running_processes")
        if running is not None:
            running.pop(process.pid, None)
            self.update_running_summary()
        record_event(
            process_manifest,
            "external_process_finished",
            stage="dxf",
            pid=process.pid,
            workspace=str(run),
            log=str(log),
            exit_code=exit_code,
        )
        if exit_code == 0:
            mark_needs_review(
                process_manifest,
                "dxf",
                "Orchestrator finished successfully. Review generated DWGs and logs, then record accepted artifacts.",
                [log],
            )
        else:
            item = process_manifest["stages"]["dxf"]
            item["status"] = "warning"
            item["notes"] = (
                f"Orchestrator failed with exit code {exit_code}. Review {log}. "
                "Review the Python orchestrator log and the endpoint-security event."
            )
            record_artifact(process_manifest, "dxf", log)
            messagebox.showerror("DXF orchestrator failed", item["notes"], parent=self)
        save_manifest(process_manifest, process_manifest_path)
        if self.manifest_path == process_manifest_path:
            self.manifest = process_manifest
            self.refresh()
        outcome = "finished and needs review" if exit_code == 0 else "failed"
        messagebox.showinfo(
            "DXF automation finished",
            f"Job {process_manifest['job']['number']} DXF automation {outcome}.\n\n"
            f"Log: {log}",
            parent=self,
        )

    def review_drawings(self) -> list[Path]:
        window = tk.Toplevel(self)
        window.title("Include or exclude drawing files")
        window.geometry("760x520")
        window.transient(self)
        window.grab_set()
        ttk.Label(
            window,
            text="Check only files that are safe plate inputs for automation.",
            padding=8,
        ).pack(anchor="w")
        frame = ttk.Frame(window)
        frame.pack(fill="both", expand=True, padx=8)
        canvas = tk.Canvas(frame)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas)
        body.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        variables = []
        for candidate in self.candidates:
            variable = tk.BooleanVar(value=candidate.selected)
            variables.append(variable)
            ttk.Checkbutton(
                body,
                text=f"{candidate.path.name} — {candidate.kind}",
                variable=variable,
            ).pack(anchor="w", pady=2)
        result: list[Path] = []

        def accept():
            result.extend(c.path for c, v in zip(self.candidates, variables) if v.get())
            window.destroy()

        buttons = ttk.Frame(window, padding=8)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Confirm Selection", command=accept).pack(side="right")
        ttk.Button(buttons, text="Cancel", command=window.destroy).pack(
            side="right", padx=5
        )
        self.wait_window(window)
        return result

    def run_comparison(self) -> None:
        nesting = self.manifest["paths"].get("nesting") or filedialog.askdirectory(
            title="Select the Nesting folder", parent=self
        )
        if not nesting:
            return
        if not self.manifest["paths"].get("nesting"):
            set_optional_path(self.manifest, "nesting", Path(nesting))
        parts = filedialog.askopenfilename(
            title="Select Parts List CSV", filetypes=[("CSV", "*.csv")], parent=self
        )
        solidworks = filedialog.askopenfilename(
            title="Select SolidWorks/model export CSV",
            filetypes=[("CSV", "*.csv")],
            parent=self,
        )
        if not parts or not solidworks or not self._accept_warnings("comparison"):
            return
        output = (
            Path(self.manifest["workspace"]["reports"])
            / f"comparison-{self.manifest['job']['revision']}"
        )
        output.mkdir(parents=True, exist_ok=True)
        log = Path(self.manifest["workspace"]["logs"]) / "comparison.log"
        command = command_comparison(
            sys.executable,
            self.repo,
            Path(nesting),
            Path(parts),
            Path(solidworks),
            output,
            self.bundled_tool("Engineering Production Comparison.exe"),
        )
        with log.open("a", encoding="utf-8") as handle:
            result = subprocess.run(
                command,
                cwd=self.repo,
                stdout=handle,
                stderr=subprocess.STDOUT,
                check=False,
            )
        record_event(
            self.manifest,
            "external_process_finished",
            stage="comparison",
            exit_code=result.returncode,
            log=str(log),
        )
        if result.returncode:
            raise JobError(
                f"Comparison exited with code {result.returncode}. Review {log}"
            )
        self.manifest["comparison"] = parse_comparison_summary(output)
        if self.manifest["comparison"]["status"] == "not_available":
            raise JobError(self.manifest["comparison"]["message"] + f" Review {log}")
        mark_needs_review(
            self.manifest,
            "comparison",
            f"{self.manifest['comparison']['message']}: {self.manifest['comparison']['errors']} errors, {self.manifest['comparison']['warnings']} warnings.",
            [log],
        )

    def launch_solidworks_macro(self, stage: str) -> None:
        macro = self.repo / (
            "solidworks/cad-batch-converter/Main.RunBatch.swp"
            if stage == "plate_model"
            else "solidworks/auto-bom/AutoBOMProperties.swp"
        )
        if not macro.is_file():
            raise JobError(f"Macro was not found: {macro}")
        if stage == "plate_model":
            prepared_root = self.manifest["stages"]["dxf"].get(
                "workspace", self.manifest["workspace"]["working"]
            )
            source_value = filedialog.askdirectory(
                title="Select the reviewed folder containing prepared DWGs",
                initialdir=prepared_root,
                parent=self,
            )
            if not source_value:
                return
            source = Path(source_value)
            if not any(source.glob("*.dwg")) and not any(source.glob("*.DWG")):
                raise JobError("The selected plate-input folder contains no DWG files.")
            thickness = simpledialog.askfloat(
                "Plate thickness",
                "Enter the extrusion thickness in inches for every DWG in "
                "this folder:",
                initialvalue=inferred_plate_thickness(source),
                minvalue=0.001,
                maxvalue=20.0,
                parent=self,
            )
            if thickness is None:
                return
            run_name = safe_name(source.name)
            filtered = (
                Path(self.manifest["workspace"]["working"])
                / "Filtered DWGs"
                / run_name
            )
            output = (
                Path(self.manifest["workspace"]["staging"])
                / "SolidWorks Parts"
                / run_name
            )
            filtered.mkdir(parents=True, exist_ok=True)
            output.mkdir(parents=True, exist_ok=True)
            values = {
                "MACROS_SOURCE_FOLDER": str(source),
                "MACROS_FILTERED_FOLDER": str(filtered),
                "MACROS_OUTPUT_FOLDER": str(output),
                "MACROS_EXTRUDE_DEPTH_METERS": format(thickness * 0.0254, ".12g"),
            }
        if not self._accept_warnings(stage):
            return
        if stage == "plate_model":
            os.environ.update(values)
            if sys.platform == "win32":
                key = r"HKCU\Software\VB and VBA Program Settings\EngineeringMacros\CadBatch"
                names = {
                    "MACROS_SOURCE_FOLDER": "SourceFolder",
                    "MACROS_FILTERED_FOLDER": "FilteredFolder",
                    "MACROS_OUTPUT_FOLDER": "OutputFolder",
                    "MACROS_EXTRUDE_DEPTH_METERS": "ExtrudeDepthMeters",
                }
                for env_name, value in values.items():
                    subprocess.run(
                        [
                            "reg.exe",
                            "add",
                            key,
                            "/v",
                            names[env_name],
                            "/t",
                            "REG_SZ",
                            "/d",
                            value,
                            "/f",
                        ],
                        check=True,
                        capture_output=True,
                    )
        if stage == "plate_model":
            solidworks = self.settings.get("solidworks_executable", "").strip()
            if solidworks and Path(solidworks).is_file():
                subprocess.Popen([solidworks])
            open_path(macro.parent)
            messagebox.showinfo(
                "Run one SolidWorks thickness group",
                plate_macro_instructions(
                    source, macro, thickness, len(plate_run_folders(source.parent))
                ),
                parent=self,
            )
            event = "cad_macro_guidance_opened"
            review_message = (
                "SolidWorks macro instructions opened. Run one configured "
                "thickness group, then review CAD output and BatchLog.txt."
            )
        else:
            open_path(macro)
            event = "cad_macro_launch_initiated"
            review_message = (
                "Macro launch initiated. Review CAD output and logs; launch "
                "does not mean the engineering task succeeded."
            )
        record_event(self.manifest, event, stage=stage, macro=str(macro))
        mark_needs_review(
            self.manifest,
            stage,
            review_message,
        )

    def open_stage_folder(self) -> None:
        def operation():
            self.require_job()
            stage = self.selected_stage()
            key = (
                "model_3d"
                if stage in {"plate_model", "manual_model", "assembly", "autobom"}
                else "nesting"
                if stage in {"nesting", "comparison"}
                and self.manifest["paths"].get("nesting")
                else "cut_files"
                if stage == "dxf"
                else "engineering_root"
            )
            open_path(Path(self.manifest["paths"][key]))

        self.handle(operation)

    def record_file(self) -> None:
        def operation():
            self.require_job()
            selected = filedialog.askopenfilename(
                title="Select a generated or reviewed file",
                initialdir=self.manifest["workspace"]["staging"],
                parent=self,
            )
            if selected:
                record_artifact(self.manifest, self.selected_stage(), Path(selected))
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()

        self.handle(operation)

    def finish_stage(self) -> None:
        def operation():
            self.require_job()
            notes = simpledialog.askstring(
                "Positive review confirmation",
                "Describe what you inspected and confirmed:",
                parent=self,
            )
            if notes is not None:
                stage = self.selected_stage()
                if stage == "bom":
                    workbook_value = self.manifest["stages"]["bom"].get(
                        "parts_list_workbook", ""
                    )
                    workbook = Path(workbook_value) if workbook_value else None
                    if workbook is None or not workbook.is_file():
                        selected = filedialog.askopenfilename(
                            title="Select the reviewed Parts List workbook",
                            filetypes=[("Excel workbook", "*.xlsx")],
                            parent=self,
                        )
                        if not selected:
                            raise JobError(
                                "Select the reviewed Parts List workbook before "
                                "completing this step."
                            )
                        workbook = Path(selected)
                        self.manifest["stages"]["bom"]["parts_list_workbook"] = str(
                            workbook
                        )
                    csv_path = export_parts_list_csv(
                        workbook,
                        Path(self.manifest["workspace"]["source_copies"])
                        / "Parts List.csv",
                    )
                    record_artifact(self.manifest, "bom", workbook)
                    record_artifact(self.manifest, "bom", csv_path)
                complete_stage(self.manifest, stage, notes)
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()

        self.handle(operation)

    def reopen(self) -> None:
        def operation():
            self.require_job()
            reason = simpledialog.askstring(
                "Reopen stage",
                "Why is this stage being reopened?",
                initialvalue="Additional work or revised inputs",
                parent=self,
            )
            if reason:
                reopen_stage(self.manifest, self.selected_stage(), reason)
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()

        self.handle(operation)

    def change_job_revision(self) -> None:
        def operation():
            self.require_job()
            value = simpledialog.askstring(
                "Change revision",
                "New revision (recorded artifacts retain their old revision):",
                initialvalue=self.manifest["job"]["revision"],
                parent=self,
            )
            if value:
                change_revision(self.manifest, value)
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()

        self.handle(operation)

    def set_optional_folder(self) -> None:
        def operation():
            self.require_job()
            key = simpledialog.askstring(
                "Optional folder",
                "Enter: nesting, part_checking, or forming",
                parent=self,
            )
            if not key:
                return
            folder = filedialog.askdirectory(
                title=f"Select {key.replace('_', ' ').title()} folder",
                initialdir=self.manifest["root"],
                parent=self,
            )
            if folder:
                set_optional_path(self.manifest, key.strip().lower(), Path(folder))
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()

        self.handle(operation)

    def promote(self) -> None:
        def operation():
            self.require_job()
            selected = filedialog.askopenfilenames(
                title="Select approved staged files",
                initialdir=self.manifest["workspace"]["staging"],
                parent=self,
            )
            if not selected:
                return
            staging_root = Path(self.manifest["workspace"]["staging"]).resolve()
            outside = [
                Path(path)
                for path in selected
                if not Path(path).resolve().is_relative_to(staging_root)
            ]
            if outside:
                raise JobError(
                    "Promotion only accepts files inside Assistant Staging. "
                    f"Not staged: {outside[0]}"
                )
            destination = filedialog.askdirectory(
                title="Select production destination folder", parent=self
            )
            if not destination:
                return
            plan = plan_promotions(map(Path, selected), Path(destination))
            self.show_promotion_table(plan)

        self.handle(operation)

    def show_promotion_table(self, plan: list[PromotionItem]) -> None:
        """Review and execute a promotion plan without sequential dialogs."""
        window = tk.Toplevel(self)
        window.title("Copy Approved Files to Production")
        window.geometry("1180x620")
        window.minsize(900, 480)
        window.transient(self)
        window.grab_set()

        ttk.Label(
            window,
            text=(
                "Review every source and destination. Conflicts default to Do Not "
                "Replace. Use Back Up + Replace only after engineering review."
            ),
            padding=10,
            wraplength=1120,
        ).pack(fill="x")

        columns = ("source", "destination", "conflict", "action", "result")
        tree_frame = ttk.Frame(window, padding=(10, 0))
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        widths = {
            "source": 300,
            "destination": 360,
            "conflict": 90,
            "action": 145,
            "result": 150,
        }
        for column in columns:
            tree.heading(column, text=column.replace("_", " ").title())
            tree.column(column, width=widths[column], anchor="w")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        actions: dict[str, str] = {}
        for index, item in enumerate(plan):
            iid = str(index)
            actions[iid] = item.action
            tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    str(item.source),
                    str(item.destination),
                    "YES" if item.conflict else "No",
                    self.promotion_action_label(item.action),
                    "Waiting for approval",
                ),
                tags=("conflict" if item.conflict else "new",),
            )
        tree.tag_configure("conflict", foreground="#a31621")
        tree.tag_configure("new", foreground="#176b32")
        tree.tag_configure("failed", foreground="#a31621")
        tree.tag_configure("finished", foreground="#176b32")

        detail = tk.StringVar(
            value="Select one or more rows, then choose the intended action."
        )
        ttk.Label(window, textvariable=detail, padding=(10, 7)).pack(fill="x")

        def set_action(action: str) -> None:
            selected_rows = tree.selection()
            if not selected_rows:
                detail.set("Select at least one file row first.")
                return
            rejected = 0
            for iid in selected_rows:
                item = plan[int(iid)]
                if action == "backup_replace" and not item.conflict:
                    rejected += 1
                    continue
                actions[iid] = action
                values = list(tree.item(iid, "values"))
                values[3] = self.promotion_action_label(action)
                values[4] = "Waiting for approval"
                tree.item(iid, values=values)
            detail.set(
                "Updated selected rows."
                if not rejected
                else f"Updated conflicts; {rejected} non-conflicting row(s) cannot use Back Up + Replace."
            )

        controls = ttk.Frame(window, padding=(10, 0, 10, 8))
        controls.pack(fill="x")
        ttk.Button(
            controls, text="Copy New File", command=lambda: set_action("copy")
        ).pack(side="left", padx=3)
        ttk.Button(
            controls, text="Do Not Replace", command=lambda: set_action("skip")
        ).pack(side="left", padx=3)
        ttk.Button(
            controls,
            text="Back Up Existing + Replace",
            command=lambda: set_action("backup_replace"),
        ).pack(side="left", padx=3)

        execute_button = ttk.Button(controls, text="Execute Approved Plan")
        execute_button.pack(side="right", padx=3)
        ttk.Button(controls, text="Close", command=window.destroy).pack(
            side="right", padx=3
        )

        def execute() -> None:
            approved = [
                PromotionItem(
                    item.source,
                    item.destination,
                    item.conflict,
                    actions[str(index)],
                )
                for index, item in enumerate(plan)
            ]
            copy_count = sum(item.action != "skip" for item in approved)
            if copy_count == 0:
                detail.set("Every row is Do Not Replace; there is nothing to copy.")
                return
            if not messagebox.askyesno(
                "Execute approved promotion plan?",
                f"{copy_count} file(s) will be copied or replaced. Existing files "
                "marked for replacement will be backed up first. Continue?",
                parent=window,
            ):
                return
            results = promote_files(self.manifest, approved)
            save_manifest(self.manifest, self.manifest_path)
            for index, result in enumerate(results):
                iid = str(index)
                values = list(tree.item(iid, "values"))
                values[4] = result["status"].replace("_", " ").title()
                if result.get("error"):
                    values[4] += f": {result['error']}"
                tree.item(
                    iid,
                    values=values,
                    tags=("failed" if result["status"] == "failed" else "finished",),
                )
            counts = {
                status: sum(result["status"] == status for result in results)
                for status in ("copied", "replaced", "skipped", "failed")
            }
            report = next((result.get("report") for result in results), "")
            detail.set(
                f"Finished: {counts['copied']} copied, {counts['replaced']} replaced, "
                f"{counts['skipped']} skipped, {counts['failed']} failed. "
                + (
                    f"Report: {report}"
                    if report
                    else "Promotion report could not be written; see row results and manifest events."
                )
            )
            execute_button.configure(state="disabled")
            self.refresh()

        execute_button.configure(command=lambda: self.handle(execute))

    @staticmethod
    def promotion_action_label(action: str) -> str:
        return {
            "copy": "Copy New File",
            "skip": "Do Not Replace",
            "backup_replace": "Back Up + Replace",
        }.get(action, action)

    def edit_settings(self) -> None:
        window = tk.Toplevel(self)
        window.title("Per-user settings (not stored in the repository)")
        window.geometry("820x330")
        window.transient(self)
        entries = {}
        labels = (
            ("macros_repo", "Shared Macros repository"),
            ("parts_list_template", "Standard Parts List template"),
            ("default_jobs_parent", "Default jobs parent"),
            ("autocad_executable", "AutoCAD executable"),
            ("autocad_console", "AutoCAD console executable"),
            ("solidworks_executable", "SolidWorks executable"),
        )
        for row, (key, label) in enumerate(labels):
            ttk.Label(window, text=label).grid(
                row=row, column=0, sticky="w", padx=8, pady=7
            )
            var = tk.StringVar(value=self.settings.get(key, ""))
            entries[key] = var
            ttk.Entry(window, textvariable=var, width=72).grid(
                row=row, column=1, sticky="ew", padx=5
            )

            def browse(k=key, v=var):
                value = (
                    filedialog.askdirectory(parent=window)
                    if k in {"macros_repo", "default_jobs_parent"}
                    else filedialog.askopenfilename(parent=window)
                )
                if value:
                    v.set(value)

            ttk.Button(window, text="Browse…", command=browse).grid(
                row=row, column=2, padx=5
            )
        window.columnconfigure(1, weight=1)

        def save():
            self.settings.update(
                {key: var.get().strip() for key, var in entries.items()}
            )
            save_settings(self.settings)
            window.destroy()

        ttk.Button(window, text="Save", command=lambda: self.handle(save)).grid(
            row=len(labels), column=2, pady=12
        )

    def open_workspace(self) -> None:
        self.handle(
            lambda: (
                self.require_job(),
                open_path(Path(self.manifest["workspace"]["assistant"])),
            )
        )

    def open_staging(self) -> None:
        self.handle(
            lambda: (
                self.require_job(),
                open_path(Path(self.manifest["workspace"]["staging"])),
            )
        )

    def open_logs(self) -> None:
        self.handle(
            lambda: (
                self.require_job(),
                open_path(Path(self.manifest["workspace"]["logs"])),
            )
        )

    def open_recent_file(self) -> None:
        def operation():
            self.require_job()
            selected = self.recent_tree.selection()
            if not selected:
                return
            index = int(selected[0].split("-")[-1])
            artifact = self.manifest.get("recent_files", [])[index]
            path = Path(artifact["path"])
            if not path.exists():
                raise JobError(f"The recorded file is no longer available: {path}")
            open_path(path)

        self.handle(operation)

    def open_comparison_report(self) -> None:
        def operation():
            self.require_job()
            comparison = self.manifest.get("comparison") or {}
            report = Path(comparison.get("html", ""))
            if not report.is_file():
                raise JobError(
                    "No completed HTML comparison report is recorded for this job."
                )
            open_path(report)

        self.handle(operation)


if __name__ == "__main__":
    JobAssistant().mainloop()
