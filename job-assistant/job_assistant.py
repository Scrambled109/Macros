"""Guided Windows interface for the engineering production workflow."""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from job_core import (
    STAGES,
    Check,
    JobError,
    change_revision,
    complete_stage,
    create_job,
    load_local_config,
    load_manifest,
    preflight,
    record_artifact,
    save_manifest,
    start_stage,
)


HERE = Path(__file__).resolve().parent
REPO = HERE.parent

STAGE_FOLDERS = {
    "bom": "source_boms",
    "dxf": "cut_files",
    "plate_model": "cad_models",
    "manual_model": "cad_models",
    "assembly": "cad_models",
    "autobom": "cad_models",
    "model_review": "exports",
    "nesting": "nests",
    "exports": "exports",
    "comparison": "reports",
    "final": "final_records",
}


class JobAssistant(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Engineering Job Assistant")
        self.geometry("1040x700")
        self.minsize(900, 580)
        self.manifest = None
        self.manifest_path: Path | None = None
        self.config = load_local_config(REPO)
        self._build()

    def _build(self) -> None:
        toolbar = ttk.Frame(self, padding=10)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="New Job", command=self.new_job).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Open Job", command=self.open_job).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Change Revision", command=self.change_job_revision).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Record Artifact", command=self.record_selected_artifact).pack(side="left", padx=3)

        self.heading = ttk.Label(self, text="Open or create a job to begin.", font=("Segoe UI", 13, "bold"))
        self.heading.pack(fill="x", padx=14, pady=(2, 8))

        columns = ("stage", "status", "review", "artifacts")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=13)
        for name, width in (("stage", 280), ("status", 120), ("review", 120), ("artifacts", 90)):
            self.tree.heading(name, text=name.title())
            self.tree.column(name, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=14)

        actions = ttk.Frame(self, padding=10)
        actions.pack(fill="x")
        ttk.Button(actions, text="Run Preflight", command=self.run_preflight).pack(side="left", padx=3)
        ttk.Button(actions, text="Open Stage Folder", command=self.open_stage_folder).pack(side="left", padx=3)
        ttk.Button(actions, text="Launch Stage Tool", command=self.launch_stage).pack(side="left", padx=3)
        ttk.Button(actions, text="Complete Reviewed Stage", command=self.finish_stage).pack(side="left", padx=3)

        self.output = tk.Text(self, height=12, wrap="word", state="disabled")
        self.output.pack(fill="both", padx=14, pady=(0, 14))

    def log(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.insert("end", text)
        self.output.configure(state="disabled")

    def selected_stage(self) -> str:
        selected = self.tree.selection()
        if not selected:
            raise JobError("Select a workflow stage first.")
        return selected[0]

    def require_job(self) -> None:
        if self.manifest is None or self.manifest_path is None:
            raise JobError("Open or create a job first.")

    def handle(self, operation) -> None:
        try:
            operation()
        except (JobError, OSError, subprocess.SubprocessError) as exc:
            messagebox.showerror("Job Assistant", str(exc), parent=self)

    def new_job(self) -> None:
        def operation() -> None:
            parent_value = filedialog.askdirectory(
                title="Choose the parent jobs folder",
                initialdir=self.config.get("default_jobs_directory") or None,
                parent=self,
            )
            if not parent_value:
                return
            number = simpledialog.askstring("New Job", "Job number:", parent=self)
            name = simpledialog.askstring("New Job", "Job name:", parent=self)
            revision = simpledialog.askstring("New Job", "Initial revision:", initialvalue="A", parent=self)
            if not all((number, name, revision)):
                return
            template_value = self.config.get("templates_directory", "")
            templates = Path(template_value) if template_value else None
            path = create_job(Path(parent_value), number, name, revision, templates)
            self._load(path)
        self.handle(operation)

    def open_job(self) -> None:
        def operation() -> None:
            selected = filedialog.askopenfilename(
                title="Open job manifest",
                filetypes=[("Job manifest", "job_manifest.json"), ("JSON", "*.json")],
                parent=self,
            )
            if selected:
                self._load(Path(selected))
        self.handle(operation)

    def change_job_revision(self) -> None:
        def operation() -> None:
            self.require_job()
            revision = simpledialog.askstring(
                "Change Revision",
                "New job revision (existing artifacts retain their recorded revision):",
                initialvalue=self.manifest["job"]["revision"],
                parent=self,
            )
            if revision:
                change_revision(self.manifest, revision)
                save_manifest(self.manifest, self.manifest_path)
                self.refresh()
        self.handle(operation)

    def _load(self, path: Path) -> None:
        self.manifest = load_manifest(path)
        self.manifest_path = path
        self.refresh()

    def refresh(self) -> None:
        def operation() -> None:
            self.require_job()
            self.manifest = load_manifest(self.manifest_path)
            self.heading.configure(
                text=f"{self.manifest['job']['number']} — {self.manifest['job']['name']}  |  Revision {self.manifest['job']['revision']}"
            )
            self.tree.delete(*self.tree.get_children())
            for key, label in STAGES:
                item = self.manifest["stages"][key]
                self.tree.insert(
                    "", "end", iid=key,
                    values=(label, item["status"].replace("_", " ").title(), "Yes" if item["reviewed"] else "No", len(item["artifacts"])),
                )
        self.handle(operation)

    def _checks(self, stage: str) -> list[Check]:
        self.require_job()
        checks = preflight(self.manifest, stage, self.config)
        self.log("\n".join(f"[{check.level.upper()}] {check.message}" + (f"\n  Fix: {check.correction}" if check.correction else "") for check in checks))
        return checks

    def run_preflight(self) -> None:
        self.handle(lambda: self._checks(self.selected_stage()))

    def open_stage_folder(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            path = self.manifest["paths"][STAGE_FOLDERS[stage]]
            os.startfile(path)
        self.handle(operation)

    def launch_stage(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            checks = self._checks(stage)
            blocks = [check for check in checks if check.level == "block"]
            warnings = [check for check in checks if check.level == "warning"]
            if blocks:
                raise JobError("Preflight has blocking errors. Correct them before launching.")
            if warnings and not messagebox.askyesno(
                "Acknowledge warnings",
                "Preflight reported warnings. Have you reviewed them and do you want to continue?",
                parent=self,
            ):
                return
            start_stage(self.manifest, stage, warnings)
            save_manifest(self.manifest, self.manifest_path)
            self._launch(stage)
            self.refresh()
        self.handle(operation)

    def _launch(self, stage: str) -> None:
        commands = {
            "bom": [sys.executable, str(REPO / "data-tools/bom-converter/bom_converter.py")],
            "dxf": ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(REPO / "autocad/dxf-orchestrator/Master_Orchestrator.ps1")],
            "comparison": [sys.executable, str(REPO / "data-tools/production-comparison/compare_production_parts.py")],
        }
        swp = {
            "plate_model": REPO / "solidworks/cad-batch-converter/Main.RunBatch.swp",
            "autobom": REPO / "solidworks/auto-bom/AutoBOMProperties.swp",
        }
        if stage in commands:
            working_directory = (
                self.manifest["paths"]["cut_files"]
                if stage == "dxf"
                else self.manifest["root"]
            )
            subprocess.Popen(commands[stage], cwd=working_directory)
        elif stage in swp:
            if stage == "plate_model":
                source = Path(self.manifest["paths"]["cut_files"])
                filtered = source / "FilteredDWGs"
                filtered.mkdir(exist_ok=True)
                key_path = r"Software\VB and VBA Program Settings\EngineeringMacros\CadBatch"
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    winreg.SetValueEx(key, "SourceFolder", 0, winreg.REG_SZ, str(source))
                    winreg.SetValueEx(key, "FilteredFolder", 0, winreg.REG_SZ, str(filtered))
                    winreg.SetValueEx(key, "OutputFolder", 0, winreg.REG_SZ, self.manifest["paths"]["cad_models"])
                os.environ["MACROS_SOURCE_FOLDER"] = str(source)
                os.environ["MACROS_FILTERED_FOLDER"] = str(filtered)
                os.environ["MACROS_OUTPUT_FOLDER"] = self.manifest["paths"]["cad_models"]
            os.startfile(str(swp[stage]))
        else:
            os.startfile(self.manifest["paths"][STAGE_FOLDERS[stage]])
            messagebox.showinfo("Manual checkpoint", "The stage folder was opened. Complete the documented review before marking this stage complete.", parent=self)

    def record_selected_artifact(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            selected = filedialog.askopenfilename(title="Select an artifact produced or approved in this stage", parent=self)
            if not selected:
                return
            record_artifact(self.manifest, stage, Path(selected))
            save_manifest(self.manifest, self.manifest_path)
            self.refresh()
        self.handle(operation)

    def finish_stage(self) -> None:
        def operation() -> None:
            self.require_job()
            stage = self.selected_stage()
            notes = simpledialog.askstring("Review checkpoint", "Describe what you reviewed and any approved exceptions:", parent=self)
            if notes is None:
                return
            complete_stage(self.manifest, stage, notes)
            save_manifest(self.manifest, self.manifest_path)
            self.refresh()
        self.handle(operation)


if __name__ == "__main__":
    JobAssistant().mainloop()
