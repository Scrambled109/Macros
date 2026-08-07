"""Windows-oriented dashboard and guided stages for engineering jobs."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import uuid
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from job_core import (
    MANIFEST_NAME,
    STAGES,
    STAGE_GUIDANCE,
    JobError,
    OutputMoveItem,
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
    normalize_plate_model_filenames,
    inferred_plate_thickness,
    parse_comparison_summary,
    plan_completed_outputs,
    prepare_dxf_workspace,
    move_completed_outputs,
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
STEEL_BLUE = "#57a0d3"
DARK_BG = "#070b0f"
DARK_SURFACE = "#0d141a"
DARK_CARD = "#121b23"
DARK_INPUT = "#080d12"
TEXT_PRIMARY = "#eef5f9"
TEXT_MUTED = "#a9bac6"


def open_path(path: Path) -> None:
    if sys.platform != "win32":
        raise JobError("Open Folder is available in the packaged Windows application.")
    os.startfile(str(path))


class JobAssistant(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Engineering Job Assistant")
        self.geometry("980x650")
        self.minsize(820, 540)
        self.settings = load_settings(repo_root=DEFAULT_REPO)
        self.manifest: dict | None = None
        self.manifest_path: Path | None = None
        self.active_stage: str | None = None
        self.candidates = []
        self.running_processes: dict[int, dict] = {}
        self.notice_path: Path | None = None
        self.stage_detail_text = ""
        self.recent_tree: ttk.Treeview | None = None
        self.dashboard_recent_tree: ttk.Treeview | None = None
        self.dashboard_recent_files: list[dict] = []
        self.activity_log_path: Path | None = None
        self.activity_log_signature: tuple | None = None
        self._configure_style()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.close_application)

    @property
    def repo(self) -> Path:
        configured = Path(self.settings.get("macros_repo") or DEFAULT_REPO)
        # Older settings can retain a disconnected network drive or a checkout
        # that was moved.  The running assistant's own checkout is a safe source
        # fallback and prevents child tools from exiting immediately because
        # their script path no longer exists.
        if (configured / "job-assistant" / "job_core.py").is_file():
            return configured
        return DEFAULT_REPO

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

    def _configure_style(self) -> None:
        self.configure(background=DARK_BG)
        self.option_add("*Menu.background", DARK_SURFACE)
        self.option_add("*Menu.foreground", TEXT_PRIMARY)
        self.option_add("*Menu.activeBackground", STEEL_BLUE)
        self.option_add("*Menu.activeForeground", "#ffffff")
        self.option_add("*Toplevel.background", DARK_BG)
        style = ttk.Style(self)
        try:
            # Clam is intentionally used on Windows too.  It gives the assistant
            # a consistent, flat ttk surface and avoids the tall, beveled Vista
            # controls that made the dashboard look crowded.
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=DARK_BG)
        style.configure("Card.TFrame", background=DARK_CARD)
        style.configure("Brand.TFrame", background=STEEL_BLUE)
        style.configure(
            "TLabel",
            background=DARK_BG,
            foreground=TEXT_PRIMARY,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Card.TLabel",
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Muted.TLabel",
            background=DARK_BG,
            foreground=TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "CardMuted.TLabel",
            background=DARK_CARD,
            foreground=TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "TButton",
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            bordercolor="#40515f",
            lightcolor=DARK_CARD,
            darkcolor=DARK_CARD,
            font=("Segoe UI", 10),
            padding=(11, 7),
            relief="flat",
        )
        style.map(
            "TButton",
            background=[("active", "#2b3c49"), ("pressed", "#344a5a")],
            bordercolor=[("focus", STEEL_BLUE), ("active", "#668094")],
        )
        style.configure(
            "TMenubutton",
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            bordercolor="#40515f",
            lightcolor=DARK_CARD,
            darkcolor=DARK_CARD,
            font=("Segoe UI", 10),
            padding=(10, 7),
            relief="flat",
        )
        style.map("TMenubutton", background=[("active", "#2b3c49")])
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10),
            foreground="#ffffff",
            background=STEEL_BLUE,
            bordercolor=STEEL_BLUE,
            lightcolor=STEEL_BLUE,
            darkcolor=STEEL_BLUE,
            padding=(14, 8),
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            foreground=[("disabled", "#c4d4df")],
            background=[
                ("disabled", "#42677f"),
                ("active", "#6eb1df"),
                ("pressed", "#438bbd"),
            ],
        )
        style.configure(
            "Heading.TLabel",
            background=DARK_BG,
            foreground=TEXT_PRIMARY,
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "Next.TLabel",
            background=DARK_BG,
            foreground="#c7d5de",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Notice.TLabel",
            background="#173a2b",
            foreground="#a4e5bd",
            padding=(10, 7),
        )
        style.configure(
            "Warning.TLabel",
            background="#493719",
            foreground="#ffd28a",
            padding=(10, 7),
        )
        style.configure(
            "StepTitle.TLabel",
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            font=("Segoe UI Semibold", 12),
        )
        style.configure(
            "Status.TLabel",
            background=DARK_CARD,
            foreground=TEXT_MUTED,
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "StatusComplete.TLabel",
            background=DARK_CARD,
            foreground="#72d99a",
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "StatusAttention.TLabel",
            background=DARK_CARD,
            foreground="#ffc36f",
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "Horizontal.TProgressbar",
            background=STEEL_BLUE,
            troughcolor="#293944",
            bordercolor="#293944",
            lightcolor=STEEL_BLUE,
            darkcolor=STEEL_BLUE,
        )
        style.configure(
            "Treeview",
            background=DARK_CARD,
            fieldbackground=DARK_CARD,
            foreground=TEXT_PRIMARY,
            borderwidth=0,
            rowheight=38,
            font=("Segoe UI", 10),
        )
        style.map(
            "Treeview",
            background=[("selected", "#315f7d")],
            foreground=[("selected", "#ffffff")],
        )
        style.configure(
            "Treeview.Heading",
            background="#293944",
            foreground="#d6e2e9",
            borderwidth=0,
            font=("Segoe UI Semibold", 9),
            padding=(8, 7),
        )
        style.configure(
            "BrandTitle.TLabel",
            background=STEEL_BLUE,
            foreground="#ffffff",
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "BrandMuted.TLabel",
            background=STEEL_BLUE,
            foreground="#eaf6fc",
            font=("Segoe UI", 9),
        )
        style.configure(
            "TEntry",
            fieldbackground=DARK_INPUT,
            foreground=TEXT_PRIMARY,
            insertcolor=TEXT_PRIMARY,
            bordercolor="#40515f",
        )
        style.configure(
            "TCombobox",
            fieldbackground=DARK_INPUT,
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            arrowcolor=TEXT_PRIMARY,
        )
        style.configure(
            "TSpinbox",
            fieldbackground=DARK_INPUT,
            background=DARK_CARD,
            foreground=TEXT_PRIMARY,
            arrowcolor=TEXT_PRIMARY,
        )

    def _build(self) -> None:
        self._build_menus()

        content = ttk.Frame(self, padding=(22, 18, 22, 16))
        content.pack(fill="both", expand=True)

        header = ttk.Frame(content, style="Brand.TFrame", padding=(16, 10))
        header.pack(fill="x", pady=(0, 8))
        self.running_summary = ttk.Label(
            header,
            text="",
            style="BrandMuted.TLabel",
        )
        self.running_summary.pack(side="right", padx=(12, 0))
        self.running_summary.pack_forget()
        self.heading = ttk.Label(
            header,
            text="Engineering Job Assistant",
            style="BrandTitle.TLabel",
        )
        self.heading.pack(side="left")
        logo_path = HERE / "assets" / "steel_america_logo.png"
        self.logo_image = None
        if logo_path.is_file():
            try:
                self.logo_image = tk.PhotoImage(file=str(logo_path)).subsample(3, 3)
                ttk.Label(
                    header,
                    image=self.logo_image,
                    style="BrandMuted.TLabel",
                ).pack(side="right", padx=(14, 0))
            except tk.TclError:
                self.logo_image = None

        progress_row = ttk.Frame(content)
        progress_row.pack(fill="x", pady=(8, 14))
        self.next_action = ttk.Label(
            progress_row,
            text="Set up a new job or open an existing job.",
            style="Next.TLabel",
            wraplength=680,
        )
        self.next_action.pack(side="left", fill="x", expand=True)
        self.progress_text = ttk.Label(
            progress_row,
            text="No job open",
            style="Muted.TLabel",
        )
        self.progress_text.pack(side="right", padx=(12, 0))
        self.progress = ttk.Progressbar(progress_row, length=130)
        self.progress.pack(side="right", padx=(12, 0))

        self.message_area = ttk.Frame(content)
        self.message_area.pack(fill="x")
        self.notice_row = ttk.Frame(self.message_area)
        self.notice_label = ttk.Label(
            self.notice_row,
            text="",
            style="Notice.TLabel",
            wraplength=760,
        )
        self.notice_label.pack(side="left", fill="x", expand=True)
        self.notice_open_button = ttk.Button(
            self.notice_row,
            text="Open Result",
            command=self.open_notice_path,
        )
        self.warning_summary = ttk.Label(
            self.message_area,
            text="",
            style="Warning.TLabel",
            wraplength=900,
        )

        self.workflow_card = ttk.Frame(content, style="Card.TFrame", padding=1)
        self.workflow_card.pack(fill="x")
        card_header = ttk.Frame(
            self.workflow_card,
            style="Card.TFrame",
            padding=(14, 11, 14, 8),
        )
        card_header.pack(fill="x")
        ttk.Label(
            card_header,
            text="Workflow",
            style="StepTitle.TLabel",
        ).pack(side="left")
        ttk.Label(
            card_header,
            text="Select a step to work on it",
            style="CardMuted.TLabel",
        ).pack(side="right")

        self.tree = ttk.Treeview(
            self.workflow_card,
            columns=("stage", "status"),
            show="headings",
            height=len(STAGES),
        )
        self.tree.heading("stage", text="Step", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")
        self.tree.column("stage", width=650, minwidth=300, anchor="w")
        self.tree.column("status", width=170, minwidth=120, anchor="w")
        self.tree.pack(fill="x", padx=1)
        self.tree.bind("<<TreeviewSelect>>", self._stage_selected)
        self.tree.tag_configure("not_started", foreground="#98a8b3")
        self.tree.tag_configure("ready", foreground="#87c9f3")
        self.tree.tag_configure("in_progress", foreground="#ffd078")
        self.tree.tag_configure("needs_review", foreground="#ffb66d")
        self.tree.tag_configure("complete", foreground="#79dfa1")
        self.tree.tag_configure("warning", foreground="#ff8d95")

        ttk.Separator(self.workflow_card).pack(fill="x")
        selected = ttk.Frame(
            self.workflow_card,
            style="Card.TFrame",
            padding=(15, 12, 15, 14),
        )
        selected.pack(fill="x")
        selected_text = ttk.Frame(selected, style="Card.TFrame")
        selected_text.pack(side="left", fill="both", expand=True)
        selected_heading = ttk.Frame(selected_text, style="Card.TFrame")
        selected_heading.pack(fill="x")
        self.step_title = ttk.Label(
            selected_heading,
            text="No job open",
            style="StepTitle.TLabel",
        )
        self.step_title.pack(side="left")
        self.step_status = ttk.Label(
            selected_heading,
            text="",
            style="Status.TLabel",
        )
        self.step_status.pack(side="left", padx=(10, 0))
        self.step_summary = ttk.Label(
            selected_text,
            text="Choose Set Up Job or Open Existing to begin.",
            style="CardMuted.TLabel",
            wraplength=650,
        )
        self.step_summary.pack(anchor="w", fill="x", pady=(4, 0))

        self.action_buttons = ttk.Frame(selected, style="Card.TFrame")
        self.action_buttons.pack(side="right", padx=(14, 0))
        self.setup_action = ttk.Button(
            self.action_buttons,
            text="Set Up Job",
            command=self.setup_job,
            style="Primary.TButton",
        )
        self.setup_action.pack(side="left", padx=(0, 6))
        self.open_action = ttk.Button(
            self.action_buttons,
            text="Open Existing",
            command=self.open_job,
        )
        self.open_action.pack(side="left")
        self.primary_step_button = ttk.Button(
            self.action_buttons,
            text="Start Step",
            command=self.run_stage,
            style="Primary.TButton",
        )
        self.more_button = ttk.Menubutton(self.action_buttons, text="More")
        self.more_menu = tk.Menu(self.more_button, tearoff=False)
        self._populate_step_menu(self.more_menu)
        self.more_button.configure(menu=self.more_menu)

        dashboard = ttk.Frame(content)
        dashboard.pack(fill="both", expand=True, pady=(12, 0))
        dashboard.columnconfigure(0, weight=3)
        dashboard.columnconfigure(1, weight=2)
        dashboard.rowconfigure(0, weight=1)

        activity_card = ttk.Frame(
            dashboard, style="Card.TFrame", padding=(14, 11, 14, 14)
        )
        activity_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        activity_card.columnconfigure(0, weight=1)
        activity_card.rowconfigure(1, weight=1)
        activity_header = ttk.Frame(activity_card, style="Card.TFrame")
        activity_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            activity_header, text="Activity", style="StepTitle.TLabel"
        ).pack(side="left")
        self.activity_status = ttk.Label(
            activity_header, text="No process running", style="CardMuted.TLabel"
        )
        self.activity_status.pack(side="right")
        activity_body = ttk.Frame(activity_card, style="Card.TFrame")
        activity_body.grid(row=1, column=0, sticky="nsew")
        activity_body.columnconfigure(0, weight=1)
        activity_body.rowconfigure(0, weight=1)
        self.activity_text = tk.Text(
            activity_body,
            wrap="none",
            height=10,
            padx=10,
            pady=9,
            relief="flat",
            background=DARK_INPUT,
            foreground="#c9d7df",
            insertbackground=TEXT_PRIMARY,
            selectbackground="#315f7d",
            font=("Cascadia Mono", 9),
        )
        activity_scroll = ttk.Scrollbar(
            activity_body, orient="vertical", command=self.activity_text.yview
        )
        self.activity_text.configure(yscrollcommand=activity_scroll.set)
        self.activity_text.grid(row=0, column=0, sticky="nsew")
        activity_scroll.grid(row=0, column=1, sticky="ns")
        self._set_activity_text(
            "Run a workflow step to see its live log here. The latest job log "
            "remains visible after the process finishes."
        )

        side = ttk.Frame(dashboard)
        side.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        side.columnconfigure(0, weight=1)
        side.rowconfigure(0, weight=1)

        recent_card = ttk.Frame(
            side, style="Card.TFrame", padding=(14, 11, 14, 12)
        )
        recent_card.grid(row=0, column=0, sticky="nsew")
        recent_card.columnconfigure(0, weight=1)
        recent_card.rowconfigure(1, weight=1)
        recent_header = ttk.Frame(recent_card, style="Card.TFrame")
        recent_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            recent_header, text="Recent outputs", style="StepTitle.TLabel"
        ).pack(side="left")
        ttk.Button(
            recent_header,
            text="Open Selected",
            command=self.open_dashboard_recent_file,
        ).pack(side="right")
        self.dashboard_recent_tree = ttk.Treeview(
            recent_card,
            columns=("file", "revision"),
            show="headings",
            height=5,
        )
        self.dashboard_recent_tree.heading("file", text="File", anchor="w")
        self.dashboard_recent_tree.heading("revision", text="Rev", anchor="w")
        self.dashboard_recent_tree.column("file", width=300, minwidth=150, anchor="w")
        self.dashboard_recent_tree.column(
            "revision", width=55, minwidth=45, anchor="center", stretch=False
        )
        self.dashboard_recent_tree.grid(row=1, column=0, sticky="nsew")
        self.dashboard_recent_tree.bind(
            "<Double-1>", lambda _event: self.open_dashboard_recent_file()
        )

        folders_card = ttk.Frame(
            side, style="Card.TFrame", padding=(14, 11, 14, 12)
        )
        folders_card.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(
            folders_card, text="Job folders", style="StepTitle.TLabel"
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        folder_specs = (
            ("Engineering", "paths", "engineering_root"),
            ("Cut Files", "paths", "cut_files"),
            ("3D Models", "paths", "model_3d"),
            ("Staging", "workspace", "staging"),
            ("Logs", "workspace", "logs"),
        )
        self.folder_buttons = []
        for index, (label, section, key) in enumerate(folder_specs):
            button = ttk.Button(
                folders_card,
                text=label,
                command=lambda s=section, k=key: self.open_job_location(s, k),
                state="disabled",
            )
            button.grid(
                row=1 + index // 3,
                column=index % 3,
                sticky="ew",
                padx=(0 if index % 3 == 0 else 5, 0),
                pady=(0, 5),
            )
            self.folder_buttons.append(button)
        for column in range(3):
            folders_card.columnconfigure(column, weight=1)
        self.after(750, self.refresh_activity_panel)

    def _build_menus(self) -> None:
        menu_bar = tk.Menu(self, tearoff=False)

        job_menu = tk.Menu(menu_bar, tearoff=False)
        job_menu.add_command(label="Set Up / Attach Job…", command=self.setup_job)
        job_menu.add_command(label="Open Job…", command=self.open_job)
        job_menu.add_command(label="Change Revision…", command=self.change_job_revision)
        job_menu.add_separator()
        job_menu.add_command(label="Settings…", command=self.edit_settings)
        job_menu.add_separator()
        job_menu.add_command(label="Exit", command=self.close_application)
        menu_bar.add_cascade(label="Job", menu=job_menu)

        step_menu = tk.Menu(menu_bar, tearoff=False)
        self._populate_step_menu(step_menu)
        menu_bar.add_cascade(label="Step", menu=step_menu)

        tools_menu = tk.Menu(menu_bar, tearoff=False)
        tools_menu.add_command(label="Job Folders…", command=self.set_optional_folder)
        tools_menu.add_separator()
        tools_menu.add_command(
            label="Open Assistant Workspace", command=self.open_workspace
        )
        tools_menu.add_command(label="Open Staging", command=self.open_staging)
        tools_menu.add_command(label="Open Logs", command=self.open_logs)
        tools_menu.add_command(
            label="Open Comparison Report", command=self.open_comparison_report
        )
        menu_bar.add_cascade(label="Tools", menu=tools_menu)

        view_menu = tk.Menu(menu_bar, tearoff=False)
        view_menu.add_command(label="Job Details…", command=self.show_job_details)
        view_menu.add_command(
            label="Selected Step Technical Details…",
            command=self.show_technical_details,
        )
        view_menu.add_separator()
        view_menu.add_command(label="Refresh", command=self.refresh, accelerator="F5")
        menu_bar.add_cascade(label="View", menu=view_menu)

        self.configure(menu=menu_bar)
        self.bind("<F5>", lambda _event: self.refresh())

    def _populate_step_menu(self, menu: tk.Menu) -> None:
        menu.add_command(label="Start Selected Step", command=self.run_stage)
        menu.add_command(label="Check Readiness…", command=self.run_checks)
        menu.add_command(
            label="Move Completed Outputs…", command=self.move_outputs
        )
        menu.add_separator()
        menu.add_command(label="Open Step Folder", command=self.open_stage_folder)
        menu.add_command(label="Record File…", command=self.record_file)
        menu.add_command(label="Mark Complete…", command=self.finish_stage)
        menu.add_command(label="Reopen Step…", command=self.reopen)
        menu.add_separator()
        menu.add_command(
            label="Technical Details…", command=self.show_technical_details
        )

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
            self.running_summary.pack_forget()
            return
        jobs = sorted(
            {item["job_number"] for item in self.running_processes.values()}
        )
        text = f"Running {count} for {', '.join(jobs)}"
        self.running_summary.configure(text=text)
        if not self.running_summary.winfo_manager():
            self.running_summary.pack(side="right", padx=(12, 0))

    def post_background_notice(
        self,
        title: str,
        message: str,
        *,
        level: str = "info",
        path: Path | None = None,
    ) -> None:
        """Show process completion in the dashboard without a blocking dialog.

        A modal completion message can land behind another application on
        Windows. Tk then disables the assistant while that hidden dialog waits,
        making the entire dashboard appear frozen. Background work always posts
        here instead; dialogs remain reserved for operator-initiated decisions.
        """

        self.notice_path = Path(path) if path else None
        self.notice_label.configure(
            text=f"{title}: {message}",
            style="Warning.TLabel" if level == "warning" else "Notice.TLabel",
        )
        if not self.notice_row.winfo_manager():
            self.notice_row.pack(fill="x", pady=(0, 8))
        if level == "warning":
            # The completion notice already contains the actionable failure.
            # Do not stack the dashboard's summary of the same warning above it.
            self.warning_summary.pack_forget()
        if self.notice_path:
            if not self.notice_open_button.winfo_manager():
                self.notice_open_button.pack(side="right", padx=(8, 0))
        else:
            self.notice_open_button.pack_forget()
        self.bell()

    def open_notice_path(self) -> None:
        def operation() -> None:
            if not self.notice_path:
                raise JobError("There is no result path to open yet.")
            target = self.notice_path
            if target.is_file():
                target = target.parent
            if not target.exists():
                raise JobError(f"The result location is no longer available: {target}")
            open_path(target)

        self.handle(operation)

    def open_review_and_offer_completion(
        self,
        stage: str,
        result_path: Path,
        process_manifest: dict,
        process_manifest_path: Path,
        *,
        ask_to_complete: bool = True,
    ) -> bool:
        """Open required review output, then offer one-click completion.

        Completion prompts are only shown for the job still displayed in the
        assistant. Background work for another job remains available through
        its dashboard notice without mutating the currently open job.
        """

        if self.manifest_path != process_manifest_path:
            return False
        # Unit-level process polling tests use an uninitialized Tk subclass.
        # The real application always owns ``tk``; skipping UI work here keeps
        # the workflow mutation logic independently testable.
        if "tk" not in self.__dict__:
            return False
        target = Path(result_path)
        if target.exists():
            try:
                open_path(target)
            except (JobError, OSError):
                # The dashboard's Open Result button remains available if the
                # Windows shell cannot open the result automatically.
                pass
        self.deiconify()
        self.lift()
        self.focus_force()
        if not ask_to_complete:
            return False
        label = process_manifest["stages"][stage]["label"]
        if not messagebox.askyesno(
            "Mark this step complete?",
            f"The required review output for:\n\n{label}\n\nhas been opened. "
            "After checking it, select Yes to mark the step complete. "
            "Select No if corrections are still needed.",
            parent=self,
        ):
            return False
        if stage == "bom":
            workbook = Path(
                process_manifest["stages"]["bom"].get(
                    "parts_list_workbook", result_path
                )
            )
            csv_path = export_parts_list_csv(
                workbook,
                Path(process_manifest["workspace"]["source_copies"])
                / "Parts List.csv",
            )
            record_artifact(process_manifest, "bom", workbook)
            record_artifact(process_manifest, "bom", csv_path)
        complete_stage(
            process_manifest,
            stage,
            "Operator confirmed the automatically opened review output.",
        )
        save_manifest(process_manifest, process_manifest_path)
        self.manifest = process_manifest
        self.refresh()
        self.show_stage()
        return True

    def close_application(self) -> None:
        if self.running_processes and not messagebox.askyesno(
            "Processes are still running",
            "External automation is still running. Closing the assistant will "
            "stop status monitoring and completion notifications, but will not "
            "stop the external tools.\n\nClose the assistant anyway?",
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
        title = job["number"]
        if job.get("name") and job["name"] != job["number"]:
            title += f" — {job['name']}"
        self.heading.configure(
            text=f"{title}  ·  Rev {job['revision']}"
        )
        self.next_action.configure(text=recommended_next_action(self.manifest))
        warnings = dashboard_warnings(self.manifest)
        if warnings:
            extra = f"  (+{len(warnings) - 1} more in Job Details)" if len(warnings) > 1 else ""
            self.warning_summary.configure(
                text=f"Needs attention: {warnings[0]}{extra}"
            )
            if not self.warning_summary.winfo_manager():
                self.warning_summary.pack(fill="x", pady=(0, 8))
        else:
            self.warning_summary.pack_forget()
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
        self.setup_action.pack_forget()
        self.open_action.pack_forget()
        if not self.primary_step_button.winfo_manager():
            self.primary_step_button.pack(side="left", padx=(0, 6))
        if not self.more_button.winfo_manager():
            self.more_button.pack(side="left")
        self.refresh_recent_outputs()
        for button in self.folder_buttons:
            button.configure(state="normal")
        self.show_stage()

    def _set_activity_text(self, content: str) -> None:
        self.activity_text.configure(state="normal")
        self.activity_text.delete("1.0", "end")
        self.activity_text.insert("1.0", content)
        self.activity_text.configure(state="disabled")
        self.activity_text.see("end")

    def refresh_activity_panel(self) -> None:
        """Tail the active process log, or retain the most recent job log."""
        try:
            log = None
            stage = None
            if self.manifest_path:
                for process in reversed(list(self.running_processes.values())):
                    if Path(process.get("manifest_path", "")) == self.manifest_path:
                        candidate = Path(process.get("log", ""))
                        if candidate.is_file():
                            log = candidate
                            stage = process.get("stage")
                            break
                if log is None and self.manifest:
                    logs = Path(self.manifest["workspace"]["logs"])
                    candidates = [item for item in logs.glob("*.log") if item.is_file()]
                    if candidates:
                        log = max(candidates, key=lambda item: item.stat().st_mtime)
            if log and log.is_file():
                stat = log.stat()
                signature = (log, stat.st_mtime_ns, stat.st_size)
                if signature != self.activity_log_signature:
                    with log.open("rb") as handle:
                        start = max(0, stat.st_size - 65536)
                        handle.seek(start)
                        data = handle.read()
                    if start:
                        data = data.split(b"\n", 1)[-1]
                    lines = data.decode("utf-8", errors="replace").splitlines()
                    self._set_activity_text(
                        "\n".join(lines[-80:]) or "The log is empty."
                    )
                    self.activity_log_signature = signature
                self.activity_log_path = log
                if stage:
                    labels = dict(STAGES)
                    self.activity_status.configure(
                        text=f"Running · {labels.get(stage, stage)}"
                    )
                else:
                    self.activity_status.configure(text=f"Latest · {log.name}")
            elif self.manifest:
                self.activity_status.configure(text="No process running")
                self._set_activity_text(
                    "No log has been created for this job yet. Start a step and "
                    "its output will stream here."
                )
        except (OSError, tk.TclError):
            pass
        finally:
            try:
                self.after(750, self.refresh_activity_panel)
            except tk.TclError:
                pass

    def refresh_recent_outputs(self) -> None:
        tree = self.dashboard_recent_tree
        if tree is None or not tree.winfo_exists():
            return
        tree.delete(*tree.get_children())
        self.dashboard_recent_files = [
            artifact
            for artifact in self.manifest.get("recent_files", [])
            if Path(artifact.get("path", "")).suffix.lower() != ".log"
        ][:5]
        for index, artifact in enumerate(self.dashboard_recent_files):
            tree.insert(
                "",
                "end",
                iid=f"dashboard-recent-{index}",
                values=(artifact.get("name", ""), artifact.get("revision", "")),
            )

    def open_dashboard_recent_file(self) -> None:
        def operation() -> None:
            self.require_job()
            selected = self.dashboard_recent_tree.selection()
            if not selected:
                raise JobError("Select a recent output first.")
            index = int(selected[0].rsplit("-", 1)[-1])
            path = Path(self.dashboard_recent_files[index]["path"])
            if not path.exists():
                raise JobError(f"The recorded file is no longer available: {path}")
            open_path(path)

        self.handle(operation)

    def open_job_location(self, section: str, key: str) -> None:
        def operation() -> None:
            self.require_job()
            path = Path(self.manifest[section][key])
            if not path.exists():
                raise JobError(f"The job folder is no longer available: {path}")
            open_path(path)

        self.handle(operation)

    def show_stage(self) -> None:
        if not self.manifest or not self.active_stage:
            return
        stage = self.selected_stage()
        guide = STAGE_GUIDANCE[stage]
        item = self.manifest["stages"][stage]
        checks = stage_checks(self.manifest, stage)
        status = item["status"].replace("_", " ").title()
        self.step_title.configure(text=item["label"])
        self.step_status.configure(
            text=status,
            style=(
                "StatusComplete.TLabel"
                if item["status"] == "complete"
                else "StatusAttention.TLabel"
                if item["status"] in {"needs_review", "warning"}
                else "Status.TLabel"
            ),
        )
        self.step_summary.configure(text=guide["action"])
        self.primary_step_button.configure(
            text="Run Again" if item["status"] == "complete" else "Start Step"
        )
        self.stage_detail_text = (
            f"{item['label']}\n\n"
            f"Status: {status}\n\n"
            f"Tool / macro\n{guide['tool']}\n\n"
            f"Changes made\n{guide['changes']}\n\n"
            "Workspace locations\n"
            f"Staging: {self.manifest['workspace']['staging']}\n"
            f"Logs: {self.manifest['workspace']['logs']}\n\n"
            "Warnings and overrides\n"
            "Sequence warnings can be continued after confirmation and are "
            "recorded with the Windows username. Missing inputs or tools block "
            "launch.\n\n"
            "All readiness checks\n"
            + "\n\n".join(
                f"{c.level.upper()} — {c.message}\nResolution: "
                f"{c.correction or 'No action required.'}"
                for c in checks
            )
        )

    def show_job_details(self) -> None:
        if not self.manifest:
            self.handle(lambda: self.require_job())
            return

        job = self.manifest["job"]
        paths = self.manifest["paths"]
        warnings = dashboard_warnings(self.manifest)
        complete = sum(
            self.manifest["stages"][key]["status"] == "complete"
            for key, _label in STAGES
        )
        details = (
            f"Job: {job['number']} — {job.get('name') or job['number']}\n"
            f"Revision: {job['revision']}\n"
            f"Progress: {complete} of {len(STAGES)} assistant steps complete\n\n"
            "Production folders\n"
            f"Engineering Process: {paths['engineering_root']}\n"
            f"3D Model: {paths['model_3d']}\n"
            f"Cut Files: {paths['cut_files']}\n"
            f"Part Checking: {paths.get('part_checking') or 'Not selected'}\n"
            f"Nesting: {paths.get('nesting') or 'Not selected'}\n\n"
            "Assistant workspace\n"
            f"Staging: {self.manifest['workspace']['staging']}\n"
            f"Logs: {self.manifest['workspace']['logs']}\n\n"
            "Items needing attention\n"
            + ("\n".join(f"• {warning}" for warning in warnings) if warnings else "None")
        )

        window = tk.Toplevel(self)
        window.title("Job Details")
        window.geometry("900x640")
        window.minsize(700, 500)
        window.transient(self)

        body = ttk.Frame(window, padding=14)
        body.pack(fill="both", expand=True)
        info = tk.Text(
            body,
            wrap="word",
            height=18,
            padx=10,
            pady=10,
            relief="flat",
            background=DARK_INPUT,
            foreground=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            font=("Segoe UI", 10),
        )
        info.insert("1.0", details)
        info.configure(state="disabled")
        info.pack(fill="both", expand=True)

        ttk.Label(
            body,
            text="Recently recorded files",
            font=("Segoe UI Semibold", 10),
        ).pack(anchor="w", pady=(12, 5))
        recent_tree = ttk.Treeview(
            body,
            columns=("name", "revision", "location"),
            show="headings",
            height=5,
        )
        recent_tree.heading("name", text="File")
        recent_tree.heading("revision", text="Rev")
        recent_tree.heading("location", text="Location")
        recent_tree.column("name", width=260, anchor="w")
        recent_tree.column("revision", width=55, anchor="center")
        recent_tree.column("location", width=480, anchor="w")
        for index, artifact in enumerate(self.manifest.get("recent_files", [])[:10]):
            recent_tree.insert(
                "",
                "end",
                iid=f"recent-{index}",
                values=(
                    artifact.get("name", ""),
                    artifact.get("revision", ""),
                    artifact.get("path", ""),
                ),
            )
        recent_tree.pack(fill="x")
        recent_tree.bind("<Double-1>", lambda _event: self.open_recent_file())
        self.recent_tree = recent_tree

        controls = ttk.Frame(body)
        controls.pack(fill="x", pady=(10, 0))
        ttk.Button(
            controls,
            text="Open Selected File",
            command=self.open_recent_file,
        ).pack(side="right", padx=(6, 0))
        ttk.Button(controls, text="Close", command=window.destroy).pack(side="right")

    def show_technical_details(self) -> None:
        if not self.manifest or not self.active_stage:
            self.handle(lambda: self.require_job())
            return
        window = tk.Toplevel(self)
        window.title("Selected Step — Technical Details")
        window.geometry("860x620")
        window.minsize(650, 420)
        window.transient(self)
        text = tk.Text(
            window,
            wrap="word",
            padx=12,
            pady=12,
            background=DARK_INPUT,
            foreground=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
        )
        scrollbar = ttk.Scrollbar(window, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.insert("1.0", self.stage_detail_text)
        text.configure(state="disabled")
        text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

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
                raise JobError(f"Unknown assistant step: {stage}")
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
        log = Path(self.manifest["workspace"]["logs"]) / "bom-converter.log"
        record_event(
            self.manifest,
            "external_process_requested",
            stage="bom",
            command=command,
            output=output,
            log=str(log),
        )
        save_manifest(self.manifest, self.manifest_path)
        handle = log.open("a", encoding="utf-8")
        handle.write(f"Windows command: {subprocess.list2cmdline(command)}\n")
        handle.write(
            f"Arguments: {command!r}\nOutput: {output}\nStage: bom\n\n"
        )
        handle.flush()
        # A remembered repository can be unavailable (especially a disconnected
        # network drive).  Passing that stale path as cwd raises WinError 267
        # before the packaged converter can even start.  The command already
        # contains absolute executable/script paths, so it is safe to inherit the
        # assistant's working directory when the preferred repository is absent.
        try:
            process = subprocess.Popen(
                command,
                cwd=existing_working_directory(self.repo),
                stdout=handle,
                stderr=subprocess.STDOUT,
            )
        except Exception:
            handle.close()
            raise
        record_event(
            self.manifest,
            "external_process_launched",
            stage="bom",
            command=command,
            pid=process.pid,
            output=output,
            log=str(log),
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
        self.running_processes[process.pid] = {
            "stage": "bom",
            "job_number": self.manifest["job"]["number"],
            "manifest_path": str(self.manifest_path),
            "result": str(output),
            "log": str(log),
        }
        self.update_running_summary()
        self._poll_bom_process(
            process,
            handle,
            log,
            Path(output),
            self.manifest,
            self.manifest_path,
        )

    def _poll_bom_process(
        self,
        process,
        handle,
        log: Path,
        output: Path,
        process_manifest: dict,
        process_manifest_path: Path,
    ) -> None:
        exit_code = process.poll()
        if exit_code is None:
            self.after(
                500,
                self._poll_bom_process,
                process,
                handle,
                log,
                output,
                process_manifest,
                process_manifest_path,
            )
            return
        handle.write(f"\nProcess exit code: {exit_code}\n")
        handle.close()
        self.running_processes.pop(process.pid, None)
        self.update_running_summary()
        record_event(
            process_manifest,
            "external_process_finished",
            stage="bom",
            pid=process.pid,
            exit_code=exit_code,
            output=str(output),
            log=str(log),
        )
        succeeded = exit_code == 0 and output.is_file()
        if succeeded:
            mark_needs_review(
                process_manifest,
                "bom",
                "Parts List created. The workbook was opened for required review.",
                [output],
            )
        else:
            item = process_manifest["stages"]["bom"]
            item["status"] = "warning"
            item["notes"] = (
                "The BOM converter closed without creating the selected Parts "
                f"List output: {output}. Review the converter log: {log}"
            )
            record_artifact(process_manifest, "bom", log)
        save_manifest(process_manifest, process_manifest_path)
        if self.manifest_path == process_manifest_path:
            self.manifest = process_manifest
            self.refresh()
            self.show_stage()
        self.post_background_notice(
            "BOM conversion finished" if succeeded else "BOM conversion incomplete",
            (
                "The Parts List is ready for review."
                if succeeded
                else "No completed Parts List was found."
            ),
            level="info" if succeeded else "warning",
            path=output if succeeded else log,
        )
        if succeeded:
            self.open_review_and_offer_completion(
                "bom", output, process_manifest, process_manifest_path
            )

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
            autocad_console=self.settings.get("autocad_console") or None,
            tool_executable=self.bundled_tool("Engineering DXF Orchestrator.exe"),
            workers=self.settings.get("autocad_workers", 2),
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
        save_manifest(process_manifest, process_manifest_path)
        if self.manifest_path == process_manifest_path:
            self.manifest = process_manifest
            self.refresh()
        outcome = "finished and needs review" if exit_code == 0 else "failed"
        self.post_background_notice(
            "DXF automation finished",
            f"Job {process_manifest['job']['number']} {outcome}.",
            level="warning" if exit_code else "info",
            path=run if exit_code == 0 else log,
        )
        if exit_code == 0:
            self.open_review_and_offer_completion(
                "dxf", run, process_manifest, process_manifest_path
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
        canvas = tk.Canvas(frame, background=DARK_BG, highlightthickness=0)
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
        if any(
            item.get("stage") == "comparison"
            for item in self.running_processes.values()
        ):
            raise JobError(
                "A production comparison is already running. Wait for its "
                "completion notification before starting another."
            )

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
        record_event(
            self.manifest,
            "external_process_requested",
            stage="comparison",
            command=command,
            output=str(output),
            log=str(log),
        )
        save_manifest(self.manifest, self.manifest_path)

        handle = log.open("a", encoding="utf-8")
        handle.write(f"Windows command: {subprocess.list2cmdline(command)}\n")
        handle.write(f"Arguments: {command!r}\nOutput: {output}\nStage: comparison\n\n")
        handle.flush()
        try:
            process = subprocess.Popen(
                command,
                cwd=existing_working_directory(self.repo),
                stdout=handle,
                stderr=subprocess.STDOUT,
            )
        except Exception:
            handle.close()
            raise

        record_event(
            self.manifest,
            "external_process_launched",
            stage="comparison",
            pid=process.pid,
            output=str(output),
            log=str(log),
        )
        self.running_processes[process.pid] = {
            "stage": "comparison",
            "job_number": self.manifest["job"]["number"],
            "manifest_path": str(self.manifest_path),
            "log": str(log),
        }
        self.update_running_summary()
        save_manifest(self.manifest, self.manifest_path)
        self._poll_comparison_process(
            process,
            handle,
            log,
            output,
            self.manifest,
            self.manifest_path,
        )

    def _poll_comparison_process(
        self,
        process,
        handle,
        log: Path,
        output: Path,
        process_manifest: dict,
        process_manifest_path: Path,
    ) -> None:
        exit_code = process.poll()
        if exit_code is None:
            self.after(
                500,
                self._poll_comparison_process,
                process,
                handle,
                log,
                output,
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
            stage="comparison",
            pid=process.pid,
            exit_code=exit_code,
            output=str(output),
            log=str(log),
        )

        title = "Production comparison finished"
        if exit_code:
            item = process_manifest["stages"]["comparison"]
            item["status"] = "warning"
            item["notes"] = (
                f"Comparison exited with code {exit_code}. Review {log}"
            )
            record_artifact(process_manifest, "comparison", log)
            title = "Production comparison failed"
        else:
            comparison = parse_comparison_summary(output)
            process_manifest["comparison"] = comparison
            if comparison["status"] == "not_available":
                item = process_manifest["stages"]["comparison"]
                item["status"] = "warning"
                item["notes"] = comparison["message"] + f" Review {log}"
                record_artifact(process_manifest, "comparison", log)
                title = "Production comparison incomplete"
            else:
                mark_needs_review(
                    process_manifest,
                    "comparison",
                    f"{comparison['message']}: {comparison['errors']} errors, "
                    f"{comparison['warnings']} warnings.",
                    [log],
                )

        save_manifest(process_manifest, process_manifest_path)
        if self.manifest_path == process_manifest_path:
            self.manifest = process_manifest
            self.refresh()
            self.show_stage()

        notes = process_manifest["stages"]["comparison"].get("notes", "")
        message = f"Job {process_manifest['job']['number']}\n\n{notes}\n\nLog: {log}"
        self.post_background_notice(
            title,
            message.replace("\n\n", " "),
            level=(
                "warning" if title.endswith(("failed", "incomplete")) else "info"
            ),
            path=output if output.exists() else log,
        )
        if exit_code == 0 and process_manifest["stages"]["comparison"]["status"] == "needs_review":
            workbook = output / "production_part_comparison.xlsx"
            review_target = workbook if workbook.is_file() else output
            self.open_review_and_offer_completion(
                "comparison",
                review_target,
                process_manifest,
                process_manifest_path,
            )

    def launch_solidworks_macro(self, stage: str) -> None:
        macro = self.repo / (
            "solidworks/cad-batch-converter/Main.RunBatch.swp"
            if stage == "plate_model"
            else "solidworks/utilities/(MOD)2(SECONDARY).swp"
        )
        if not macro.is_file():
            raise JobError(f"Macro was not found: {macro}")
        runner = self.repo / "solidworks/cad-batch-converter/run_macro.py"
        if not runner.is_file():
            raise JobError(f"SolidWorks runner was not found: {runner}")
        if stage != "plate_model":
            if self._accept_warnings(stage):
                self._start_solidworks_runner(stage, macro, runner)
            return

        # Completed DXF outputs may already have been moved out of the
        # assistant workspace. Production Cut Files is therefore the primary
        # location; the orchestrator run is only a fallback.
        cut_files = Path(self.manifest["paths"]["cut_files"])
        prepared_root = (
            cut_files
            if cut_files.is_dir()
            else Path(
                self.manifest["stages"]["dxf"].get(
                    "workspace", self.manifest["workspace"]["working"]
                )
            )
        )
        source_value = filedialog.askdirectory(
            title="Select the reviewed folder containing prepared DWGs",
            initialdir=prepared_root,
            parent=self,
        )
        if not source_value:
            return
        source = Path(source_value)
        drawings = sorted(
            (
                path
                for path in source.iterdir()
                if path.is_file() and path.suffix.casefold() == ".dwg"
            ),
            key=lambda path: path.name.casefold(),
        )
        if not drawings:
            raise JobError("The selected plate-input folder contains no DWG files.")
        thickness = simpledialog.askfloat(
            "Plate thickness",
            "Enter the extrusion thickness in inches for every DWG in this folder:",
            initialvalue=inferred_plate_thickness(source),
            minvalue=0.001,
            maxvalue=20.0,
            parent=self,
        )
        if thickness is None:
            return
        if not self._accept_warnings(stage):
            return

        run_name = safe_name(source.name)
        output = (
            Path(self.manifest["workspace"]["staging"])
            / "SolidWorks Parts"
            / run_name
        )
        output.mkdir(parents=True, exist_ok=True)

        local_base = Path(
            os.environ.get(
                "LOCALAPPDATA",
                Path.home() / "AppData" / "Local",
            )
        ) / "EngineeringJobAssistant" / "PlateBatches"
        local_run = (
            local_base
            / f"{safe_name(self.manifest['job']['number'])}-{run_name}-"
            f"{uuid.uuid4().hex[:8]}"
        )
        local_source = local_run / "Source DWGs"
        filtered = local_run / "Filtered DWGs"
        automation_output = local_run / "SolidWorks Parts"
        for folder in (local_source, filtered, automation_output):
            folder.mkdir(parents=True, exist_ok=False)
        (local_run / ".engineering-job-assistant-plate-batch").write_text(
            "Temporary local plate batch; safe for the assistant to remove after "
            "verified publication.\n",
            encoding="utf-8",
        )
        copied = 0
        try:
            for drawing in drawings:
                shutil.copy2(drawing, local_source / drawing.name)
                copied += 1
        except OSError as exc:
            raise JobError(
                f"Could not stage DWGs on the local drive at {local_source}: {exc}"
            ) from exc
        if copied == 0:
            raise JobError("No DWGs were copied into the local speed workspace.")

        values = {
            "MACROS_SOURCE_FOLDER": str(local_source),
            "MACROS_FILTERED_FOLDER": str(filtered),
            "MACROS_OUTPUT_FOLDER": str(automation_output),
            "MACROS_EXTRUDE_DEPTH_METERS": format(thickness * 0.0254, ".12g"),
        }
        os.environ.update(values)
        if sys.platform == "win32":
            import winreg

            key_path = (
                r"Software\VB and VBA Program Settings\EngineeringMacros\CadBatch"
            )
            names = {
                "MACROS_SOURCE_FOLDER": "SourceFolder",
                "MACROS_FILTERED_FOLDER": "FilteredFolder",
                "MACROS_OUTPUT_FOLDER": "OutputFolder",
                "MACROS_EXTRUDE_DEPTH_METERS": "ExtrudeDepthMeters",
            }
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                for env_name, value in values.items():
                    winreg.SetValueEx(
                        key, names[env_name], 0, winreg.REG_SZ, value
                    )

        record_event(
            self.manifest,
            "cad_batch_configured",
            stage=stage,
            macro=str(macro),
            source=str(source),
            local_source=str(local_source),
            local_output=str(automation_output),
            output=str(output),
            local_dwgs=copied,
            solidworks_mode="single_persistent_local_session",
            solidworks_workers=1,
            thickness_inches=thickness,
        )
        self._start_solidworks_runner(
            stage,
            macro,
            runner,
            source=source,
            output=output,
            automation_source=local_source,
            automation_output=automation_output,
            cleanup_workspace=local_run,
            thickness=thickness,
        )

    def _start_solidworks_runner(
        self,
        stage: str,
        macro: Path,
        runner: Path,
        *,
        source: Path | None = None,
        output: Path | None = None,
        automation_source: Path | None = None,
        automation_output: Path | None = None,
        cleanup_workspace: Path | None = None,
        thickness: float | None = None,
    ) -> None:
        """Run one verified macro batch through one persistent SolidWorks session."""

        for item in self.running_processes.values():
            if item.get("stage") in {"plate_model", "autobom"}:
                raise JobError(
                    "A SolidWorks automation run is already active. Wait for it "
                    "to finish before starting another SolidWorks step."
                )
        command = [
            sys.executable,
            "-u",
            str(runner),
            "--macro",
            str(macro),
        ]
        solidworks = self.settings.get("solidworks_executable", "").strip()
        if solidworks:
            if not Path(solidworks).is_file():
                raise JobError(
                    f"Configured SolidWorks executable was not found: {solidworks}"
                )
            command.extend(["--solidworks-executable", solidworks])
        if stage == "plate_model" and output is not None:
            try:
                runner_source = runner.read_text(encoding="utf-8")
            except OSError:
                runner_source = ""
            normalize_output = automation_output or output
            if "--normalize-output" in runner_source:
                command.extend(["--normalize-output", str(normalize_output)])
            if automation_output is not None and "--publish-output" in runner_source:
                command.extend(["--publish-output", str(output)])
            if cleanup_workspace is not None and "--cleanup-workspace" in runner_source:
                command.extend(["--cleanup-workspace", str(cleanup_workspace)])
        suffix = safe_name(source.name if source else stage)
        log = Path(self.manifest["workspace"]["logs"]) / f"solidworks-{suffix}.log"
        handle = log.open("a", encoding="utf-8")
        handle.write(f"Windows command: {subprocess.list2cmdline(command)}\n")
        if source:
            handle.write(f"Source: {source}\n")
        if output:
            handle.write(f"Published output: {output}\n")
        if automation_output:
            handle.write(f"Local automation output: {automation_output}\n")
        if cleanup_workspace:
            handle.write(f"Local speed workspace: {cleanup_workspace}\n")
        handle.write("SolidWorks mode: one persistent local session\n")
        if thickness is not None:
            handle.write(f"Thickness: {thickness:g} in\n")
        handle.flush()
        try:
            process = subprocess.Popen(
                command,
                cwd=existing_working_directory(self.repo),
                stdout=handle,
                stderr=subprocess.STDOUT,
            )
        except Exception:
            handle.close()
            raise
        record_event(
            self.manifest,
            "external_process_launched",
            stage=stage,
            pid=process.pid,
            macro=str(macro),
            source=str(source) if source else "",
            output=str(output) if output else "",
            solidworks_mode="single_persistent_local_session",
            solidworks_workers=1,
            log=str(log),
        )
        self.running_processes[process.pid] = {
            "stage": stage,
            "job_number": self.manifest["job"]["number"],
            "manifest_path": str(self.manifest_path),
            "log": str(log),
        }
        self.update_running_summary()
        save_manifest(self.manifest, self.manifest_path)
        self._poll_solidworks_process(
            process,
            handle,
            stage,
            log,
            output,
            self.manifest,
            self.manifest_path,
        )

    def _poll_solidworks_process(
        self,
        process,
        handle,
        stage: str,
        log: Path,
        output: Path | None,
        process_manifest: dict,
        process_manifest_path: Path,
    ) -> None:
        exit_code = process.poll()
        if exit_code is None:
            self.after(
                500,
                self._poll_solidworks_process,
                process,
                handle,
                stage,
                log,
                output,
                process_manifest,
                process_manifest_path,
            )
            return
        rename_error = ""
        renamed_parts: list[Path] = []
        runner_handles_normalization = False
        if stage == "plate_model":
            runner = self.repo / "solidworks/cad-batch-converter/run_macro.py"
            try:
                runner_handles_normalization = (
                    "--normalize-output" in runner.read_text(encoding="utf-8")
                )
            except OSError:
                pass
        if (
            exit_code == 0
            and stage == "plate_model"
            and output is not None
            and not runner_handles_normalization
        ):
            try:
                renamed_parts = normalize_plate_model_filenames(output)
                handle.write(
                    f"\nNormalized {len(renamed_parts)} SolidWorks part filename(s).\n"
                )
                for part in renamed_parts:
                    handle.write(f"  {part.name}\n")
                record_event(
                    process_manifest,
                    "plate_model_filenames_normalized",
                    output=str(output),
                    renamed=[part.name for part in renamed_parts],
                )
            except JobError as exc:
                rename_error = str(exc)
                handle.write(f"\nFilename normalization failed: {rename_error}\n")
                record_event(
                    process_manifest,
                    "plate_model_filename_normalization_failed",
                    output=str(output),
                    error=rename_error,
                )
        handle.write(f"\nExit code: {exit_code}\n")
        handle.close()
        self.running_processes.pop(process.pid, None)
        self.update_running_summary()
        record_event(
            process_manifest,
            "external_process_finished",
            stage=stage,
            pid=process.pid,
            exit_code=exit_code,
            output=str(output) if output else "",
            log=str(log),
        )
        review_target = output or Path(process_manifest["paths"]["model_3d"])
        successful = exit_code == 0 and not rename_error
        if successful:
            note = (
                f"SolidWorks macro finished and normalized {len(renamed_parts)} "
                "part filename(s). The result location was opened for required review."
                if stage == "plate_model"
                else "SolidWorks macro finished. The result location was opened for "
                "required review."
            )
            mark_needs_review(
                process_manifest,
                stage,
                note,
                [review_target, log],
            )
        else:
            item = process_manifest["stages"][stage]
            item["status"] = "warning"
            item["notes"] = (
                f"SolidWorks filename cleanup failed: {rename_error}. Review {log}"
                if rename_error
                else f"SolidWorks automation exited with code {exit_code}. Review {log}"
            )
            record_artifact(process_manifest, stage, log)
        save_manifest(process_manifest, process_manifest_path)
        if self.manifest_path == process_manifest_path:
            self.manifest = process_manifest
            self.refresh()
            self.show_stage()
        self.post_background_notice(
            "SolidWorks automation finished" if successful else "SolidWorks automation failed",
            (
                "Review the generated models."
                if successful
                else (
                    f"{rename_error} Review {log.name}."
                    if rename_error
                    else f"Review {log.name}."
                )
            ),
            level="info" if successful else "warning",
            path=review_target if successful else log,
        )
        if not successful or self.manifest_path != process_manifest_path:
            return
        self.open_review_and_offer_completion(
            stage,
            review_target,
            process_manifest,
            process_manifest_path,
            ask_to_complete=stage != "plate_model",
        )
        if stage != "plate_model":
            return
        if messagebox.askyesno(
            "Process another thickness folder?",
            "This folder has finished and its output is open for review.\n\n"
            "Select Yes to choose the next plate folder. Select No when all "
            "plate folders are finished.",
            parent=self,
        ):
            self.after(0, lambda: self.launch_solidworks_macro("plate_model"))
        else:
            self.open_review_and_offer_completion(
                stage,
                review_target,
                process_manifest,
                process_manifest_path,
            )

    def open_stage_folder(self) -> None:
        def operation():
            self.require_job()
            stage = self.selected_stage()
            key = (
                "model_3d"
                if stage in {"plate_model", "autobom"}
                else "nesting"
                if stage == "comparison"
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

    def move_outputs(self) -> None:
        def operation() -> None:
            self.require_job()
            plan = plan_completed_outputs(self.manifest)
            if not plan:
                complete = [
                    label
                    for key, label in STAGES
                    if key in {"dxf", "plate_model"}
                    and self.manifest["stages"][key]["status"] == "complete"
                ]
                if complete:
                    raise JobError(
                        "No completed cut files or plate models remain in the "
                        "assistant workspace. They may already have been moved."
                    )
                raise JobError(
                    "Complete and positively review the Cut Files or Plate Models "
                    "step before moving its outputs to production."
                )
            self.show_output_move_table(plan)

        self.handle(operation)

    def show_output_move_table(self, plan: list[OutputMoveItem]) -> None:
        """Show one review screen for automatic, recoverable production moves."""

        window = tk.Toplevel(self)
        window.title("Move Completed Outputs")
        window.geometry("1220x650")
        window.minsize(900, 480)
        window.transient(self)
        window.grab_set()

        cut_count = sum(item.category == "cut_file" for item in plan)
        plate_count = sum(item.category == "plate_model" for item in plan)
        conflict_count = sum(item.conflict for item in plan)
        ttk.Label(
            window,
            text=(
                f"Ready: {cut_count} cut file(s) and {plate_count} plate model(s). "
                f"{conflict_count} destination conflict(s) will be backed up "
                "automatically before replacement. Source files are removed only "
                "after a verified copy reaches production."
            ),
            padding=10,
            wraplength=1160,
        ).pack(fill="x")

        columns = ("type", "run", "source", "destination", "conflict", "result")
        tree_frame = ttk.Frame(window, padding=(10, 0))
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        widths = {
            "type": 100,
            "run": 130,
            "source": 290,
            "destination": 350,
            "conflict": 90,
            "result": 150,
        }
        for column in columns:
            tree.heading(column, text=column.title())
            tree.column(column, width=widths[column], anchor="w")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for index, item in enumerate(plan):
            tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    "Cut file" if item.category == "cut_file" else "Plate model",
                    item.run_name,
                    str(item.source),
                    str(item.destination),
                    "Back up + replace" if item.conflict else "No",
                    "Ready",
                ),
                tags=("conflict" if item.conflict else "new",),
            )
        tree.tag_configure("conflict", foreground="#ffb66d")
        tree.tag_configure("new", foreground="#79dfa1")
        tree.tag_configure("failed", foreground="#ff8d95")
        tree.tag_configure("finished", foreground="#79dfa1")

        detail = tk.StringVar(
            value="Review the destinations, then move all completed outputs."
        )
        ttk.Label(window, textvariable=detail, padding=(10, 7)).pack(fill="x")
        controls = ttk.Frame(window, padding=(10, 0, 10, 10))
        controls.pack(fill="x")
        execute_button = ttk.Button(
            controls,
            text="Move All Completed Outputs",
            style="Primary.TButton",
        )
        execute_button.pack(side="right", padx=3)
        ttk.Button(controls, text="Cancel", command=window.destroy).pack(
            side="right", padx=3
        )

        def execute() -> None:
            if not messagebox.askyesno(
                "Move completed outputs?",
                f"Move all {len(plan)} listed file(s) to the selected production "
                "folders? Existing files will be backed up before replacement.",
                parent=window,
            ):
                return
            execute_button.configure(state="disabled")
            detail.set("Moving outputs and verifying each copy…")
            window.update_idletasks()
            results = move_completed_outputs(self.manifest, plan)
            save_manifest(self.manifest, self.manifest_path)
            for index, result in enumerate(results):
                values = list(tree.item(str(index), "values"))
                values[5] = result["status"].replace("_", " ").title()
                if result.get("error"):
                    values[5] += f": {result['error']}"
                tree.item(
                    str(index),
                    values=values,
                    tags=(
                        "failed" if result["status"] == "failed" else "finished",
                    ),
                )
            counts = {
                status: sum(result["status"] == status for result in results)
                for status in ("moved", "replaced", "already_current", "failed")
            }
            report_value = next(
                (result.get("report") for result in results if result.get("report")),
                "",
            )
            detail.set(
                f"Finished: {counts['moved']} moved, {counts['replaced']} replaced, "
                f"{counts['already_current']} already current, "
                f"{counts['failed']} failed."
            )
            self.post_background_notice(
                "Completed output move finished",
                detail.get(),
                level="warning" if counts["failed"] else "info",
                path=Path(report_value) if report_value else None,
            )
            self.refresh()

        execute_button.configure(command=lambda: self.handle(execute))

    def edit_settings(self) -> None:
        window = tk.Toplevel(self)
        window.title("Per-user settings (not stored in the repository)")
        window.geometry("820x390")
        window.transient(self)
        entries = {}
        labels = (
            ("macros_repo", "Shared Macros repository"),
            ("parts_list_template", "Standard Parts List template"),
            ("default_jobs_parent", "Default jobs parent"),
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
        workers_row = len(labels)
        ttk.Label(window, text="Parallel AutoCAD processes").grid(
            row=workers_row, column=0, sticky="w", padx=8, pady=7
        )
        workers = tk.IntVar(value=self.settings.get("autocad_workers", 2))
        ttk.Spinbox(
            window,
            from_=1,
            to=4,
            textvariable=workers,
            width=8,
            state="readonly",
        ).grid(row=workers_row, column=1, sticky="w", padx=5)
        ttk.Label(
            window,
            text="2 recommended. Each process converts one DXF; bevel files get (B).",
        ).grid(row=workers_row + 1, column=1, sticky="w", padx=5)
        window.columnconfigure(1, weight=1)

        def save():
            self.settings.update(
                {key: var.get().strip() for key, var in entries.items()}
            )
            self.settings["autocad_workers"] = workers.get()
            save_settings(self.settings)
            window.destroy()

        ttk.Button(window, text="Save", command=lambda: self.handle(save)).grid(
            row=workers_row + 2, column=2, pady=12
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
            if self.recent_tree is None or not self.recent_tree.winfo_exists():
                raise JobError("Open Job Details to select a recently recorded file.")
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
