from __future__ import annotations

import threading
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from . import __version__
from .engine import convert_epls, load_metadata


class ConverterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"EPL-to-Parts-List Converter v{__version__}")
        self.geometry("920x820")
        self.minsize(760, 640)
        self.epl_files: list[str] = []
        self.bop_files: list[str] = []
        self.metadata_path = ctk.StringVar()
        self.plate_template_path = ctk.StringVar()
        self.output_dir = ctk.StringVar(value=str(Path.home() / "Documents"))
        self.output_name = ctk.StringVar(value="Parts List")
        self._build()

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(
            self,
            text="EPL-to-Parts-List Converter",
            font=ctk.CTkFont(size=25, weight="bold"),
        )
        title.grid(row=0, column=0, padx=24, pady=(18, 4), sticky="w")
        subtitle = ctk.CTkLabel(
            self,
            text=(
                f"Version {__version__}  •  Select both EPL and BOP/BOM files. "
                "The BOP/BOM determines which parts are exported."
            ),
            text_color=("gray35", "gray70"),
        )
        subtitle.grid(row=1, column=0, padx=24, pady=(0, 10), sticky="w")

        content = ctk.CTkScrollableFrame(self, fg_color="transparent")
        content.grid(row=2, column=0, padx=18, pady=(0, 4), sticky="nsew")
        content.grid_columnconfigure(0, weight=1)

        self.file_box, self.epl_count_label = self._file_panel(
            content,
            row=0,
            title="1. EPL WORKBOOKS",
            description="Engineering source data. Select one or more EPL workbooks.",
            button_text="Choose EPL File(s)",
            command=self._select_epls,
        )
        self.bop_box, self.bop_count_label = self._file_panel(
            content,
            row=1,
            title="2. BOP / BOM WORKBOOKS — REQUIRED",
            description=(
                "Scope authority. Only EPL parts listed in the selected BOP/BOM are exported."
            ),
            button_text="Choose Required BOP/BOM File(s)",
            command=self._select_bops,
            required=True,
        )

        options = ctk.CTkFrame(content)
        options.grid(row=2, column=0, padx=6, pady=7, sticky="ew")
        options.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            options,
            text="3. OUTPUT OPTIONS",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, padx=14, pady=(12, 4), sticky="w")
        self._row(options, 1, "Optional metadata", self.metadata_path, self._select_metadata)
        self._row(
            options,
            2,
            "Plates template",
            self.plate_template_path,
            self._select_plate_template,
        )
        self._row(options, 3, "Output folder", self.output_dir, self._select_output_dir)
        ctk.CTkLabel(options, text="Output name", width=125, anchor="w").grid(
            row=4, column=0, padx=(14, 8), pady=(8, 14)
        )
        ctk.CTkEntry(options, textvariable=self.output_name).grid(
            row=4, column=1, columnspan=2, padx=(0, 14), pady=(8, 14), sticky="ew"
        )

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=3, column=0, padx=24, pady=(6, 4), sticky="ew")
        actions.grid_columnconfigure(0, weight=1)
        self.convert_button = ctk.CTkButton(
            actions,
            text="Convert BOP-Scoped Parts",
            height=44,
            width=220,
            command=self._start_conversion,
        )
        self.convert_button.grid(row=0, column=1, sticky="e")
        self.progress = ctk.CTkProgressBar(actions, mode="indeterminate", width=260)
        self.progress.grid(row=0, column=0, sticky="w")
        self.progress.grid_remove()
        self.result_label = ctk.CTkLabel(self, text="", justify="left", anchor="w")
        self.result_label.grid(row=4, column=0, padx=24, pady=(4, 14), sticky="ew")

    def _file_panel(
        self,
        parent: ctk.CTkScrollableFrame,
        row: int,
        title: str,
        description: str,
        button_text: str,
        command,
        required: bool = False,
    ) -> tuple[ctk.CTkTextbox, ctk.CTkLabel]:
        panel = ctk.CTkFrame(
            parent,
            border_width=2 if required else 0,
            border_color="#D97706" if required else None,
        )
        panel.grid(row=row, column=0, padx=6, pady=7, sticky="ew")
        panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            panel,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#D97706" if required else ("gray10", "gray90"),
        ).grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")
        ctk.CTkLabel(
            panel,
            text=description,
            text_color=("gray35", "gray70"),
        ).grid(row=1, column=0, columnspan=2, padx=14, pady=(0, 8), sticky="w")
        button = ctk.CTkButton(panel, text=button_text, command=command, width=235)
        button.grid(row=2, column=0, padx=14, pady=(0, 8), sticky="w")
        count = ctk.CTkLabel(panel, text="No files selected", anchor="e")
        count.grid(row=2, column=1, padx=14, pady=(0, 8), sticky="e")
        box = ctk.CTkTextbox(panel, height=72)
        box.grid(row=3, column=0, columnspan=2, padx=14, pady=(0, 14), sticky="ew")
        box.insert("1.0", "Nothing selected")
        box.configure(state="disabled")
        return box, count

    def _row(self, parent: ctk.CTkFrame, row: int, label: str, variable: ctk.StringVar, command) -> None:
        ctk.CTkLabel(parent, text=label, width=125, anchor="w").grid(
            row=row, column=0, padx=(14, 8), pady=8
        )
        ctk.CTkEntry(parent, textvariable=variable).grid(
            row=row, column=1, padx=(0, 8), pady=8, sticky="ew"
        )
        ctk.CTkButton(parent, text="Browse…", width=90, command=command).grid(
            row=row, column=2, padx=(0, 14), pady=8
        )

    def _select_epls(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Select EPL workbooks",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if selected:
            self.epl_files = list(selected)
            self.file_box.configure(state="normal")
            self.file_box.delete("1.0", "end")
            self.file_box.insert("1.0", "\n".join(Path(path).name for path in selected))
            self.file_box.configure(state="disabled")
            self.epl_count_label.configure(text=f"{len(selected)} EPL file(s) selected")

    def _select_metadata(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select optional part metadata",
            filetypes=[("Metadata", "*.json *.csv"), ("All files", "*.*")],
        )
        if selected:
            self.metadata_path.set(selected)

    def _select_bops(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Select BOP workbooks",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if selected:
            self.bop_files = list(selected)
            self.bop_box.configure(state="normal")
            self.bop_box.delete("1.0", "end")
            self.bop_box.insert("1.0", "\n".join(Path(path).name for path in selected))
            self.bop_box.configure(state="disabled")
            self.bop_count_label.configure(text=f"{len(selected)} BOP/BOM file(s) selected")

    def _select_plate_template(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select optional Plates template",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.plate_template_path.set(selected)

    def _select_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="Choose output folder")
        if selected:
            self.output_dir.set(selected)

    def _start_conversion(self) -> None:
        if not self.epl_files:
            messagebox.showwarning("No EPL files", "Select at least one EPL workbook.")
            return
        if not self.bop_files:
            messagebox.showwarning(
                "No BOP files",
                "Select at least one BOP workbook. The BOP defines which parts are in scope.",
            )
            return
        self.convert_button.configure(state="disabled")
        self.progress.grid()
        self.progress.start()
        self.result_label.configure(text="Converting…")
        threading.Thread(target=self._convert, daemon=True).start()

    def _convert(self) -> None:
        try:
            metadata = load_metadata(self.metadata_path.get()) if self.metadata_path.get() else None
            result = convert_epls(
                self.epl_files,
                self.output_dir.get(),
                self.output_name.get(),
                metadata,
                bop_paths=self.bop_files,
                plate_template_path=self.plate_template_path.get() or None,
            )
            summary = result.to_dict()["summary"]
            message = (
                f"Done: {summary['plates_exported']} plates and "
                f"{summary['shapes_exported']} shapes exported.\n"
                f"{summary['assemblies_omitted']} assemblies omitted; "
                f"{summary['epl_rows_out_of_scope']} EPL row(s) outside BOP scope; "
                f"{summary['unclassified_items']} unclassified item(s).\n"
                f"Outputs: {result.plates_path.parent}"
            )
            self.after(0, partial(self._finish, message, None))
        except Exception as exc:
            # Exception variables are cleared when the except block exits.
            # Capture the text now so Tkinter can safely run this callback later.
            self.after(0, partial(self._finish, "", str(exc)))

    def _finish(self, message: str, error: str | None) -> None:
        self.progress.stop()
        self.progress.grid_remove()
        self.convert_button.configure(state="normal")
        if error:
            self.result_label.configure(text="Conversion did not finish.")
            messagebox.showerror("Conversion error", error)
        else:
            self.result_label.configure(text=message)
            messagebox.showinfo("Conversion complete", message)


def run_ui() -> None:
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = ConverterApp()
    app.mainloop()
