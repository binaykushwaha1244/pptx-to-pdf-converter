"""
PPTX to PDF Converter (Windows)
- Batch-convert PowerPoint .pptx files to PDF.
- Uses Microsoft PowerPoint when available, otherwise LibreOffice.
- Keeps files local on your computer.
"""

import os
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

APP_TITLE = "PPTX to PDF Converter"


def find_libreoffice():
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


def convert_with_powerpoint(source: Path, destination: Path):
    import win32com.client
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = None
    try:
        presentation = powerpoint.Presentations.Open(
            str(source.resolve()), ReadOnly=1, Untitled=0, WithWindow=0
        )
        presentation.SaveAs(str(destination.resolve()), 32)  # ppSaveAsPDF
    finally:
        if presentation is not None:
            presentation.Close()
        powerpoint.Quit()


def convert_with_libreoffice(source: Path, output_dir: Path, soffice: str):
    profile_dir = output_dir / ".lo_profile"
    profile_dir.mkdir(exist_ok=True)
    cmd = [
        soffice,
        "-env:UserInstallation=" + profile_dir.resolve().as_uri(),
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir.resolve()),
        str(source.resolve()),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    expected = output_dir / (source.stem + ".pdf")
    if result.returncode != 0 or not expected.exists():
        details = (result.stderr or result.stdout or "No error details returned.").strip()
        raise RuntimeError(details)
    return expected


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("720x510")
        self.root.minsize(620, 440)

        self.files = []
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents"))
        self.recursive = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Choose one or more PowerPoint files to begin.")

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PPTX to PDF Converter", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Convert one presentation or batch-convert a folder. Files stay on your computer.",
            wraplength=660
        ).pack(anchor="w", pady=(4, 14))

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Add PPTX files…", command=self.add_files).pack(side="left")
        ttk.Button(buttons, text="Add folder…", command=self.add_folder).pack(side="left", padx=8)
        ttk.Button(buttons, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(buttons, text="Clear list", command=self.clear_list).pack(side="right")

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True, pady=10)
        self.file_list = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=10)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scrollbar.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        output_frame = ttk.LabelFrame(outer, text="Save PDFs to", padding=10)
        output_frame.pack(fill="x", pady=(0, 10))
        ttk.Entry(output_frame, textvariable=self.output_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(output_frame, text="Browse…", command=self.choose_output).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))
        ttk.Label(outer, textvariable=self.status, wraplength=660).pack(anchor="w")
        self.convert_button = ttk.Button(outer, text="Convert to PDF", command=self.start_conversion)
        self.convert_button.pack(anchor="e", pady=(12, 0))

    def add_files(self):
        chosen = filedialog.askopenfilenames(
            title="Select PowerPoint files",
            filetypes=[("PowerPoint presentations", "*.pptx"), ("All files", "*.*")]
        )
        self._add_paths(chosen)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder containing PPTX files")
        if not folder:
            return
        path = Path(folder)
        found = sorted(path.rglob("*.pptx") if self.recursive.get() else path.glob("*.pptx"))
        self._add_paths([str(p) for p in found])
        if not found:
            messagebox.showinfo(APP_TITLE, "No .pptx files were found in that folder.")

    def _add_paths(self, paths):
        existing = {str(Path(p).resolve()) for p in self.files}
        for raw in paths:
            p = Path(raw)
            if p.is_file() and p.suffix.lower() == ".pptx":
                resolved = str(p.resolve())
                if resolved not in existing:
                    self.files.append(resolved)
                    self.file_list.insert("end", resolved)
                    existing.add(resolved)
        self.status.set(f"{len(self.files)} PowerPoint file(s) queued.")

    def remove_selected(self):
        selected = list(self.file_list.curselection())
        for index in reversed(selected):
            self.file_list.delete(index)
            del self.files[index]
        self.status.set(f"{len(self.files)} PowerPoint file(s) queued.")

    def clear_list(self):
        self.files.clear()
        self.file_list.delete(0, "end")
        self.status.set("File list cleared.")

    def choose_output(self):
        folder = filedialog.askdirectory(title="Choose output folder")
        if folder:
            self.output_dir.set(folder)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning(APP_TITLE, "Add at least one .pptx file first.")
            return
        out = Path(self.output_dir.get()).expanduser()
        try:
            out.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not create output folder:\n{exc}")
            return

        self.convert_button.configure(state="disabled")
        self.progress.configure(maximum=len(self.files), value=0)
        threading.Thread(target=self._convert_all, args=(out,), daemon=True).start()

    def _convert_all(self, output_dir):
        successes, failures = [], []
        soffice = find_libreoffice()

        try:
            import win32com.client  # noqa: F401
            has_powerpoint = True
        except Exception:
            has_powerpoint = False

        for index, raw in enumerate(self.files, start=1):
            source = Path(raw)
            destination = output_dir / (source.stem + ".pdf")
            try:
                self._set_status(f"Converting {index}/{len(self.files)}: {source.name}")
                if has_powerpoint:
                    convert_with_powerpoint(source, destination)
                    if not destination.exists():
                        raise RuntimeError("PowerPoint did not create the expected PDF.")
                elif soffice:
                    convert_with_libreoffice(source, output_dir, soffice)
                else:
                    raise RuntimeError(
                        "No supported converter found. Install Microsoft PowerPoint and run "
                        "'pip install pywin32', or install LibreOffice."
                    )
                successes.append(source.name)
            except Exception as exc:
                failures.append((source.name, str(exc)))

            self.root.after(0, self.progress.configure, {"value": index})

        def finish():
            self.convert_button.configure(state="normal")
            summary = f"Finished. Converted {len(successes)} of {len(self.files)} file(s)."
            if failures:
                summary += "\n\nFailed files:\n" + "\n".join(
                    f"• {name}: {error[:220]}" for name, error in failures
                )
            self.status.set(summary.replace("\n", " "))
            if failures:
                messagebox.showwarning(APP_TITLE, summary)
            else:
                messagebox.showinfo(APP_TITLE, summary + f"\n\nPDFs saved in:\n{output_dir}")
        self.root.after(0, finish)

    def _set_status(self, value):
        self.root.after(0, self.status.set, value)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
