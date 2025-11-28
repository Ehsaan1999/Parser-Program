# qc_app_tk.py

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from qc_core import run_qc


class QCApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Gallo QC Automation")
        self.geometry("640x420")
        self.resizable(False, False)

        self.job_folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")
        self.result_summary_var = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # Job frame
        frm_job = ttk.LabelFrame(self, text="Job")
        frm_job.place(x=10, y=10, width=620, height=80)

        ttk.Label(frm_job, text="Job folder:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm_job, textvariable=self.job_folder_var, width=55).grid(row=0, column=1, **pad)
        ttk.Button(frm_job, text="Browse...", command=self._choose_job_folder).grid(row=0, column=2, **pad)

        # Run frame
        frm_run = ttk.LabelFrame(self, text="Run QC")
        frm_run.place(x=10, y=100, width=620, height=80)

        ttk.Button(frm_run, text="Run QC", command=self._run_qc_clicked).grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        ttk.Label(frm_run, textvariable=self.status_var).grid(
            row=0, column=1, padx=10, pady=10, sticky="w"
        )

        # Summary frame
        frm_res = ttk.LabelFrame(self, text="Summary")
        frm_res.place(x=10, y=190, width=620, height=210)

        ttk.Label(
            frm_res,
            textvariable=self.result_summary_var,
            justify="left",
            anchor="nw",
        ).pack(fill="both", expand=True, padx=10, pady=10)

    # ---------- UI helpers ----------

    def _choose_job_folder(self):
        folder = filedialog.askdirectory(title="Select Job Folder")
        if folder:
            self.job_folder_var.set(folder)

    # ---------- Run QC ----------

    def _run_qc_clicked(self):
        job_folder = self.job_folder_var.get().strip()
        if not job_folder or not os.path.isdir(job_folder):
            messagebox.showerror("Error", "Please select a valid job folder.")
            return

        self.status_var.set("Running QC...")
        self.result_summary_var.set("")
        self.update_idletasks()

        thread = threading.Thread(target=self._run_qc_worker, daemon=True)
        thread.start()

    def _run_qc_worker(self):
        try:
            summary, report_path = run_qc(self.job_folder_var.get().strip())
            counts = summary.counts

            msg = (
                f"Job {summary.job_number}\n"
                f"Witness: {summary.witness_name or '(unknown)'}\n\n"
                f"Exact matches: {counts['EXACT_MATCH']}\n"
                f"Partial matches: {counts['PARTIAL_MATCH']}\n"
                f"No matches: {counts['NO_MATCH']}\n"
                f"Missing: {counts['MISSING']}\n\n"
                f"Report: {report_path}"
            )

            def update_success():
                self.status_var.set("QC complete.")
                self.result_summary_var.set(msg)
                if messagebox.askyesno("QC complete", "Open report folder?"):
                    self._open_folder(os.path.dirname(report_path))

            self.after(0, update_success)

        except Exception as e:
            def update_error():
                self.status_var.set("Error.")
                messagebox.showerror("Error", f"An error occurred:\n{e}")

            self.after(0, update_error)

    def _open_folder(self, path: str):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


if __name__ == "__main__":
    app = QCApp()
    app.mainloop()