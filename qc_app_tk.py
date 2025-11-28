
# ============================
# Module 6: Tkinter Frontend
# ============================

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import sys

from qc_core import TXTParser, PDFParser, RBLoader, run_all_comparisons, organize_results, generate_qc_pdf_report

class QCApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gallo QC Automation")
        self.geometry("600x300")
        self.folder = tk.StringVar()
        tk.Label(self, text="Job Folder:").pack(pady=5)
        tk.Entry(self, textvariable=self.folder, width=50).pack()
        tk.Button(self, text="Browse", command=self.browse).pack(pady=5)
        tk.Button(self, text="Run QC", command=self.run_qc).pack(pady=10)
        self.out = tk.Label(self, text="")
        self.out.pack()

    def browse(self):
        p = filedialog.askdirectory()
        if p:
            self.folder.set(p)

    def run_qc(self):
        folder = self.folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error","Select valid folder.")
            return
        txt_path = self._find(folder, ".txt")
        pdf_path = self._find(folder, ".pdf")
        if not txt_path or not pdf_path:
            messagebox.showerror("Error","TXT or PDF missing.")
            return

        try:
            txt = TXTParser().load(txt_path)
            pdf = PDFParser().load(pdf_path)

            jobno = self._extract_job(folder)
            rb  = RBLoader().get_job_data(jobno)

            allres = run_all_comparisons(txt, pdf, rb)
            grouped = organize_results(allres)

            out_path = os.path.join(folder, f"{jobno}_QC_Report.pdf")
            generate_qc_pdf_report(out_path, grouped, allres, jobno, txt_data=txt, pdf_data=pdf, rb_data=rb)

            self.out.configure(text=f"QC Complete\nReport: {out_path}")
            if messagebox.askyesno("Open","Open folder?"):
                self._open(os.path.dirname(out_path))

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _find(self, folder, ext):
        for f in os.listdir(folder):
            if f.lower().endswith(ext):
                return os.path.join(folder,f)
        return None

    def _extract_job(self, folder):
        import re
        m = re.findall(r"\d{5,}", folder)
        return m[-1] if m else "UNKNOWN"

    def _open(self, path):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

if __name__ == "__main__":
    QCApp().mainloop()
