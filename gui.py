"""
Simple point-and-click GUI for make_bilingual_pptx.py.

This file does NOT change the conversion logic at all — it imports
make_bilingual_pptx.py as-is and just calls its existing functions,
the same way the original command-line script does in main().

Packaged into a Windows .exe via PyInstaller (see the GitHub Actions
workflow in .github/workflows/build-windows-exe.yml).
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from pptx import Presentation

import make_bilingual_pptx as core


class TextRedirector:
    """
    Sends print() output into the on-screen log box instead of a real
    console. A windowed .exe (no console window) has no stdout, so any
    bare print() call — including the "Korean-only / English-only"
    warnings inside match_verses() — would otherwise crash the app.
    """

    def __init__(self, widget):
        self.widget = widget

    def write(self, msg):
        if not msg:
            return
        self.widget.configure(state="normal")
        self.widget.insert("end", msg)
        self.widget.see("end")
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()

    def flush(self):
        pass


class App(tk.Tk):
    ACCENT = "#D4AF37"
    BG = "#1a1a1a"
    FG = "white"

    def __init__(self):
        super().__init__()
        self.title("Bilingual Scripture Slide Maker")
        self.geometry("580x480")
        self.configure(bg=self.BG)
        self.resizable(False, False)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        tk.Label(
            self, text="Bilingual Scripture Slide Maker",
            font=("Segoe UI", 15, "bold"), fg=self.ACCENT, bg=self.BG,
        ).pack(pady=(18, 6))

        tk.Label(
            self, text="1. Choose your combined Korean + English text file",
            fg=self.FG, bg=self.BG,
        ).pack(anchor="w", padx=18, pady=(14, 2))
        row1 = tk.Frame(self, bg=self.BG)
        row1.pack(fill="x", padx=18)
        tk.Entry(row1, textvariable=self.input_path, state="readonly").pack(
            side="left", fill="x", expand=True
        )
        tk.Button(row1, text="Browse...", command=self.pick_input).pack(
            side="left", padx=(8, 0)
        )

        tk.Label(
            self, text="2. Choose where to save the PowerPoint",
            fg=self.FG, bg=self.BG,
        ).pack(anchor="w", padx=18, pady=(14, 2))
        row2 = tk.Frame(self, bg=self.BG)
        row2.pack(fill="x", padx=18)
        tk.Entry(row2, textvariable=self.output_path, state="readonly").pack(
            side="left", fill="x", expand=True
        )
        tk.Button(row2, text="Save As...", command=self.pick_output).pack(
            side="left", padx=(8, 0)
        )

        self.convert_btn = tk.Button(
            self, text="Convert", command=self.run_conversion,
            bg=self.ACCENT, fg="black", font=("Segoe UI", 11, "bold"),
            padx=20, pady=6,
        )
        self.convert_btn.pack(pady=18)

        tk.Label(self, text="Status", fg=self.FG, bg=self.BG).pack(
            anchor="w", padx=18
        )
        self.log = scrolledtext.ScrolledText(
            self, height=12, state="disabled", bg="#0d0d0d", fg="white",
            insertbackground="white",
        )
        self.log.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        # Route print() (including warnings from inside make_bilingual_pptx.py)
        # into the log box instead of a nonexistent console.
        sys.stdout = TextRedirector(self.log)
        sys.stderr = TextRedirector(self.log)

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Choose combined Korean + English text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                base = os.path.splitext(os.path.basename(path))[0]
                self.output_path.set(
                    os.path.join(os.path.dirname(path), f"{base}.pptx")
                )

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save PowerPoint as",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")],
        )
        if path:
            self.output_path.set(path)

    def run_conversion(self):
        in_file = self.input_path.get()
        out_file = self.output_path.get()

        if not in_file:
            messagebox.showwarning(
                "Missing file", "Please choose an input text file first."
            )
            return
        if not out_file:
            messagebox.showwarning(
                "Missing output", "Please choose where to save the PowerPoint."
            )
            return

        self.convert_btn.config(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        try:
            print(f"Parsing {in_file} ...")
            ko_lines, en_lines = core.split_bilingual_file(in_file)

            ko_verses = core.parse_verses_from_lines(ko_lines)
            print(f"  -> {len(ko_verses)} Korean verses after merging")

            en_verses = core.parse_verses_from_lines(en_lines)
            print(f"  -> {len(en_verses)} English verses after merging")

            print("Matching verses by number...")
            pairs = core.match_verses(ko_verses, en_verses)
            print(f"  -> {len(pairs)} matched pairs")

            prs = Presentation()
            prs.slide_width = core.SLIDE_W
            prs.slide_height = core.SLIDE_H

            for ref_ko, text_ko, ref_en, text_en in pairs:
                core.make_slide(prs, ref_ko, text_ko, ref_en, text_en)

            prs.save(out_file)
            print(f"Done! Saved {len(pairs)} slides -> {out_file}")
            messagebox.showinfo(
                "Success", f"Saved {len(pairs)} slides to:\n{out_file}"
            )
        except Exception as e:
            print(f"ERROR: {e}")
            messagebox.showerror("Conversion failed", str(e))
        finally:
            self.convert_btn.config(state="normal")


if __name__ == "__main__":
    App().mainloop()
