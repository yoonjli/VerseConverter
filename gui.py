"""BilingualSlideMaker - Windows GUI for make_bilingual_pptx.py

Ways to give it a file:
  1. Drag the .txt onto the window (needs tkinterdnd2)
  2. Drag the .txt onto BilingualSlideMaker.exe in File Explorer
  3. Browse...
Output defaults to <same folder>/<same name>.pptx
"""
import os, sys, threading, queue
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import make_bilingual_pptx as core

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAVE_DND = True
except Exception:
    HAVE_DND = False


class TextRedirector:
    """Send print() output to the on-screen log (a windowed .exe has no console)."""
    def __init__(self, q): self.q = q
    def write(self, s):
        if s: self.q.put(s)
    def flush(self): pass


def default_output(in_path):
    return os.path.splitext(in_path)[0] + ".pptx"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Bilingual Slide Maker")
        root.geometry("760x560")
        self.q = queue.Queue()

        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()

        self.drop = tk.Label(
            root, relief="groove", bd=2, height=4, bg="#f4f4f4",
            text=("Drag & drop the combined .txt file here" if HAVE_DND
                  else "Click Browse... to choose the combined .txt file"))
        self.drop.pack(fill="x", padx=12, pady=(12, 6))

        row1 = tk.Frame(root); row1.pack(fill="x", padx=12, pady=3)
        tk.Label(row1, text="Input .txt:", width=11, anchor="w").pack(side="left")
        tk.Entry(row1, textvariable=self.in_var).pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="Browse...", command=self.browse_in).pack(side="left", padx=(6, 0))

        row2 = tk.Frame(root); row2.pack(fill="x", padx=12, pady=3)
        tk.Label(row2, text="Save as .pptx:", width=11, anchor="w").pack(side="left")
        tk.Entry(row2, textvariable=self.out_var).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="Save As...", command=self.browse_out).pack(side="left", padx=(6, 0))

        self.btn = tk.Button(root, text="Convert", command=self.convert, height=2)
        self.btn.pack(fill="x", padx=12, pady=8)

        self.log = scrolledtext.ScrolledText(root, height=16, state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        if HAVE_DND:
            for w in (root, self.drop):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop)

        # File dropped onto the .exe icon arrives as argv[1]
        if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
            self.set_input(sys.argv[1])

        self.poll()

    # ---- input / output handling
    def set_input(self, path):
        self.in_var.set(path)
        self.out_var.set(default_output(path))
        self.drop.config(text=f"Loaded: {os.path.basename(path)}")

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)   # handles {paths with spaces}
        if files:
            self.set_input(files[0])

    def browse_in(self):
        p = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if p: self.set_input(p)

    def browse_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".pptx",
                                         filetypes=[("PowerPoint", "*.pptx")],
                                         initialfile=os.path.basename(self.out_var.get() or "slides.pptx"))
        if p: self.out_var.set(p)

    # ---- log
    def poll(self):
        try:
            while True:
                s = self.q.get_nowait()
                self.log.config(state="normal"); self.log.insert("end", s)
                self.log.see("end"); self.log.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    # ---- conversion
    def convert(self):
        in_path = self.in_var.get().strip().strip('"')
        out_path = self.out_var.get().strip().strip('"')
        if not os.path.isfile(in_path):
            messagebox.showerror("Input file", "Please choose or drop the input .txt file."); return
        if not out_path:
            out_path = default_output(in_path)
        if not out_path.lower().endswith(".pptx"):
            out_path += ".pptx"
        self.out_var.set(out_path)
        self.btn.config(state="disabled")
        threading.Thread(target=self.work, args=(in_path, out_path), daemon=True).start()

    def work(self, in_path, out_path):
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = TextRedirector(self.q)
        try:
            from pptx import Presentation
            print(f"Parsing {in_path} ...")
            ko_lines, en_lines = core.split_bilingual_file(in_path)
            ko = core.parse_verses_from_lines(ko_lines)
            print(f"  -> {len(ko)} Korean verses")
            en = core.parse_verses_from_lines(en_lines)
            print(f"  -> {len(en)} English verses")
            pairs = core.match_verses(ko, en)
            print(f"  -> {len(pairs)} matched pairs")

            prs = Presentation()
            prs.slide_width, prs.slide_height = core.SLIDE_W, core.SLIDE_H
            adjustments = []
            for ref_ko, t_ko, ref_en, t_en in pairs:
                adjustments.extend(core.make_slide(prs, ref_ko, t_ko, ref_en, t_en) or [])
            prs.save(out_path)
            print(f"Saved {len(pairs)} slides -> {out_path}")
            core.print_review_summary(adjustments)
        except Exception as e:
            print(f"\nERROR: {e}")
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.root.after(0, lambda: self.btn.config(state="normal"))


if __name__ == "__main__":
    root = TkinterDnD.Tk() if HAVE_DND else tk.Tk()
    App(root)
    root.mainloop()
