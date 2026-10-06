```python
#!/usr/bin/env python3
"""
PDF Merger & Converter - Desktop App
Features:
- Add PDF, Word, and PowerPoint files
- Auto-converts Word and PPT to PDF using MS Office COM interface
- Reorder manually or auto-sort
- Merge and save as a new PDF
"""

import os
import re
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    raise SystemExit("Missing dependency, run: pip install pypdf")

try:
    import comtypes
    import comtypes.client  # Explicitly import client module first

    # Fix comtypes cache permission error after PyInstaller bundling
    if getattr(sys, 'frozen', False):
        comtypes_cache = os.path.join(tempfile.gettempdir(), 'comtypes_gen')
        if not os.path.exists(comtypes_cache):
            os.makedirs(comtypes_cache)
        comtypes.client.gen_dir = comtypes_cache
except ImportError:
    raise SystemExit("Missing dependency, run: pip install comtypes")


def natural_sort_key(path):
    name = os.path.basename(path).lower()
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


class PDFMergerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Merger & Converter (MS Office)")
        self.geometry("620x520")
        self.minsize(520, 440)
        
        self.file_paths = [] 
        self.temp_dir = tempfile.TemporaryDirectory()
        self.sort_field = tk.StringVar()
        self.sort_order = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        top_frame = tk.Frame(self, pady=8)
        top_frame.pack(fill="x", padx=10)

        self.btn_add = tk.Button(top_frame, text="Add Files", width=13, command=self.add_files)
        self.btn_add.pack(side="left", padx=4)
        self.btn_remove = tk.Button(top_frame, text="Remove Selected", width=15, command=self.remove_selected)
        self.btn_remove.pack(side="left", padx=4)
        self.btn_clear = tk.Button(top_frame, text="Clear All", width=13, command=self.clear_all)
        self.btn_clear.pack(side="left", padx=4)

        list_frame = tk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=10, pady=4)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame, selectmode="extended", yscrollcommand=scrollbar.set, activestyle="dotbox"
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        order_frame = tk.Frame(self, pady=6)
        order_frame.pack(fill="x", padx=10)

        self.btn_up = tk.Button(order_frame, text="↑ Move Up", width=12, command=self.move_up)
        self.btn_up.pack(side="left", padx=4)
        self.btn_down = tk.Button(order_frame, text="↓ Move Down", width=12, command=self.move_down)
        self.btn_down.pack(side="left", padx=4)

        self.count_label = tk.Label(order_frame, text="0 file(s)")
        self.count_label.pack(side="right", padx=4)

        sort_frame = tk.Frame(self, pady=4)
        sort_frame.pack(fill="x", padx=10)

        self.sort_label = tk.Label(sort_frame, text="Sort by:")
        self.sort_label.pack(side="left", padx=(0, 4))

        self.sort_field_menu = ttk.Combobox(
            sort_frame, textvariable=self.sort_field, state="readonly", width=12,
            values=["File Name", "Modified Date"]
        )
        self.sort_field_menu.current(0)
        self.sort_field_menu.pack(side="left", padx=4)

        self.sort_order_menu = ttk.Combobox(
            sort_frame, textvariable=self.sort_order, state="readonly", width=12,
            values=["Ascending", "Descending"]
        )
        self.sort_order_menu.current(0)
        self.sort_order_menu.pack(side="left", padx=4)

        self.btn_sort = tk.Button(sort_frame, text="Sort", width=10, command=self.apply_sort)
        self.btn_sort.pack(side="left", padx=4)

        bottom_frame = tk.Frame(self, pady=10)
        bottom_frame.pack(fill="x", padx=10)

        self.progress = ttk.Progressbar(bottom_frame, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.btn_merge = tk.Button(
            bottom_frame, text="Merge PDFs & Save", bg="#2f6fed", fg="white",
            font=("", 11, "bold"), height=2, command=self.merge_pdfs
        )
        self.btn_merge.pack(fill="x")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Add Files", 
            filetypes=[
                ("Supported Files", "*.pdf *.docx *.doc *.pptx *.ppt"),
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx *.doc"),
                ("PowerPoint files", "*.pptx *.ppt")
            ]
        )
        
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            target_path = p

            if ext in ['.doc', '.docx', '.ppt', '.pptx']:
                self.title(f"Converting {os.path.basename(p)}...")
                self.update()
                try:
                    # Add dynamic=True to avoid generating COM cache, solving permission issues after bundling
                    target_path = self.convert_to_pdf_ms(p)
                except Exception as e:
                    messagebox.showerror("Conversion Failed", f"Failed to convert {os.path.basename(p)}:\n{e}")
                    self.title("PDF Merger & Converter (MS Office)")
                    continue
                self.title("PDF Merger & Converter (MS Office)")

            if target_path not in self.file_paths:
                self.file_paths.append(target_path)
                self.listbox.insert("end", os.path.basename(target_path))
                
        self._update_count()

    def convert_to_pdf_ms(self, input_path):
        in_path = os.path.abspath(input_path)
        base_name = os.path.basename(input_path).rsplit('.', 1)[0]
        out_path = os.path.abspath(os.path.join(self.temp_dir.name, f"{base_name}.pdf"))
        
        ext = os.path.splitext(input_path)[1].lower()
        
        if ext in ['.doc', '.docx']:
            # Use dynamic=True for late binding
            word = comtypes.client.CreateObject('Word.Application', dynamic=True)
            word.Visible = False
            doc = word.Documents.Open(in_path)
            doc.SaveAs(out_path, 17) 
            doc.Close()
            word.Quit()
        elif ext in ['.ppt', '.pptx']:
            powerpoint = comtypes.client.CreateObject('Powerpoint.Application', dynamic=True)
            ppt = powerpoint.Presentations.Open(in_path, WithWindow=False)
            ppt.SaveAs(out_path, 32) 
            ppt.Close()
            powerpoint.Quit()
            
        return out_path

    def remove_selected(self):
        selected = list(self.listbox.curselection())
        for idx in reversed(selected):
            self.listbox.delete(idx)
            del self.file_paths[idx]
        self._update_count()

    def clear_all(self):
        self.listbox.delete(0, "end")
        self.file_paths.clear()
        self._update_count()

    def move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0: return
        for idx in sel: self._swap(idx, idx - 1)
        self.listbox.selection_clear(0, "end")
        for idx in sel: self.listbox.selection_set(idx - 1)

    def move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[-1] == len(self.file_paths) - 1: return
        for idx in reversed(sel): self._swap(idx, idx + 1)
        self.listbox.selection_clear(0, "end")
        for idx in sel: self.listbox.selection_set(idx + 1)

    def apply_sort(self):
        if not self.file_paths: return
        is_date = self.sort_field_menu.current() == 1
        is_desc = self.sort_order_menu.current() == 1
        key_func = (lambda p: os.path.getmtime(p)) if is_date else natural_sort_key
        self.file_paths.sort(key=key_func, reverse=is_desc)
        self.listbox.delete(0, "end")
        for p in self.file_paths:
            self.listbox.insert("end", os.path.basename(p))

    def _swap(self, i, j):
        self.file_paths[i], self.file_paths[j] = self.file_paths[j], self.file_paths[i]
        text_i = self.listbox.get(i)
        text_j = self.listbox.get(j)
        self.listbox.delete(min(i, j), max(i, j))
        first, second = (text_j, text_i) if i < j else (text_i, text_j)
        self.listbox.insert(min(i, j), first)
        self.listbox.insert(max(i, j), second)

    def _update_count(self):
        self.count_label.config(text=f"{len(self.file_paths)} file(s)")

    def merge_pdfs(self):
        if len(self.file_paths) < 2:
            messagebox.showwarning("Notice", "Please add at least 2 files before merging.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save merged PDF", defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")], initialfile="merged.pdf"
        )
        if not save_path: return

        writer = PdfWriter()
        self.progress["maximum"] = len(self.file_paths)
        self.progress["value"] = 0

        try:
            for i, path in enumerate(self.file_paths):
                reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)
                self.progress["value"] = i + 1
                self.update_idletasks()
            with open(save_path, "wb") as f:
                writer.write(f)
        except Exception as e:
            messagebox.showerror("Merge failed", f"An error occurred:\n{e}")
            return
        finally:
            self.progress["value"] = 0

        messagebox.showinfo("Done", f"Successfully merged {len(self.file_paths)} file(s)!\nSaved to:\n{save_path}")

if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()