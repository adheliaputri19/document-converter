
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import threading
import os


ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import sys
sys.path.append(ROOT_DIR)

from conversion.engine import ConversionEngine
from conversion.compressor import DocumentCompressor
from utils.file_handler import FileHandler



class GUIManager:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Document Converter & Compressor")
        self.root.geometry("850x680")
        self.root.minsize(750, 500)

        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.conversion_type = tk.StringVar(value="doc_to_pdf")
        self.method = tk.StringVar(value="auto")
        self.compress_files = []
        self.compress_output = tk.StringVar(value=os.getcwd())
        self.compress_level = tk.StringVar(value="medium")

        # Engine
        self.engine = ConversionEngine()
        self.has_ms_word = self.engine.check_ms_word_installation()
        self.compressor = DocumentCompressor()

        self._setup_ui()

    def _setup_ui(self):
        main = ttk.Frame(self.root, padding=15)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        nb = ttk.Notebook(main)
        nb.grid(row=0, column=0, sticky="nsew")
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        self.tab_convert = ttk.Frame(nb)
        self.tab_compress = ttk.Frame(nb)
        nb.add(self.tab_convert, text="Konversi")
        nb.add(self.tab_compress, text="Kompres Ukuran")

        self._setup_convert_tab()
        self._setup_compress_tab()

    def _setup_convert_tab(self):
        f = self.tab_convert
        ttk.Label(f, text="Konversi Dokumen", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=15)
        info = f"MS Word: {'Terdeteksi' if self.has_ms_word else 'Tidak'}"
        ttk.Label(f, text=info, foreground="green" if self.has_ms_word else "red").grid(row=1, column=0, columnspan=3)

        # Input
        ttk.Label(f, text="File Input:").grid(row=2, column=0, sticky="e", pady=5)
        ttk.Entry(f, textvariable=self.input_path, width=50).grid(row=2, column=1, pady=5)
        ttk.Button(f, text="Browse", command=self._browse_input).grid(row=2, column=2, pady=5)

        # Output
        ttk.Label(f, text="File Output:").grid(row=3, column=0, sticky="e", pady=5)
        ttk.Entry(f, textvariable=self.output_path, width=50).grid(row=3, column=1, pady=5)
        ttk.Button(f, text="Browse", command=self._browse_output).grid(row=3, column=2, pady=5)

        # Tipe
        ttk.Label(f, text="Tipe Konversi:").grid(row=4, column=0, sticky="e", pady=5)
        combo = ttk.Combobox(f, textvariable=self.conversion_type, values=["doc_to_pdf", "pdf_to_docx", "pdf_to_doc"], state="readonly")
        combo.grid(row=4, column=1, pady=5)
        combo.bind("<<ComboboxSelected>>", self._update_output_suggestion)

        # Metode
        ttk.Label(f, text="Metode (PDF→DOCX):").grid(row=5, column=0, sticky="e", pady=5)
        ttk.Combobox(f, textvariable=self.method, values=["auto", "pdf2docx", "pymupdf", "text_only"], state="readonly").grid(row=5, column=1, pady=5)

        # Tombol
        btn = ttk.Button(f, text="MULAI KONVERSI", command=self._start_convert)
        btn.grid(row=6, column=1, pady=25)

        # Progress
        self.progress = ttk.Progressbar(f, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

    def _browse_input(self):
        file = filedialog.askopenfilename(filetypes=[("Dokumen", "*.doc *.docx *.pdf")])
        if file:
            self.input_path.set(file)
            self._update_output_suggestion()

    def _browse_output(self):
        ext_map = {".pdf": "*.pdf", ".docx": "*.docx", ".doc": "*.doc"}
        typ = self.conversion_type.get()
        default_ext = { "doc_to_pdf": ".pdf", "pdf_to_docx": ".docx", "pdf_to_doc": ".doc" }.get(typ, ".pdf")
        file = filedialog.asksaveasfilename(defaultextension=default_ext, filetypes=[("File", ext_map[default_ext])])
        if file:
            self.output_path.set(file)

    def _update_output_suggestion(self, event=None):
        inp = self.input_path.get()
        if inp:
            self.output_path.set(FileHandler.auto_generate_output_path(inp, self.conversion_type.get()))

    def _start_convert(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showwarning("Error", "Pilih file input dan output dulu!")
            return
        self.progress.start()
        threading.Thread(target=self._run_convert, daemon=True).start()

    def _run_convert(self):
        try:
            kwargs = {'method': self.method.get()} if self.conversion_type.get() == 'pdf_to_docx' else {}
            self.engine.convert(self.conversion_type.get(), self.input_path.get(), self.output_path.get(), **kwargs)
            messagebox.showinfo("Sukses", "Konversi selesai!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal: {str(e)}")
        finally:
            self.progress.stop()

    def _setup_compress_tab(self):
        f = self.tab_compress
        ttk.Label(f, text="Kompres Ukuran File", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=15)

        # Level
        ttk.Label(f, text="Level Kompresi:").grid(row=1, column=0, sticky="e", pady=5)
        ttk.Combobox(f, textvariable=self.compress_level, values=["low", "medium", "high"], state="readonly").grid(row=1, column=1, pady=5)

        # Daftar file
        ttk.Label(f, text="File untuk dikompres:").grid(row=2, column=0, columnspan=3, pady=5)
        frame = ttk.Frame(f)
        frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=5)
        f.rowconfigure(3, weight=1)

        self.listbox = tk.Listbox(frame, height=12)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Drag & drop
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self._drop_files)

        # Tombol
        btn_frame = ttk.Frame(f)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=5)
        ttk.Button(btn_frame, text="Tambah File", command=self._add_files).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Tambah Folder", command=self._add_folder).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Hapus Semua", command=self._clear_files).grid(row=0, column=2, padx=5)

        # Output folder
        ttk.Label(f, text="Folder Output:").grid(row=5, column=0, sticky="e", pady=5)
        ttk.Entry(f, textvariable=self.compress_output, width=50).grid(row=5, column=1, pady=5)
        ttk.Button(f, text="Browse", command=self._browse_output_folder).grid(row=5, column=2, pady=5)

        # Tombol kompres
        ttk.Button(f, text="MULAI KOMPRES", command=self._start_compress).grid(row=6, column=1, pady=20)

        # Progress
        self.c_progress = ttk.Progressbar(f, mode="determinate")
        self.c_progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

    def _add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Dokumen", "*.pdf *.docx")])
        if files:
            self.compress_files.extend(files)
            self._update_list()

    def _add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.compress_files.extend([str(p) for p in Path(folder).rglob('*') if p.suffix.lower() in {'.pdf', '.docx'}])
            self._update_list()

    def _drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        self.compress_files.extend([f for f in files if Path(f).suffix.lower() in {'.pdf', '.docx'}])
        self._update_list()

    def _update_list(self):
        self.listbox.delete(0, tk.END)
        for f in self.compress_files:
            self.listbox.insert(tk.END, Path(f).name)

    def _clear_files(self):
        self.compress_files.clear()
        self._update_list()

    def _browse_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.compress_output.set(folder)

    def _start_compress(self):
        if not self.compress_files:
            messagebox.showwarning("Error", "Pilih file dulu!")
            return
        if not os.path.isdir(self.compress_output.get()):
            messagebox.showerror("Error", "Pilih folder output yang valid!")
            return
        threading.Thread(target=self._run_compress, daemon=True).start()

    def _run_compress(self):
        total = len(self.compress_files)
        self.c_progress["maximum"] = total
        self.c_progress["value"] = 0
        level = self.compress_level.get()
        out_dir = Path(self.compress_output.get())

        for i, f in enumerate(self.compress_files):
            self.c_progress["value"] = i + 1
            self.root.update_idletasks()
            try:
                name = Path(f).name
                out_path = out_dir / f"compressed_{name}"
                self.compressor.compress(f, str(out_path), level)
            except Exception as e:
                print(f"[ERROR] {f}: {e}")

        messagebox.showinfo("Sukses", f"{total} file berhasil dikompres di:\n{out_dir}")

    def run(self):
        self.root.mainloop()