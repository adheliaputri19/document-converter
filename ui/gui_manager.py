# ui/gui_manager.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import threading
import os
import sys

# Fix path
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(ROOT_DIR)

from conversion.engine import ConversionEngine
from conversion.compressor import DocumentCompressor
from utils.file_handler import FileHandler


class GUIManager:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("📄 Document Converter & Compressor Pro")
        self.root.geometry("900x720")
        self.root.minsize(800, 600)
        
        # Set window icon if available
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        self._setup_styles()
        self._setup_variables()
        self._setup_engine()
        self._setup_ui()

    def _setup_styles(self):
        """Setup modern UI styles and themes"""
        style = ttk.Style(self.root)
        try:
            style.theme_use('clam')
        except Exception:
            pass
        
        # Configure modern color scheme
        self.colors = {
            'primary': '#2C3E50',
            'secondary': '#3498DB',
            'success': '#27AE60',
            'warning': '#F39C12',
            'danger': '#E74C3C',
            'light': '#ECF0F1',
            'dark': '#2C3E50',
            'background': '#F8F9FA'
        }
        
        # Configure styles
        style.configure('Title.TLabel', 
                       font=('Arial', 18, 'bold'),
                       foreground=self.colors['primary'],
                       background=self.colors['background'])
        
        style.configure('Subtitle.TLabel',
                       font=('Arial', 11),
                       foreground=self.colors['dark'],
                       background=self.colors['background'])
        
        style.configure('Card.TFrame',
                       background='white',
                       relief='raised',
                       borderwidth=1)
        
        style.configure('Accent.TButton',
                       font=('Arial', 10, 'bold'),
                       foreground='white',
                       background=self.colors['secondary'],
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Accent.TButton',
                 background=[('active', self.colors['primary']),
                           ('pressed', self.colors['dark'])])
        
        style.configure('Success.TButton',
                       font=('Arial', 10, 'bold'),
                       foreground='white',
                       background=self.colors['success'])
        
        style.map('Success.TButton',
                 background=[('active', '#219955'),
                           ('pressed', '#1E8449')])
        
        style.configure('Modern.TLabelframe',
                       background=self.colors['background'],
                       borderwidth=2,
                       relief='raised')
        
        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['background'],
                       foreground=self.colors['primary'],
                       font=('Arial', 10, 'bold'))

    def _setup_variables(self):
        """Setup tkinter variables"""
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.conversion_type = tk.StringVar(value="doc_to_pdf")
        self.method = tk.StringVar(value="auto")
        self.compress_files = []
        self.compress_output = tk.StringVar(value=os.path.expanduser("~/Documents"))
        self.compress_level = tk.StringVar(value="medium")
        self._status_message = tk.StringVar(value="🎯 Ready - Pilih file untuk memulai")

    def _setup_engine(self):
        """Setup conversion engine and compressor"""
        self.engine = ConversionEngine()
        self.has_ms_word = self.engine.check_ms_word_installation()
        self.compressor = DocumentCompressor()

    def _setup_ui(self):
        """Setup modern UI components"""
        # Configure root background
        self.root.configure(bg=self.colors['background'])
        
        # Main frame
        main = ttk.Frame(self.root, padding=20, style='Card.TFrame')
        main.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid weights
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        # Application Title
        self._create_app_title(main)
        
        # Header
        self._create_header(main)

        # Notebook (Tabs)
        nb = ttk.Notebook(main)
        nb.grid(row=2, column=0, sticky="nsew", pady=15)

        # Create tabs
        self.tab_convert = ttk.Frame(nb, style='Card.TFrame')
        self.tab_compress = ttk.Frame(nb, style='Card.TFrame')
        nb.add(self.tab_convert, text="🔄 Konversi Dokumen")
        nb.add(self.tab_compress, text="📦 Kompres Ukuran")

        self._setup_convert_tab()
        self._setup_compress_tab()

        # Status bar
        self._create_status_bar(main)

    def _create_app_title(self, parent):
        """Create application title above the menu"""
        title_frame = ttk.Frame(parent, style='Card.TFrame')
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Main application title
        app_title = ttk.Label(title_frame, 
                             text="Document Converter & Compressor Pro", 
                             font=('Arial', 20, 'bold'),
                             foreground=self.colors['primary'],
                             background=self.colors['background'])
        app_title.pack(pady=15)
        
        # Separator line
        separator = ttk.Separator(title_frame, orient='horizontal')
        separator.pack(fill='x', padx=20, pady=5)

    def _create_header(self, parent):
        """Create modern header"""
        header_frame = ttk.Frame(parent, style='Card.TFrame')
        header_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        # Title
        title_label = ttk.Label(header_frame, 
                               text="📄 Document Converter & Compressor Pro", 
                               style='Title.TLabel')
        title_label.pack(pady=10)
        
        # Subtitle with MS Word status
        status_icon = "✅" if self.has_ms_word else "⚠️"
        status_text = "MS Word: Terdeteksi" if self.has_ms_word else "MS Word: Tidak Terdeteksi"
        status_color = self.colors['success'] if self.has_ms_word else self.colors['warning']
        
        subtitle_frame = ttk.Frame(header_frame, style='Card.TFrame')
        subtitle_frame.pack(pady=5)
        
        ttk.Label(subtitle_frame, text=f"{status_icon} {status_text}", 
                 foreground=status_color, font=('Arial', 9, 'bold')).pack()

    def _create_status_bar(self, parent):
        """Create modern status bar"""
        status_frame = ttk.Frame(parent, style='Card.TFrame')
        status_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        
        status_label = ttk.Label(status_frame, textvariable=self._status_message,
                               font=('Arial', 9), foreground=self.colors['dark'])
        status_label.pack(side='left', padx=10, pady=5)

    def _setup_convert_tab(self):
        """Setup modern conversion tab"""
        f = self.tab_convert
        f.configure(padding=20)
        
        # Configure grid
        f.columnconfigure(1, weight=1)

        # File selection section
        file_frame = ttk.LabelFrame(f, text="📁 Seleksi File", style='Modern.TLabelframe')
        file_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)

        # Input file
        ttk.Label(file_frame, text="File Input:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky="w", pady=8, padx=10)
        
        input_frame = ttk.Frame(file_frame)
        input_frame.grid(row=0, column=1, columnspan=2, sticky="ew", pady=8, padx=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(input_frame, textvariable=self.input_path, font=('Arial', 10)).grid(
            row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(input_frame, text="📂 Browse", 
                  command=self._browse_input, style='Accent.TButton').grid(row=0, column=1)

        # Output file
        ttk.Label(file_frame, text="File Output:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky="w", pady=8, padx=10)
        
        output_frame = ttk.Frame(file_frame)
        output_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=8, padx=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_frame, textvariable=self.output_path, font=('Arial', 10)).grid(
            row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(output_frame, text="📂 Browse", 
                  command=self._browse_output, style='Accent.TButton').grid(row=0, column=1)

        # Conversion settings section
        settings_frame = ttk.LabelFrame(f, text="⚙️ Pengaturan Konversi", style='Modern.TLabelframe')
        settings_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        settings_frame.columnconfigure(1, weight=1)

        # Conversion type
        ttk.Label(settings_frame, text="Tipe Konversi:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky="w", pady=12, padx=10)
        
        type_combo = ttk.Combobox(settings_frame, textvariable=self.conversion_type, 
                                values=["doc_to_pdf", "pdf_to_docx", "pdf_to_doc"], 
                                state="readonly", font=('Arial', 10))
        type_combo.grid(row=0, column=1, sticky="ew", pady=12, padx=(0, 10))
        type_combo.bind("<<ComboboxSelected>>", self._update_output_suggestion)

        # Method selection
        ttk.Label(settings_frame, text="Metode Konversi:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky="w", pady=12, padx=10)
        
        ttk.Combobox(settings_frame, textvariable=self.method, 
                    values=["auto", "pdf2docx", "pymupdf", "text_only"], 
                    state="readonly", font=('Arial', 10)).grid(
                    row=1, column=1, sticky="ew", pady=12, padx=(0, 10))

        # Convert button
        button_frame = ttk.Frame(f)
        button_frame.grid(row=2, column=0, columnspan=3, pady=20)
        
        convert_btn = ttk.Button(button_frame, text="🚀 MULAI KONVERSI", 
                               command=self._start_convert, style='Success.TButton',
                               padding=(30, 12))
        convert_btn.pack()

        # Progress bar
        self.progress = ttk.Progressbar(f, mode="indeterminate", style='Accent.Horizontal.TProgressbar')
        self.progress.grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)

    def _setup_compress_tab(self):
        """Setup modern compression tab"""
        f = self.tab_compress
        f.configure(padding=20)
        f.columnconfigure(0, weight=1)

        # Compression settings
        settings_frame = ttk.LabelFrame(f, text="⚙️ Pengaturan Kompresi", style='Modern.TLabelframe')
        settings_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        settings_frame.columnconfigure(1, weight=1)

        # Compression level
        ttk.Label(settings_frame, text="Level Kompresi:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky="w", pady=12, padx=10)
        
        ttk.Combobox(settings_frame, textvariable=self.compress_level, 
                    values=["low", "medium", "high"], 
                    state="readonly", font=('Arial', 10)).grid(
                    row=0, column=1, sticky="ew", pady=12, padx=(0, 10))

        # Output folder
        ttk.Label(settings_frame, text="Folder Output:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky="w", pady=12, padx=10)
        
        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=1, column=1, sticky="ew", pady=12, padx=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_frame, textvariable=self.compress_output, font=('Arial', 10)).grid(
            row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(output_frame, text="📁 Pilih", 
                  command=self._browse_output_folder, style='Accent.TButton').grid(row=0, column=1)

        # File list section
        list_frame = ttk.LabelFrame(f, text="📋 Daftar File untuk Kompresi", style='Modern.TLabelframe')
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)

        # Listbox with scrollbar in a frame
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        list_container.columnconfigure(0, weight=1)
        list_container.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_container, height=12, font=('Arial', 10),
                                 bg='white', relief='solid', borderwidth=1)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Drag & drop functionality
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self._drop_files)

        # Info text for drag & drop
        info_label = ttk.Label(list_frame, 
                              text="💡 Drag & drop file PDF/DOCX ke area di atas",
                              font=('Arial', 9), foreground=self.colors['dark'])
        info_label.grid(row=1, column=0, pady=5)

        # File management buttons
        btn_frame = ttk.Frame(f)
        btn_frame.grid(row=2, column=0, pady=10)
        
        ttk.Button(btn_frame, text="➕ Tambah File", 
                  command=self._add_files, style='Accent.TButton').grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="📁 Tambah Folder", 
                  command=self._add_folder, style='Accent.TButton').grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="🗑️ Hapus Semua", 
                  command=self._clear_files, style='Accent.TButton').grid(row=0, column=2, padx=5)

        # Compress button
        compress_btn = ttk.Button(f, text="📦 MULAI KOMPRES", 
                                 command=self._start_compress, style='Success.TButton',
                                 padding=(30, 12))
        compress_btn.grid(row=3, column=0, pady=10)

        # Progress bar
        self.c_progress = ttk.Progressbar(f, mode="determinate", style='Accent.Horizontal.TProgressbar')
        self.c_progress.grid(row=4, column=0, sticky="ew", pady=10)

    # File browsing methods (unchanged functionality)
    def _browse_input(self):
        self._status_message.set("📂 Memilih file input...")
        file = filedialog.askopenfilename(
            title="Pilih File Input",
            filetypes=[("Dokumen", "*.doc *.docx *.pdf"), ("Semua File", "*.*")])
        if file:
            self.input_path.set(file)
            self._update_output_suggestion()
            self._status_message.set(f"✅ File input dipilih: {Path(file).name}")

    def _browse_output(self):
        self._status_message.set("📂 Memilih lokasi output...")
        ext_map = {".pdf": "*.pdf", ".docx": "*.docx", ".doc": "*.doc"}
        typ = self.conversion_type.get()
        default_ext = {
            "doc_to_pdf": ".pdf", 
            "pdf_to_docx": ".docx", 
            "pdf_to_doc": ".doc"
        }.get(typ, ".pdf")
        
        file = filedialog.asksaveasfilename(
            title="Simpan File Output",
            defaultextension=default_ext,
            filetypes=[("File", ext_map[default_ext])])
        
        if file:
            self.output_path.set(file)
            self._status_message.set(f"✅ Lokasi output dipilih: {Path(file).name}")

    def _browse_output_folder(self):
        self._status_message.set("📂 Memilih folder output...")
        folder = filedialog.askdirectory(title="Pilih Folder Output")
        if folder:
            self.compress_output.set(folder)
            self._status_message.set(f"✅ Folder output dipilih: {Path(folder).name}")

    def _update_output_suggestion(self, event=None):
        inp = self.input_path.get()
        if inp:
            self.output_path.set(
                FileHandler.auto_generate_output_path(inp, self.conversion_type.get()))
            self._status_message.set("✅ Path output otomatis dihasilkan")

    # File management methods (unchanged functionality)
    def _add_files(self):
        self._status_message.set("📂 Menambahkan file...")
        files = filedialog.askopenfilenames(
            title="Pilih File untuk Kompresi",
            filetypes=[("Dokumen", "*.pdf *.docx"), ("Semua File", "*.*")])
        if files:
            self.compress_files.extend(files)
            self._update_list()
            self._status_message.set(f"✅ {len(files)} file ditambahkan")

    def _add_folder(self):
        self._status_message.set("📂 Menambahkan folder...")
        folder = filedialog.askdirectory(title="Pilih Folder")
        if folder:
            new_files = [
                str(p) for p in Path(folder).rglob('*') 
                if p.suffix.lower() in {'.pdf', '.docx'}
            ]
            self.compress_files.extend(new_files)
            self._update_list()
            self._status_message.set(f"✅ {len(new_files)} file dari folder ditambahkan")

    def _drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        new_files = [
            f for f in files 
            if Path(f).suffix.lower() in {'.pdf', '.docx'}
        ]
        self.compress_files.extend(new_files)
        self._update_list()
        self._status_message.set(f"✅ {len(new_files)} file di-drop")

    def _update_list(self):
        self.listbox.delete(0, tk.END)
        for f in self.compress_files:
            self.listbox.insert(tk.END, f"📄 {Path(f).name}")
        self._status_message.set(f"📋 {len(self.compress_files)} file dalam daftar")

    def _clear_files(self):
        self.compress_files.clear()
        self._update_list()
        self._status_message.set("🗑️ Semua file dihapus dari daftar")

    # Conversion methods (unchanged functionality)
    def _start_convert(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showwarning("Peringatan", "❌ Pilih file input dan output terlebih dahulu!")
            return
        
        self._status_message.set("🔄 Memulai proses konversi...")
        self.progress.start()
        threading.Thread(target=self._run_convert, daemon=True).start()

    def _run_convert(self):
        try:
            kwargs = {'method': self.method.get()} if self.conversion_type.get() == 'pdf_to_docx' else {}
            self.engine.convert(
                self.conversion_type.get(), 
                self.input_path.get(), 
                self.output_path.get(), 
                **kwargs
            )
            self.root.after(0, lambda: messagebox.showinfo("Sukses", "✅ Konversi berhasil diselesaikan!"))
            self.root.after(0, lambda: self._status_message.set("✅ Konversi selesai!"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"❌ Gagal: {str(e)}"))
            self.root.after(0, lambda: self._status_message.set(f"❌ Error: {str(e)}"))
        finally:
            self.root.after(0, self.progress.stop)

    # Compression methods (unchanged functionality)
    def _start_compress(self):
        if not self.compress_files:
            messagebox.showwarning("Peringatan", "❌ Pilih file terlebih dahulu!")
            return
        
        if not os.path.isdir(self.compress_output.get()):
            messagebox.showerror("Error", "❌ Pilih folder output yang valid!")
            return
        
        self._status_message.set("📦 Memulai proses kompresi...")
        threading.Thread(target=self._run_compress, daemon=True).start()

    def _run_compress(self):
        total = len(self.compress_files)
        self.root.after(0, lambda: self.c_progress.config(maximum=total))
        self.root.after(0, lambda: self.c_progress.config(value=0))
        
        level = self.compress_level.get()
        out_dir = Path(self.compress_output.get())
        success_count = 0

        for i, f in enumerate(self.compress_files):
            self.root.after(0, lambda v=i+1: self.c_progress.config(value=v))
            self.root.after(0, lambda: self._status_message.set(
                f"📦 Memproses file {i+1}/{total}"))
            
            try:
                name = Path(f).name
                out_path = out_dir / f"compressed_{name}"
                self.compressor.compress(f, str(out_path), level)
                success_count += 1
            except Exception as e:
                print(f"[ERROR] {f}: {e}")

        self.root.after(0, lambda: messagebox.showinfo("Sukses", 
            f"✅ {success_count}/{total} file berhasil dikompres!\nLokasi: {out_dir}"))
        self.root.after(0, lambda: self._status_message.set(
            f"✅ Kompresi selesai: {success_count}/{total} file berhasil"))

    def run(self):
        """Start the GUI application"""
        self._status_message.set("🚀 Aplikasi siap digunakan!")
        self.root.mainloop()