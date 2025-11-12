
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from pathlib import Path
import os

from conversion.engine import ConversionEngine
from utils.file_handler import FileHandler



class GUIManager:
    """Manager untuk GUI application"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Document Converter & Compressor - DOC/DOCX ↔ PDF")
        self.root.geometry("820x640")
        self.root.minsize(700, 520)
        self.root.resizable(True, True)

        # Styling
        self.style = ttk.Style(self.root)
        try:
            # prefer a modern theme if available
            if "clam" in self.style.theme_names():
                self.style.theme_use("clam")
            else:
                self.style.theme_use(self.style.theme_names()[0])
        except Exception:
            pass

        # Configure some custom styles
        self.style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        self.style.configure("Info.TLabel", font=("Segoe UI", 10))
        self.style.configure("TFrame", background="#f6f8fb")
        self.style.configure("Banner.TFrame", background="#2b6cb0")
        self.style.configure("Banner.TLabel", font=("Segoe UI", 14, "bold"), foreground="white", background="#2b6cb0")
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        self.style.map("Accent.TButton",
                       foreground=[("active", "white"), ("!disabled", "white")],
                       background=[("active", "#1e90ff"), ("!disabled", "#2b6cb0")])

        # Initialize variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.conversion_type = tk.StringVar(value="doc_to_pdf")
        self.conversion_method = tk.StringVar(value="auto")
        self.status_text = tk.StringVar(value="Ready")

        # Initialize conversion engine
        self.conversion_engine = ConversionEngine()
        self.has_ms_word = self.conversion_engine.check_ms_word_installation()
        self.conversion_engine.has_ms_word = self.has_ms_word

        # Setup UI
        self.setup_ui()

    def setup_ui(self):
        """Setup user interface"""
        # Menu
        self._setup_menu()

         # Banner
        banner = ttk.Frame(self.root, style="Banner.TFrame", padding=(12, 12))
        banner.grid(row=0, column=0, sticky=(tk.W, tk.E))
        # buat 3 kolom: icon kiri, judul di tengah (dengan weight), spacer kanan
        banner.columnconfigure(0, weight=0)
        banner.columnconfigure(1, weight=1)
        banner.columnconfigure(2, weight=0)

        banner_icon = ttk.Label(banner, text="📝", style="Banner.TLabel")
        banner_icon.grid(row=0, column=0, sticky=tk.W, padx=(0,8))

        title_label = ttk.Label(
            banner,
            text="Document Converter & Compressor",
            style="Banner.TLabel",
            anchor="center",
            justify="center"
        )
        # tidak pakai sticky agar widget berada di tengah sel
        title_label.grid(row=0, column=1)

        # Main frame
        main_frame = ttk.Frame(self.root, padding="16")
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)

        # Info sistem
        system_info = self._get_system_info()
        info_label = ttk.Label(main_frame, text=system_info, style="Info.TLabel")
        info_label.grid(row=0, column=0, sticky=tk.W, pady=(6, 12))

        # Pilihan metode untuk PDF ke DOCX
        self._setup_conversion_method_ui(main_frame)

        # Pilihan tipe konversi
        self._setup_conversion_type_ui(main_frame)

        # Input dan output file
        self._setup_file_selection_ui(main_frame)

        # Progress dan status
        self._setup_progress_ui(main_frame)

        # Log / preview area
        self._setup_log_ui(main_frame)

        # Tombol aksi
        self._setup_action_buttons(main_frame)

        # Konfigurasi grid
        self._configure_grid_weights(main_frame)

    def _setup_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Clear", command=self.clear_fields, accelerator="Ctrl+K")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)
        # keyboard shortcuts
        self.root.bind_all("<Control-q>", lambda e: self.root.quit())
        self.root.bind_all("<Control-k>", lambda e: self.clear_fields())

    def _show_about(self):
        messagebox.showinfo("About", "Document Converter\n\nDOC/DOCX ↔ PDF\nUI diperbarui untuk pengalaman yang lebih baik.")

    def _get_system_info(self) -> str:
        """Dapatkan info sistem dan dukungan konversi"""
        if self.has_ms_word:
            return "✅ DOC → PDF   ✅ DOCX → PDF   ✅ PDF → DOCX   ✅ PDF → DOC"
        else:
            return "❌ DOC → PDF   ✅ DOCX → PDF   ✅ PDF → DOCX   ❌ PDF → DOC"

    def _setup_conversion_method_ui(self, parent):
        """Setup UI untuk pemilihan metode konversi"""
        method_frame = ttk.LabelFrame(parent, text="Metode PDF ke DOCX", padding="10")
        method_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=8)
        method_frame.columnconfigure(0, weight=1)
        method_frame.columnconfigure(1, weight=1)

        ttk.Radiobutton(method_frame, text="Auto (Rekomendasi)", 
                       variable=self.conversion_method, value="auto").grid(row=0, column=0, sticky=tk.W, padx=6, pady=2)

        ttk.Radiobutton(method_frame, text="pdf2docx (Gambar + Formatting)", 
                       variable=self.conversion_method, value="pdf2docx").grid(row=0, column=1, sticky=tk.W, padx=6, pady=2)

        ttk.Radiobutton(method_frame, text="PyMuPDF (Text + Gambar)", 
                       variable=self.conversion_method, value="pymupdf").grid(row=1, column=0, sticky=tk.W, padx=6, pady=2)

        ttk.Radiobutton(method_frame, text="Text Only (Cepat)", 
                       variable=self.conversion_method, value="text_only").grid(row=1, column=1, sticky=tk.W, padx=6, pady=2)

    def _setup_conversion_type_ui(self, parent):
        """Setup UI untuk pemilihan tipe konversi"""
        type_frame = ttk.LabelFrame(parent, text="Tipe Konversi", padding="10")
        type_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=8)
        type_frame.columnconfigure(0, weight=1)

        ttk.Radiobutton(type_frame, text="DOC/DOCX ke PDF", 
                       variable=self.conversion_type, value="doc_to_pdf",
                       command=self._on_conversion_change).grid(row=0, column=0, sticky=tk.W, pady=2)

        ttk.Radiobutton(type_frame, text="PDF ke DOCX", 
                       variable=self.conversion_type, value="pdf_to_docx",
                       command=self._on_conversion_change).grid(row=1, column=0, sticky=tk.W, pady=2)

        ttk.Radiobutton(type_frame, text="PDF ke DOC", 
                       variable=self.conversion_type, value="pdf_to_doc",
                       command=self._on_conversion_change).grid(row=2, column=0, sticky=tk.W, pady=2)

    def _setup_file_selection_ui(self, parent):
        """Setup UI untuk seleksi file"""
        # Input file
        ttk.Label(parent, text="File Input:").grid(row=3, column=0, sticky=tk.W, pady=(12, 6))

        input_frame = ttk.Frame(parent)
        input_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=6)
        input_frame.columnconfigure(0, weight=1)

        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=70)
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))

        ttk.Button(input_frame, text="Browse", style="Accent.TButton", command=self.browse_input_file).grid(row=0, column=1, padx=(8, 0))

        # Output file
        ttk.Label(parent, text="File Output:").grid(row=5, column=0, sticky=tk.W, pady=(12, 6))

        output_frame = ttk.Frame(parent)
        output_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=6)
        output_frame.columnconfigure(0, weight=1)

        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=70)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))

        ttk.Button(output_frame, text="Browse", style="Accent.TButton", command=self.browse_output_file).grid(row=0, column=1, padx=(8, 0))

    def _setup_progress_ui(self, parent):
        """Setup UI untuk progress dan status"""
        self.progress = ttk.Progressbar(parent, mode='indeterminate')
        self.progress.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(14,6))

        status_frame = ttk.Frame(parent)
        status_frame.grid(row=8, column=0, sticky=(tk.W, tk.E), pady=(0,8))
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(status_frame, textvariable=self.status_text, style="Info.TLabel")
        self.status_label.grid(row=0, column=0, sticky=tk.W)

    def _setup_log_ui(self, parent):
        """Panel log / preview untuk memberikan feedback ke user"""
        log_frame = ttk.LabelFrame(parent, text="Log / Info", padding="8")
        log_frame.grid(row=9, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(6, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_widget = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD, state="disabled")
        self.log_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # write initial info
        self._log("Ready. Pilih file input atau drag & drop (tidak tersedia)")

    def _log(self, message: str):
        """Tambahkan pesan ke panel log"""
        self.log_widget.configure(state="normal")
        self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.log_widget.configure(state="disabled")

    def _setup_action_buttons(self, parent):
        """Setup tombol aksi"""
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=10, column=0, sticky=(tk.E), pady=6)
        ttk.Button(buttons_frame, text="Konversi", style="Accent.TButton", command=self.convert_document).grid(row=0, column=0, padx=(0,8))
        ttk.Button(buttons_frame, text="Bersihkan", command=self.clear_fields).grid(row=0, column=1)

    def _configure_grid_weights(self, parent):
        """Konfigurasi grid weights untuk responsive layout"""
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(9, weight=1)  # log area grows
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

    def _on_conversion_change(self):
        """Handler ketika tipe konversi berubah"""
        current_input = self.input_path.get()
        if current_input:
            self.auto_generate_output_path(current_input)

    def browse_input_file(self):
        """Browse file input"""
        file_types = self._get_input_file_types()

        filename = filedialog.askopenfilename(filetypes=file_types)
        if filename:
            if self.conversion_type.get() == "pdf_to_doc" and not self.has_ms_word:
                self._show_ms_word_required_error()
                return

            self.input_path.set(filename)
            self.auto_generate_output_path(filename)
            self._log(f"Selected input: {filename}")

    def _get_input_file_types(self):
        """Dapatkan tipe file untuk dialog input"""
        conversion_type = self.conversion_type.get()

        if conversion_type == "doc_to_pdf":
            if self.has_ms_word:
                return [
                    ("Word Documents", "*.docx"),
                    ("Word Documents", "*.doc"),
                    ("All supported", "*.docx;*.doc")
                ]
            else:
                return [("Word Documents (.docx only)", "*.docx")]
        else:
            return [
                ("PDF Files", "*.pdf"),
                ("All PDF Files", "*.pdf")
            ]

    def _show_ms_word_required_error(self):
        """Tampilkan error Microsoft Word required"""
        messagebox.showerror(
            "Microsoft Word Diperlukan",
            "Konversi PDF ke DOC membutuhkan Microsoft Word.\n\n"
            "Microsoft Word tidak terdeteksi di sistem Anda.\n"
            "Silakan:\n"
            "• Install Microsoft Word, atau\n"
            "• Gunakan opsi 'PDF ke DOCX' sebagai alternatif"
        )

    def browse_output_file(self):
        """Browse file output"""
        file_types, default_extension = self._get_output_file_types()

        filename = filedialog.asksaveasfilename(
            filetypes=file_types,
            defaultextension=default_extension
        )
        if filename:
            self.output_path.set(filename)
            self._log(f"Selected output: {filename}")

    def _get_output_file_types(self):
        """Dapatkan tipe file untuk dialog output"""
        conversion_type = self.conversion_type.get()

        if conversion_type == "doc_to_pdf":
            return [("PDF Files", "*.pdf")], ".pdf"
        elif conversion_type == "pdf_to_docx":
            return [("Word Documents", "*.docx")], ".docx"
        else:
            return [("Word Documents", "*.doc")], ".doc"

    def auto_generate_output_path(self, input_path):
        """Generate output path otomatis"""
        try:
            output_path = FileHandler.auto_generate_output_path(
                input_path,
                self.conversion_type.get()
            )
            self.output_path.set(output_path)
            self._log(f"Auto output: {output_path}")
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            self._log(f"Error generating output path: {e}")

    def convert_document(self):
        """Eksekusi konversi dokumen"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        # Validasi input
        if not self._validate_conversion_input(input_file, output_file):
            return

        try:
            self._start_conversion()

            # Eksekusi konversi
            success = self.conversion_engine.convert(
                conversion_type=self.conversion_type.get(),
                input_file=input_file,
                output_file=output_file,
                conversion_method=self.conversion_method.get()
            )

            if success:
                self._on_conversion_success(output_file)
            else:
                raise Exception("Konversi gagal tanpa error spesifik")

        except Exception as e:
            self._on_conversion_error(e)

    def _validate_conversion_input(self, input_file: str, output_file: str) -> bool:
        """Validasi input konversi"""
        if not input_file:
            messagebox.showerror("Error", "Pilih file input terlebih dahulu!")
            return False

        if not output_file:
            messagebox.showerror("Error", "Pilih lokasi output terlebih dahulu!")
            return False

        if not os.path.exists(input_file):
            messagebox.showerror("Error", "File input tidak ditemukan!")
            return False

        return True

    def _start_conversion(self):
        """Persiapan sebelum konversi"""
        self.progress.start(10)
        self.status_text.set("Sedang mengkonversi...")
        self._log("Konversi dimulai...")
        self.root.update()

    def _on_conversion_success(self, output_file: str):
        """Handler ketika konversi sukses"""
        self.progress.stop()
        self.status_text.set("Konversi berhasil!")
        self._log(f"Konversi selesai. Output: {output_file}")

        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            messagebox.showinfo("Sukses",
                f"File berhasil dikonversi!\n\n"
                f"Output: {output_file}\n"
                f"Size: {file_size} bytes")
        else:
            raise Exception("File output tidak berhasil dibuat")

    def _on_conversion_error(self, error: Exception):
        """Handler ketika konversi error"""
        self.progress.stop()
        self.status_text.set("Error!")
        error_msg = f"Terjadi kesalahan saat konversi:\n{str(error)}"
        print(f"ERROR DETAIL: {error_msg}")
        self._log(f"Error: {error_msg}")
        messagebox.showerror("Error", error_msg)

    def clear_fields(self):
        """Bersihkan semua field"""
        self.input_path.set("")
        self.output_path.set("")
        self.status_text.set("Ready")
        self._log("Fields cleared")

    def run(self):
        """Jalankan aplikasi"""
        self._check_dependencies()
        self.root.mainloop()

    def _check_dependencies(self):
        """Cek dependencies dan tampilkan warning jika perlu"""
        try:
            from conversion.strategies import LIBRARY_AVAILABLE
            if not LIBRARY_AVAILABLE:
                self._show_installation_instructions()
        except ImportError:
            self._show_installation_instructions()

    def _show_installation_instructions(self):
        """Tampilkan instruksi instalasi"""
        messagebox.showwarning(
            "Library Tidak Ditemukan",
            "Library diperlukan tidak terinstall.\n\n"
            "Untuk hasil terbaik (dengan gambar), install:\n"
            "pip install docx2pdf pdf2docx pymupdf python-docx comtypes\n\n"
            "Minimal installation:\n"
            "pip install docx2pdf pymupdf python-docx"
        )
