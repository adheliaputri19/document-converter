import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import os

from conversion.engine import ConversionEngine
from utils.file_handler import FileHandler


class GUIManager:
    """Manager untuk GUI application"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Document Converter - DOC/DOCX ↔ PDF")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Initialize variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.conversion_type = tk.StringVar(value="doc_to_pdf")
        self.conversion_method = tk.StringVar(value="auto")
        self.status_text = tk.StringVar(value="")
        
        # Initialize conversion engine
        self.conversion_engine = ConversionEngine()
        self.has_ms_word = self.conversion_engine.check_ms_word_installation()
        self.conversion_engine.has_ms_word = self.has_ms_word
        
        # Setup UI
        self.setup_ui()
    
    def setup_ui(self):
        """Setup user interface"""
        # Frame utama
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Judul
        title_label = ttk.Label(main_frame, text="Document Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Info sistem
        system_info = self._get_system_info()
        info_label = ttk.Label(main_frame, text=system_info, font=("Arial", 9))
        info_label.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        
        # Pilihan metode untuk PDF ke DOCX
        self._setup_conversion_method_ui(main_frame)
        
        # Pilihan tipe konversi
        self._setup_conversion_type_ui(main_frame)
        
        # Input dan output file
        self._setup_file_selection_ui(main_frame)
        
        # Progress dan status
        self._setup_progress_ui(main_frame)
        
        # Tombol aksi
        self._setup_action_buttons(main_frame)
        
        # Konfigurasi grid
        self._configure_grid_weights(main_frame)
    
    def _get_system_info(self) -> str:
        """Dapatkan info sistem dan dukungan konversi"""
        if self.has_ms_word:
            return "✅ DOC → PDF   ✅ DOCX → PDF   ✅ PDF → DOCX   ✅ PDF → DOC"
        else:
            return "❌ DOC → PDF   ✅ DOCX → PDF   ✅ PDF → DOCX   ❌ PDF → DOC"
    
    def _setup_conversion_method_ui(self, parent):
        """Setup UI untuk pemilihan metode konversi"""
        method_frame = ttk.LabelFrame(parent, text="Metode PDF ke DOCX", padding="10")
        method_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Radiobutton(method_frame, text="Auto (Rekomendasi)", 
                       variable=self.conversion_method, value="auto").grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(method_frame, text="pdf2docx (Gambar + Formatting)", 
                       variable=self.conversion_method, value="pdf2docx").grid(row=0, column=1, sticky=tk.W)
        
        ttk.Radiobutton(method_frame, text="PyMuPDF (Text + Gambar)", 
                       variable=self.conversion_method, value="pymupdf").grid(row=1, column=0, sticky=tk.W)
        
        ttk.Radiobutton(method_frame, text="Text Only (Cepat)", 
                       variable=self.conversion_method, value="text_only").grid(row=1, column=1, sticky=tk.W)
    
    def _setup_conversion_type_ui(self, parent):
        """Setup UI untuk pemilihan tipe konversi"""
        ttk.Radiobutton(parent, text="DOC/DOCX ke PDF", 
                       variable=self.conversion_type, value="doc_to_pdf",
                       command=self._on_conversion_change).grid(row=3, column=0, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(parent, text="PDF ke DOCX", 
                       variable=self.conversion_type, value="pdf_to_docx",
                       command=self._on_conversion_change).grid(row=4, column=0, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(parent, text="PDF ke DOC", 
                       variable=self.conversion_type, value="pdf_to_doc",
                       command=self._on_conversion_change).grid(row=5, column=0, sticky=tk.W, pady=5)
    
    def _setup_file_selection_ui(self, parent):
        """Setup UI untuk seleksi file"""
        # Input file
        ttk.Label(parent, text="File Input:").grid(row=6, column=0, sticky=tk.W, pady=(20, 5))
        
        input_frame = ttk.Frame(parent)
        input_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=50)
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(input_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=1, padx=(5, 0))
        
        # Output file
        ttk.Label(parent, text="File Output:").grid(row=8, column=0, sticky=tk.W, pady=(20, 5))
        
        output_frame = ttk.Frame(parent)
        output_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(output_frame, text="Browse", command=self.browse_output_file).grid(row=0, column=1, padx=(5, 0))
    
    def _setup_progress_ui(self, parent):
        """Setup UI untuk progress dan status"""
        self.progress = ttk.Progressbar(parent, mode='indeterminate')
        self.progress.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        
        self.status_label = ttk.Label(parent, textvariable=self.status_text)
        self.status_label.grid(row=11, column=0, columnspan=2, pady=10)
    
    def _setup_action_buttons(self, parent):
        """Setup tombol aksi"""
        ttk.Button(parent, text="Konversi", command=self.convert_document).grid(row=12, column=0, pady=10)
        ttk.Button(parent, text="Bersihkan", command=self.clear_fields).grid(row=12, column=1, pady=10)
    
    def _configure_grid_weights(self, parent):
        """Konfigurasi grid weights untuk responsive layout"""
        parent.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
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
        except ValueError as e:
            messagebox.showerror("Error", str(e))
    
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
        self.progress.start()
        self.status_text.set("Sedang mengkonversi...")
        self.root.update()
    
    def _on_conversion_success(self, output_file: str):
        """Handler ketika konversi sukses"""
        self.progress.stop()
        self.status_text.set("Konversi berhasil!")
        
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
        messagebox.showerror("Error", error_msg)
    
    def clear_fields(self):
        """Bersihkan semua field"""
        self.input_path.set("")
        self.output_path.set("")
        self.status_text.set("")
    
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