# conversion/strategies.py
import os
import tempfile
import shutil
import sys
from abc import ABC, abstractmethod

try:
    from docx2pdf import convert as docx2pdf_convert
    LIBRARY_AVAILABLE = True
except ImportError:
    LIBRARY_AVAILABLE = False

try:
    import comtypes.client
    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Inches, Cm
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pdf2docx import Converter
    PDF2DOCX_CONVERTER_AVAILABLE = True
except ImportError:
    PDF2DOCX_CONVERTER_AVAILABLE = False


class ConversionStrategy(ABC):
    """Abstract base class for conversion strategies"""
    
    @abstractmethod
    def convert(self, input_file: str, output_file: str) -> bool:
        pass
    
    @abstractmethod
    def validate_input(self, input_file: str) -> bool:
        pass


class DocToPdfStrategy(ConversionStrategy):
    """Strategy for converting DOC/DOCX to PDF"""
    
    def __init__(self, has_ms_word: bool = False):
        self.has_ms_word = has_ms_word
        self._verified_ms_word = False
    
    def _verify_ms_word_installation(self) -> bool:
        """Verify Microsoft Word installation secara langsung"""
        if not COMTYPES_AVAILABLE:
            return False
            
        try:
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.Quit()
            return True
        except Exception as e:
            print(f"Peringatan: Microsoft Word tidak dapat diakses: {e}")
            return False
    
    def validate_input(self, input_file: str) -> bool:
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File tidak ditemukan: {input_file}")
        
        file_extension = os.path.splitext(input_file)[1].lower()
        if file_extension not in ['.doc', '.docx']:
            raise ValueError(f"Format file tidak didukung: {file_extension}")
        
        if file_extension == '.doc':
            if not self.has_ms_word:
                raise ValueError("Konversi .doc ke PDF membutuhkan Microsoft Word")
            
            # Verify MS Word on first use for .doc files
            if not self._verified_ms_word:
                if not self._verify_ms_word_installation():
                    raise ValueError("Microsoft Word tidak terdeteksi. Pastikan Word terinstall dan tersedia.")
                self._verified_ms_word = True
        
        return True
    
    def convert(self, input_file: str, output_file: str) -> bool:
        self.validate_input(input_file)
        
        file_extension = os.path.splitext(input_file)[1].lower()
        
        if file_extension == '.docx':
            if not LIBRARY_AVAILABLE:
                raise ImportError("Library docx2pdf tidak tersedia. Install dengan: pip install docx2pdf")
            
            print("Mengkonversi DOCX ke PDF menggunakan docx2pdf...")
            docx2pdf_convert(input_file, output_file)
            return True
            
        elif file_extension == '.doc':
            if self.has_ms_word and COMTYPES_AVAILABLE:
                return self._convert_doc_to_pdf_with_word(input_file, output_file)
            else:
                raise ValueError("Konversi .doc ke PDF membutuhkan Microsoft Word dan comtypes")
    
    def _convert_doc_to_pdf_with_word(self, input_file: str, output_file: str) -> bool:
        """Convert .doc to PDF menggunakan Microsoft Word"""
        try:
            input_file = os.path.abspath(input_file)
            output_file = os.path.abspath(output_file)
            
            print("Membuka Microsoft Word untuk konversi DOC ke PDF...")
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            try:
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=17)  # 17 = PDF format
                doc.Close()
                print("Konversi DOC ke PDF berhasil!")
                return True
            except Exception as e:
                # Ensure document is closed even if error occurs
                try:
                    doc.Close()
                except:
                    pass
                raise e
            finally:
                word.Quit()
                
        except Exception as e:
            raise Exception(f"Gagal mengkonversi .doc ke PDF menggunakan Word: {str(e)}")


class PdfToDocxStrategy(ConversionStrategy):
    """Strategy for converting PDF to DOCX with multiple methods"""
    
    def __init__(self, conversion_method: str = "auto"):
        self.conversion_method = conversion_method
    
    def validate_input(self, input_file: str) -> bool:
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File PDF tidak ditemukan: {input_file}")
        
        if not input_file.lower().endswith('.pdf'):
            raise ValueError("File input harus berformat PDF")
        
        file_size = os.path.getsize(input_file)
        if file_size == 0:
            raise ValueError("File PDF kosong (0 bytes)")
        
        # Basic PDF validation
        try:
            if PYMUPDF_AVAILABLE:
                with fitz.open(input_file) as doc:
                    if len(doc) == 0:
                        raise ValueError("PDF tidak memiliki halaman")
        except Exception as e:
            raise ValueError(f"File PDF tidak valid: {str(e)}")
        
        return True
    
    def convert(self, input_file: str, output_file: str) -> bool:
        self.validate_input(input_file)
        
        print(f"Memulai konversi PDF ke DOCX dengan metode: {self.conversion_method}")
        
        # METHOD 1: pdf2docx (Best for formatting + images)
        if self.conversion_method in ["auto", "pdf2docx"] and PDF2DOCX_CONVERTER_AVAILABLE:
            try:
                print("Mencoba konversi dengan pdf2docx...")
                cv = Converter(input_file)
                cv.convert(output_file, start=0, end=None)
                cv.close()
                
                if self._validate_docx_output(output_file):
                    print("Konversi dengan pdf2docx berhasil!")
                    return True
                else:
                    print("Output pdf2docx tidak valid, mencoba metode lain...")
            except Exception as e:
                print(f"Metode pdf2docx gagal: {e}")
                if self.conversion_method == "pdf2docx":
                    raise
        
        # METHOD 2: PyMuPDF with images
        if self.conversion_method in ["auto", "pymupdf"] and PYMUPDF_AVAILABLE and DOCX_AVAILABLE:
            try:
                print("Mencoba konversi dengan PyMuPDF (dengan gambar)...")
                return self._convert_pdf_to_docx_pymupdf_with_images(input_file, output_file)
            except Exception as e:
                print(f"Metode PyMuPDF dengan gambar gagal: {e}")
                if self.conversion_method == "pymupdf":
                    raise
        
        # METHOD 3: Simple text extraction (fallback)
        if self.conversion_method in ["auto", "text_only"] and PYMUPDF_AVAILABLE and DOCX_AVAILABLE:
            try:
                print("Mencoba konversi dengan PyMuPDF (text only)...")
                return self._convert_pdf_to_docx_text_only(input_file, output_file)
            except Exception as e:
                print(f"Metode text-only gagal: {e}")
                if self.conversion_method == "text_only":
                    raise
        
        raise Exception("Semua metode konversi gagal. Coba install pdf2docx untuk hasil terbaik: pip install pdf2docx")
    
    def _convert_pdf_to_docx_pymupdf_with_images(self, input_file: str, output_file: str) -> bool:
        """Convert PDF to DOCX dengan PyMuPDF termasuk gambar"""
        temp_dir = tempfile.mkdtemp()
        
        try:
            pdf_document = fitz.open(input_file)
            doc = Document()
            
            # Set page layout
            section = doc.sections[0]
            section.page_height = Cm(29.7)
            section.page_width = Cm(21.0)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                
                # Extract text
                text = page.get_text()
                
                # Extract images
                image_list = page.get_images()
                
                if image_list:
                    print(f"Halaman {page_num + 1}: Menemukan {len(image_list)} gambar")
                    for img_index, img in enumerate(image_list):
                        try:
                            xref = img[0]
                            pix = fitz.Pixmap(pdf_document, xref)
                            
                            if pix.n - pix.alpha < 4:
                                if pix.alpha:
                                    pix = fitz.Pixmap(fitz.csRGB, pix)
                                
                                img_path = os.path.join(temp_dir, f"page{page_num+1}_img{img_index+1}.png")
                                pix.save(img_path)
                                doc.add_picture(img_path, width=Inches(6.0))
                                doc.add_paragraph()
                            
                            pix = None
                        except Exception as e:
                            print(f"Gagal mengekstrak gambar {img_index + 1}: {e}")
                            continue
                
                if text.strip():
                    lines = text.split('\n')
                    for line in lines:
                        if line.strip():
                            doc.add_paragraph(line.strip())
                    
                    if page_num < len(pdf_document) - 1:
                        doc.add_page_break()
            
            doc.save(output_file)
            pdf_document.close()
            
            if self._validate_docx_output(output_file):
                print("Konversi dengan PyMuPDF (dengan gambar) berhasil!")
                return True
            else:
                raise Exception("Output DOCX tidak valid")
                
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _convert_pdf_to_docx_text_only(self, input_file: str, output_file: str) -> bool:
        """Convert PDF to DOCX text only (simple method)"""
        pdf_document = fitz.open(input_file)
        doc = Document()
        
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            
            if text.strip():
                lines = text.split('\n')
                for line in lines:
                    if line.strip():
                        doc.add_paragraph(line.strip())
                
                if page_num < len(pdf_document) - 1:
                    doc.add_page_break()
        
        doc.save(output_file)
        pdf_document.close()
        
        if self._validate_docx_output(output_file):
            print("Konversi dengan PyMuPDF (text only) berhasil!")
            return True
        else:
            raise Exception("Output DOCX tidak valid")
    
    def _validate_docx_output(self, file_path: str) -> bool:
        """Validasi file DOCX output"""
        if not os.path.exists(file_path):
            return False
        
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            return False
            
        try:
            import zipfile
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                return True
        except:
            return False


class PdfToDocStrategy(ConversionStrategy):
    """Strategy for converting PDF to DOC using Microsoft Word"""
    
    def __init__(self, pdf_to_docx_strategy: PdfToDocxStrategy, has_ms_word: bool = False):
        self.pdf_to_docx_strategy = pdf_to_docx_strategy
        self.has_ms_word = has_ms_word
        self._verified_ms_word = False
    
    def _verify_ms_word_installation(self) -> bool:
        """Verify Microsoft Word installation secara langsung"""
        if not COMTYPES_AVAILABLE:
            return False
            
        try:
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.Quit()
            return True
        except Exception as e:
            print(f"Error deteksi Microsoft Word: {e}")
            return False
    
    def validate_input(self, input_file: str) -> bool:
        if not self.has_ms_word:
            raise Exception("Konversi PDF ke DOC membutuhkan Microsoft Word")
        
        # Verify MS Word installation on first use
        if not self._verified_ms_word:
            print("Memverifikasi instalasi Microsoft Word...")
            if not self._verify_ms_word_installation():
                raise Exception("Microsoft Word tidak terdeteksi. Pastikan:\n"
                              "1. Microsoft Word terinstall\n"
                              "2. Word tersedia di PATH system\n"
                              "3. COM automation diaktifkan")
            self._verified_ms_word = True
            print("Microsoft Word terverifikasi!")
        
        if not input_file.lower().endswith('.pdf'):
            raise ValueError("File input harus berformat PDF")
        
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File PDF tidak ditemukan: {input_file}")
        
        file_size = os.path.getsize(input_file)
        if file_size == 0:
            raise ValueError("File PDF kosong (0 bytes)")
        
        return True
    
    def convert(self, input_file: str, output_file: str) -> bool:
        self.validate_input(input_file)
        
        temp_docx = tempfile.NamedTemporaryFile(suffix='.docx', delete=False).name
        
        try:
            # Convert PDF to DOCX
            print("Langkah 1: Mengkonversi PDF ke DOCX terlebih dahulu...")
            self.pdf_to_docx_strategy.convert(input_file, temp_docx)
            
            if not os.path.exists(temp_docx) or os.path.getsize(temp_docx) == 0:
                raise Exception("Konversi PDF ke DOCX gagal menghasilkan file yang valid")
            
            # Convert DOCX to DOC menggunakan Microsoft Word
            print("Langkah 2: Mengkonversi DOCX ke DOC menggunakan Microsoft Word...")
            success = self._convert_docx_to_doc_with_word(temp_docx, output_file)
            
            if success:
                print("Konversi PDF ke DOC berhasil!")
                return True
            else:
                raise Exception("Konversi DOCX ke DOC gagal")
            
        except Exception as e:
            # Clean up temp file if error occurs
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass
            raise Exception(f"Gagal mengkonversi PDF ke DOC: {str(e)}")
        finally:
            # Clean up temp file
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass
    
    def _convert_docx_to_doc_with_word(self, input_file: str, output_file: str) -> bool:
        """Convert DOCX to DOC menggunakan Microsoft Word"""
        try:
            input_file = os.path.abspath(input_file)
            output_file = os.path.abspath(output_file)
            
            print(f"Membuka Microsoft Word untuk konversi: {input_file} -> {output_file}")
            
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            try:
                # Open DOCX file
                doc = word.Documents.Open(input_file)
                
                # Save as DOC format (FileFormat=0 untuk .doc)
                doc.SaveAs(output_file, FileFormat=0)
                doc.Close()
                
                # Verify output file
                if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                    print("Konversi DOCX ke DOC selesai")
                    return True
                else:
                    raise Exception("File output tidak terbentuk atau kosong")
                
            except Exception as e:
                # Ensure document is closed even if error occurs
                try:
                    doc.Close()
                except:
                    pass
                raise e
                
            finally:
                word.Quit()
                
        except Exception as e:
            raise Exception(f"Gagal mengkonversi DOCX ke DOC menggunakan Word: {str(e)}")