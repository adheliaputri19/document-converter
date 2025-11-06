# conversion/engine.py
import os
import sys
from .strategies import DocToPdfStrategy, PdfToDocxStrategy, PdfToDocStrategy


class ConversionEngine:
    """Main engine untuk menangani semua jenis konversi dokumen"""
    
    def __init__(self, has_ms_word: bool = None):
        """
        Initialize ConversionEngine
        
        Args:
            has_ms_word: None untuk auto-detection, True/False untuk manual override
        """
        self.has_ms_word = self._detect_ms_word() if has_ms_word is None else has_ms_word
        self._setup_strategies()
        self._print_initialization_info()
    
    def _detect_ms_word(self) -> bool:
        """Auto-detect Microsoft Word installation"""
        print("Mendeteksi instalasi Microsoft Word...")
        try:
            import comtypes.client
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.Quit()
            print("✓ Microsoft Word terdeteksi")
            return True
        except Exception as e:
            print(f"✗ Microsoft Word tidak terdeteksi: {e}")
            return False
    
    def _print_initialization_info(self):
        """Print informasi tentang status engine"""
        print("\n" + "="*50)
        print("DOCUMENT CONVERSION ENGINE")
        print("="*50)
        print(f"Microsoft Word: {'TERDETEKSI' if self.has_ms_word else 'TIDAK TERDETEKSI'}")
        print(f"DOC/DOCX to PDF: {'AVAILABLE' if self.has_ms_word else 'DOCX ONLY'}")
        print(f"PDF to DOCX: AVAILABLE")
        print(f"PDF to DOC: {'AVAILABLE' if self.has_ms_word else 'NOT AVAILABLE'}")
        print("="*50 + "\n")
    
    def _setup_strategies(self):
        """Setup semua strategy yang tersedia"""
        self.strategies = {
            'doc_to_pdf': DocToPdfStrategy(self.has_ms_word),
            'pdf_to_docx': PdfToDocxStrategy(),
            'pdf_to_doc': PdfToDocStrategy(PdfToDocxStrategy(), self.has_ms_word)
        }
    
    def convert(self, conversion_type: str, input_file: str, output_file: str, **kwargs) -> bool:
        """
        Eksekusi konversi berdasarkan tipe
        
        Args:
            conversion_type: Jenis konversi ('doc_to_pdf', 'pdf_to_docx', 'pdf_to_doc')
            input_file: Path file input
            output_file: Path file output
            **kwargs: Additional parameters untuk strategy
        
        Returns:
            bool: True jika konversi berhasil
            
        Raises:
            ValueError: Jika tipe konversi tidak didukung
            Exception: Jika konversi gagal
        """
        if conversion_type not in self.strategies:
            supported = list(self.strategies.keys())
            raise ValueError(f"Tipe konversi tidak didukung: {conversion_type}. "
                           f"Yang didukung: {supported}")
        
        print(f"\nMemulai konversi: {conversion_type}")
        print(f"Input: {input_file}")
        print(f"Output: {output_file}")
        
        strategy = self.strategies[conversion_type]
        
        # Update strategy parameters jika ada
        if conversion_type == 'pdf_to_docx' and 'conversion_method' in kwargs:
            strategy.conversion_method = kwargs['conversion_method']
            print(f"Metode konversi: {kwargs['conversion_method']}")
        
        try:
            result = strategy.convert(input_file, output_file)
            if result:
                print(f"✓ Konversi berhasil: {output_file}")
            return result
        except Exception as e:
            print(f"✗ Konversi gagal: {str(e)}")
            raise
    
    def get_supported_conversions(self) -> dict:
        """Mendapatkan daftar konversi yang didukung"""
        return {
            'doc_to_pdf': {
                'description': 'DOC/DOCX ke PDF',
                'input_extensions': ['.doc', '.docx'] if self.has_ms_word else ['.docx'],
                'output_extension': '.pdf',
                'available': True
            },
            'pdf_to_docx': {
                'description': 'PDF ke DOCX',
                'input_extensions': ['.pdf'],
                'output_extension': '.docx',
                'available': True
            },
            'pdf_to_doc': {
                'description': 'PDF ke DOC',
                'input_extensions': ['.pdf'] if self.has_ms_word else [],
                'output_extension': '.doc',
                'available': self.has_ms_word
            }
        }
    
    def print_supported_conversions(self):
        """Print daftar konversi yang didukung dalam format yang mudah dibaca"""
        conversions = self.get_supported_conversions()
        
        print("\nKONVERSI YANG DIDUKUNG:")
        print("-" * 40)
        for key, info in conversions.items():
            status = "✓ AVAILABLE" if info['available'] else "✗ NOT AVAILABLE"
            input_ext = ", ".join(info['input_extensions'])
            print(f"{key:15} {info['description']:20} {input_ext:15} → {info['output_extension']:10} {status}")
        print("-" * 40)
    
    def check_ms_word_installation(self) -> bool:
        """Cek apakah Microsoft Word terinstall"""
        return self.has_ms_word
    
    def get_engine_info(self) -> dict:
        """Dapatkan informasi lengkap tentang engine"""
        return {
            'has_ms_word': self.has_ms_word,
            'supported_conversions': self.get_supported_conversions(),
            'libraries_available': {
                'comtypes': self._check_library('comtypes.client'),
                'pymupdf': self._check_library('fitz'),
                'docx': self._check_library('docx'),
                'pdf2docx': self._check_library('pdf2docx'),
                'docx2pdf': self._check_library('docx2pdf')
            }
        }
    
    def _check_library(self, library_name: str) -> bool:
        """Cek apakah library tersedia"""
        try:
            __import__(library_name)
            return True
        except ImportError:
            return False


# Utility function untuk mudah digunakan
def create_conversion_engine(force_ms_word_detection: bool = False):
    """
    Utility function untuk membuat ConversionEngine
    
    Args:
        force_ms_word_detection: Jika True, akan memaksa deteksi ulang MS Word
        
    Returns:
        ConversionEngine instance
    """
    if force_ms_word_detection:
        return ConversionEngine(has_ms_word=None)  # Force auto-detection
    else:
        return ConversionEngine()