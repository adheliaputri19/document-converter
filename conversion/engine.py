# conversion/engine.py
import os
from .strategies import DocToPdfStrategy, PdfToDocxStrategy, PdfToDocStrategy


class ConversionEngine:
    def __init__(self, has_ms_word: bool = None):
        self.has_ms_word = self._detect_ms_word() if has_ms_word is None else has_ms_word
        self._setup_strategies()
        self._print_initialization_info()
    
    def _detect_ms_word(self) -> bool:
        print("Mendeteksi Microsoft Word...")
        try:
            import comtypes.client
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.Quit()
            print("Microsoft Word terdeteksi")
            return True
        except Exception as e:
            print(f"Microsoft Word tidak terdeteksi: {e}")
            return False
    
    def _print_initialization_info(self):
        print("\n" + "="*50)
        print("DOCUMENT CONVERSION ENGINE")
        print("="*50)
        print(f"MS Word: {'TERDETEKSI' if self.has_ms_word else 'TIDAK'}")
        print(f"DOC/DOCX → PDF: {'OK' if self.has_ms_word else 'DOCX ONLY'}")
        print(f"PDF → DOCX: OK")
        print(f"PDF → DOC: {'OK' if self.has_ms_word else 'TIDAK'}")
        print("="*50 + "\n")
    
    def _setup_strategies(self):
        self.strategies = {
            'doc_to_pdf': DocToPdfStrategy(self.has_ms_word),
            'pdf_to_docx': PdfToDocxStrategy(),
            'pdf_to_doc': PdfToDocStrategy(PdfToDocxStrategy(), self.has_ms_word)
        }
    
    def convert(self, conversion_type: str, input_file: str, output_file: str, **kwargs) -> bool:
        if conversion_type not in self.strategies:
            raise ValueError(f"Tipe tidak didukung: {conversion_type}")
        
        strategy = self.strategies[conversion_type]
        if conversion_type == 'pdf_to_docx' and 'method' in kwargs:
            strategy = PdfToDocxStrategy(kwargs['method'])
        
        try:
            print(f"Konversi: {input_file} → {output_file}")
            return strategy.convert(input_file, output_file)
        except Exception as e:
            print(f"Gagal: {e}")
            return False
    
    def get_supported_conversions(self) -> dict:
        return {
            'doc_to_pdf': {
                'desc': 'DOC/DOCX → PDF',
                'in': ['.docx', '.doc'] if self.has_ms_word else ['.docx'],
                'out': '.pdf',
                'ok': True
            },
            'pdf_to_docx': {
                'desc': 'PDF → DOCX',
                'in': ['.pdf'],
                'out': '.docx',
                'ok': True
            },
            'pdf_to_doc': {
                'desc': 'PDF → DOC',
                'in': ['.pdf'] if self.has_ms_word else [],
                'out': '.doc',
                'ok': self.has_ms_word
            }
        }
    
    def print_supported_conversions(self):
        print("\nKONVERSI YANG DIDUKUNG:")
        print("-" * 45)
        for k, v in self.get_supported_conversions().items():
            status = "OK" if v['ok'] else "TIDAK"
            print(f"{k:15} {v['desc']:20} {', '.join(v['in']):15} → {v['out']:10} {status}")
        print("-" * 45)
    
    def check_ms_word_installation(self) -> bool:
        return self.has_ms_word


def create_conversion_engine(force: bool = False):
    return ConversionEngine(has_ms_word=None if force else None)