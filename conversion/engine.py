# conversion/engine.py
from .strategies import DocToPdfStrategy, PdfToDocxStrategy, PdfToDocStrategy


class ConversionEngine:
    """Main engine untuk menangani semua jenis konversi dokumen"""
    
    def __init__(self, has_ms_word: bool = False):
        self.has_ms_word = has_ms_word
        self._setup_strategies()
    
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
        """
        if conversion_type not in self.strategies:
            raise ValueError(f"Tipe konversi tidak didukung: {conversion_type}")
        
        strategy = self.strategies[conversion_type]
        
        # Update strategy parameters jika ada
        if conversion_type == 'pdf_to_docx' and 'conversion_method' in kwargs:
            strategy.conversion_method = kwargs['conversion_method']
        
        return strategy.convert(input_file, output_file)
    
    def get_supported_conversions(self) -> dict:
        """Mendapatkan daftar konversi yang didukung"""
        return {
            'doc_to_pdf': {
                'description': 'DOC/DOCX ke PDF',
                'input_extensions': ['.doc', '.docx'] if self.has_ms_word else ['.docx'],
                'output_extension': '.pdf'
            },
            'pdf_to_docx': {
                'description': 'PDF ke DOCX',
                'input_extensions': ['.pdf'],
                'output_extension': '.docx'
            },
            'pdf_to_doc': {
                'description': 'PDF ke DOC',
                'input_extensions': ['.pdf'] if self.has_ms_word else [],
                'output_extension': '.doc'
            }
        }
    
    def check_ms_word_installation(self) -> bool:
        """Cek apakah Microsoft Word terinstall"""
        if not self.has_ms_word:
            return False
            
        try:
            import comtypes.client
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            word.Quit()
            return True
        except Exception:
            return False