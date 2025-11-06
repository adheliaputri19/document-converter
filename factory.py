# factory.py
from conversion.engine import ConversionEngine
from ui.gui_manager import GUIManager


class ConverterFactory:
    """Factory class untuk membuat instance converter"""
    
    @staticmethod
    def create_gui_converter():
        """Buat GUI converter instance"""
        return GUIManager()
    
    @staticmethod
    def create_cli_converter():
        """Buat CLI converter instance"""
        from cli.cli_converter import CLIConverter
        return CLIConverter()
    
    @staticmethod
    def create_conversion_engine(has_ms_word: bool = False):
        """Buat conversion engine instance"""
        return ConversionEngine(has_ms_word)
    
    @staticmethod
    def get_supported_conversions():
        """Dapatkan daftar konversi yang didukung"""
        engine = ConversionEngine()
        return engine.get_supported_conversions()
    
    @staticmethod
    def check_system_dependencies():
        """Cek dependencies sistem"""
        engine = ConversionEngine()
        
        dependencies = {
            'Microsoft Word': engine.check_ms_word_installation(),
            'docx2pdf': True,  # Akan di-check saat runtime
            'PyMuPDF': True,   # Akan di-check saat runtime
            'python-docx': True,  # Akan di-check saat runtime
            'pdf2docx': True,  # Akan di-check saat runtime
            'comtypes': True   # Akan di-check saat runtime
        }
        
        return dependencies