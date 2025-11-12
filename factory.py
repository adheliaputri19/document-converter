# factory.py
from conversion.engine import ConversionEngine
from ui.gui_manager import GUIManager
from cli.cli_converter import CLIConverter


class ConverterFactory:
    @staticmethod
    def create_gui_converter():
        return GUIManager()
    
    @staticmethod
    def create_cli_converter():
        return CLIConverter()
    
    @staticmethod
    def create_conversion_engine(has_ms_word: bool = None):
        return ConversionEngine(has_ms_word)