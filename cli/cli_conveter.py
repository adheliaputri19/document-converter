# cli/cli_converter.py
import argparse
import sys
import os
from pathlib import Path

from conversion.engine import ConversionEngine
from utils.file_handler import FileHandler


class CLIConverter:
    """Command Line Interface untuk Document Converter"""
    
    def __init__(self):
        self.conversion_engine = ConversionEngine()
        self.has_ms_word = self.conversion_engine.check_ms_word_installation()
        self.conversion_engine.has_ms_word = self.has_ms_word
    
    def setup_parser(self) -> argparse.ArgumentParser:
        """Setup argument parser"""
        parser = argparse.ArgumentParser(
            description="Document Converter - CLI Tool untuk konversi DOC/DOCX ↔ PDF",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Contoh penggunaan:
  python cli_converter.py -i input.docx -o output.pdf -t doc_to_pdf
  python cli_converter.py -i document.pdf -o output.docx -t pdf_to_docx --method pdf2docx
  python cli_converter.py -i file.pdf -o document.doc -t pdf_to_doc
            
Supported conversions:
  • doc_to_pdf    : Convert DOC/DOCX ke PDF
  • pdf_to_docx   : Convert PDF ke DOCX
  • pdf_to_doc    : Convert PDF ke DOC (membutuhkan MS Word)
            """
        )
        
        parser.add_argument('-i', '--input', required=True, help='File input')
        parser.add_argument('-o', '--output', required=True, help='File output')
        parser.add_argument('-t', '--type', required=True, 
                          choices=['doc_to_pdf', 'pdf_to_docx', 'pdf_to_doc'],
                          help='Tipe konversi')
        parser.add_argument('-m', '--method', default='auto',
                          choices=['auto', 'pdf2docx', 'pymupdf', 'text_only'],
                          help='Metode konversi untuk PDF ke DOCX')
        parser.add_argument('-y', '--yes', action='store_true',
                          help='Auto confirm overwrite')
        parser.add_argument('-v', '--verbose', action='store_true',
                          help='Verbose output')
        
        return parser
    
    def validate_arguments(self, args: argparse.Namespace) -> bool:
        """Validasi arguments"""
        # Check input file
        if not os.path.exists(args.input):
            print(f"❌ ERROR: File input tidak ditemukan: {args.input}")
            return False
        
        # Check output directory
        output_dir = Path(args.output).parent
        if not output_dir.exists():
            print(f"❌ ERROR: Directory output tidak ditemukan: {output_dir}")
            return False
        
        # Check file overwrite
        if os.path.exists(args.output) and not args.yes:
            response = input(f"⚠️  File {args.output} sudah ada. Overwrite? (y/N): ")
            if response.lower() != 'y':
                print("❌ Konversi dibatalkan")
                return False
        
        # Check MS Word requirement
        if args.type == 'pdf_to_doc' and not self.has_ms_word:
            print("❌ ERROR: Konversi PDF ke DOC membutuhkan Microsoft Word")
            print("   Gunakan opsi 'pdf_to_docx' sebagai alternatif")
            return False
        
        return True
    
    def print_system_info(self):
        """Print informasi sistem"""
        print("🤖 Document Converter - CLI Mode")
        print("=" * 50)
        
        supported = self.conversion_engine.get_supported_conversions()
        for conv_type, info in supported.items():
            status = "✅" if info['input_extensions'] else "❌"
            print(f"{status} {info['description']}")
        
        print("=" * 50)
    
    def convert(self, args: argparse.Namespace) -> bool:
        """Eksekusi konversi"""
        try:
            if args.verbose:
                self.print_system_info()
                print(f"🔄 Memulai konversi: {args.input} → {args.output}")
            
            # Validasi arguments
            if not self.validate_arguments(args):
                return False
            
            # Eksekusi konversi
            success = self.conversion_engine.convert(
                conversion_type=args.type,
                input_file=args.input,
                output_file=args.output,
                conversion_method=args.method
            )
            
            if success:
                if args.verbose:
                    file_info = FileHandler.get_file_info(args.output)
                    print(f"✅ Konversi berhasil!")
                    print(f"   Output: {args.output}")
                    print(f"   Size: {file_info['size']} bytes")
                else:
                    print(f"✅ Konversi berhasil: {args.output}")
                return True
            else:
                print(f"❌ Konversi gagal")
                return False
                
        except Exception as e:
            print(f"❌ ERROR: {str(e)}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            return False
    
    def run(self):
        """Jalankan CLI"""
        parser = self.setup_parser()
        args = parser.parse_args()
        
        success = self.convert(args)
        sys.exit(0 if success else 1)


def main():
    """Entry point untuk CLI"""
    converter = CLIConverter()
    converter.run()


if __name__ == "__main__":
    main()