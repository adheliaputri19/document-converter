import argparse
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from conversion.engine import create_conversion_engine


class CLIConverter:
    """CLI interface untuk Document Converter"""
    
    def __init__(self):
        self.engine = create_conversion_engine()
        self.parser = self._setup_parser()
    
    def _setup_parser(self):
        """Setup argument parser"""
        parser = argparse.ArgumentParser(
            description='Document Converter - Konversi dokumen antara PDF, DOC, dan DOCX',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Contoh penggunaan:
  python main.py doc-to-pdf input.docx output.pdf
  python main.py pdf-to-docx input.pdf output.docx
  python main.py pdf-to-doc input.pdf output.doc
  python main.py --list-supported

Untuk GUI mode, jalankan tanpa argument:
  python main.py
            """
        )
        
        # Subcommands
        subparsers = parser.add_subparsers(dest='command', help='Jenis konversi')
        
        # DOC to PDF
        doc_parser = subparsers.add_parser('doc-to-pdf', help='Konversi DOC/DOCX ke PDF')
        doc_parser.add_argument('input', help='File input (DOC/DOCX)')
        doc_parser.add_argument('output', help='File output (PDF)')
        
        # PDF to DOCX
        pdf_docx_parser = subparsers.add_parser('pdf-to-docx', help='Konversi PDF ke DOCX')
        pdf_docx_parser.add_argument('input', help='File input (PDF)')
        pdf_docx_parser.add_argument('output', help='File output (DOCX)')
        pdf_docx_parser.add_argument('--method', choices=['auto', 'pdf2docx', 'pymupdf', 'text_only'],
                                   default='auto', help='Metode konversi (default: auto)')
        
        # PDF to DOC
        pdf_doc_parser = subparsers.add_parser('pdf-to-doc', help='Konversi PDF ke DOC')
        pdf_doc_parser.add_argument('input', help='File input (PDF)')
        pdf_doc_parser.add_argument('output', help='File output (DOC)')
        
        # Global options
        parser.add_argument('--list-supported', action='store_true',
                          help='Tampilkan daftar konversi yang didukung')
        parser.add_argument('--force', action='store_true',
                          help='Force overwrite file output jika sudah ada')
        parser.add_argument('--verbose', action='store_true',
                          help='Tampilkan informasi detail')
        
        return parser
    
    def _check_files(self, input_file: str, output_file: str, force: bool = False) -> bool:
        """Validasi file input dan output"""
        # Check input file
        if not os.path.exists(input_file):
            print(f"❌ Error: File input tidak ditemukan: {input_file}")
            return False
        
        if os.path.getsize(input_file) == 0:
            print(f"❌ Error: File input kosong: {input_file}")
            return False
        
        # Check output file
        if os.path.exists(output_file) and not force:
            print(f"❌ Error: File output sudah ada: {output_file}")
            print("   Gunakan --force untuk overwrite")
            return False
        
        # Check output directory
        output_dir = os.path.dirname(output_file) or '.'
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"📁 Membuat directory: {output_dir}")
            except Exception as e:
                print(f"❌ Error: Tidak dapat membuat directory output: {e}")
                return False
        
        return True
    
    def _get_conversion_type(self, command: str) -> str:
        """Map command ke conversion type"""
        conversion_map = {
            'doc-to-pdf': 'doc_to_pdf',
            'pdf-to-docx': 'pdf_to_docx',
            'pdf-to-doc': 'pdf_to_doc'
        }
        return conversion_map.get(command)
    
    def run(self):
        """Jalankan CLI converter"""
        args = self.parser.parse_args()
        
        # Handle list-supported command
        if args.list_supported or not args.command:
            self.engine.print_supported_conversions()
            
            # Print engine info jika verbose
            if args.verbose:
                print("\n🔧 ENGINE INFO:")
                info = self.engine.get_engine_info()
                for lib, available in info['libraries_available'].items():
                    status = "✓" if available else "✗"
                    print(f"   {lib:15} {status}")
            return
        
        # Validasi command
        conversion_type = self._get_conversion_type(args.command)
        if not conversion_type:
            print(f"❌ Error: Command tidak valid: {args.command}")
            self.parser.print_help()
            return
        
        # Validasi file
        if not self._check_files(args.input, args.output, args.force):
            return
        
        # Jalankan konversi
        try:
            print(f"🔄 Memulai konversi: {args.input} → {args.output}")
            start_time = time.time()
            
            # Prepare kwargs
            kwargs = {}
            if args.command == 'pdf-to-docx':
                kwargs['conversion_method'] = args.method
            
            # Execute conversion
            success = self.engine.convert(
                conversion_type=conversion_type,
                input_file=args.input,
                output_file=args.output,
                **kwargs
            )
            
            if success:
                end_time = time.time()
                file_size = os.path.getsize(args.output)
                print(f"✅ Konversi berhasil!")
                print(f"📁 Output: {args.output}")
                print(f"📊 Size: {file_size:,} bytes")
                print(f"⏱️  Waktu: {end_time - start_time:.2f} detik")
            else:
                print(f"❌ Konversi gagal")
                
        except Exception as e:
            print(f"❌ Error selama konversi: {str(e)}")
            
            # Additional help untuk error spesifik
            if "Microsoft Word" in str(e):
                print("\n💡 Tips untuk Microsoft Word:")
                print("   - Pastikan Microsoft Word terinstall")
                print("   - Jalankan sebagai administrator jika perlu")
                print("   - Coba buka Word manual sekali untuk aktivasi")
            
            if "library" in str(e).lower() or "import" in str(e).lower():
                print("\n💡 Tips untuk dependencies:")
                print("   - Install required libraries: pip install pymupdf python-docx pdf2docx docx2pdf comtypes")


def main():
    """Entry point untuk CLI"""
    converter = CLIConverter()
    converter.run()


if __name__ == "__main__":
    main()