# main.py
import sys
import os

# Add current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def main():
    """Main entry point untuk aplikasi Document Converter"""
    
   if len(sys.argv) > 1:
    mode = sys.argv[1]

    if mode == "compress":
        # Fitur Kompres Dokumen
        from conversion.doc_compressor import DocumentCompressor
        compressor = DocumentCompressor()
        
        print("Masukkan path dokumen yang ingin dikompres (pisahkan dengan koma):")
        file_input = input("> ")
        files = [f.strip() for f in file_input.split(",")]

        output_name = input("Nama file zip (kosongkan untuk default): ").strip()
        if not output_name:
            output_name = "compressed_documents.zip"

        compressor.compress_documents(files, output_name)

    else:
        # Mode CLI Converter biasa
        from cli.cli_converter import CLIConverter
        converter = CLIConverter()
        converter.run()


if __name__ == "__main__":
    main()