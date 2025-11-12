# cli/cli_converter.py
import argparse
import os
import sys
import time
from pathlib import Path

# Fix path agar bisa import dari root
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(ROOT_DIR)

from conversion.engine import create_conversion_engine
from conversion.compressor import DocumentCompressor
from utils.file_handler import FileHandler


class CLIConverter:
    def __init__(self):
        self.engine = create_conversion_engine()
        self.compressor = DocumentCompressor()
        self.parser = self._setup_parser()

    def _setup_parser(self):
        parser = argparse.ArgumentParser(
            description="Document Converter & Compressor",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Contoh:
  python main.py doc-to-pdf input.docx output.pdf
  python main.py pdf-to-docx input.pdf output.docx --method pdf2docx
  python main.py pdf-to-doc input.pdf output.doc
  python main.py compress file.pdf kecil.pdf --level high
  python main.py compress-folder ./data/ ./output/ --force
  python main.py list-supported
            """
        )
        sub = parser.add_subparsers(dest='command')

        # Konversi
        p = sub.add_parser('doc-to-pdf', help='DOC/DOCX → PDF')
        p.add_argument('input', help='File input')
        p.add_argument('output', help='File output')

        p = sub.add_parser('pdf-to-docx', help='PDF → DOCX')
        p.add_argument('input', help='File input')
        p.add_argument('output', help='File output')
        p.add_argument('--method', choices=['auto', 'pdf2docx', 'pymupdf', 'text_only'], default='auto')

        p = sub.add_parser('pdf-to-doc', help='PDF → DOC (butuh MS Word)')
        p.add_argument('input', help='File input')
        p.add_argument('output', help='File output')

        # Kompres
        p = sub.add_parser('compress', help='Kompres file')
        p.add_argument('input', help='File input')
        p.add_argument('output', help='File output')
        p.add_argument('--level', choices=['low', 'medium', 'high'], default='medium')

        p = sub.add_parser('compress-folder', help='Kompres folder')
        p.add_argument('input_folder', help='Folder input')
        p.add_argument('output_folder', help='Folder output')
        p.add_argument('--level', choices=['low', 'medium', 'high'], default='medium')
        p.add_argument('--force', action='store_true')

        sub.add_parser('list-supported', help='Lihat konversi yang didukung')

        return parser

    def run(self):
        args = self.parser.parse_args()
        if not args.command:
            self.parser.print_help()
            return

        try:
            if args.command == 'doc-to-pdf':
                self.engine.convert('doc_to_pdf', args.input, args.output)
                print("Konversi selesai!")

            elif args.command == 'pdf-to-docx':
                self.engine.convert('pdf_to_docx', args.input, args.output, method=args.method)
                print("Konversi selesai!")

            elif args.command == 'pdf-to-doc':
                self.engine.convert('pdf_to_doc', args.input, args.output)
                print("Konversi selesai!")

            elif args.command == 'compress':
                start = time.time()
                self.compressor.compress(args.input, args.output, args.level)
                size_in = os.path.getsize(args.input)
                size_out = os.path.getsize(args.output)
                reduction = 100 * (1 - size_out / size_in)
                print(f"Kompres selesai! -{reduction:.1f}% | {time.time()-start:.2f}s")

            elif args.command == 'compress-folder':
                self._compress_folder(args)

            elif args.command == 'list-supported':
                self.engine.print_supported_conversions()

        except Exception as e:
            print(f"GAGAL: {e}")

    def _compress_folder(self, args):
        in_dir = Path(args.input_folder)
        out_dir = Path(args.output_folder)
        if not in_dir.is_dir():
            print("Folder input tidak ada!")
            return
        out_dir.mkdir(parents=True, exist_ok=True)

        files = [f for f in in_dir.rglob('*') if f.suffix.lower() in {'.pdf', '.docx'}]
        if not files:
            print("Tidak ada file PDF/DOCX.")
            return

        print(f"Kompres {len(files)} file...")
        for i, f in enumerate(files, 1):  # Fixed typo: enumerate51 → enumerate(files, 1)
            rel = f.relative_to(in_dir)
            out = out_dir / f"compressed_{rel.name}"
            out.parent.mkdir(parents=True, exist_ok=True)
            if not args.force and out.exists():
                print(f"  [SKIP] {out}")
                continue
            try:
                self.compressor.compress(str(f), str(out), args.level)
                print(f"  [{i}] {rel} → {out.name}")
            except Exception as e:
                print(f"  [ERROR] {f}: {e}")

        print("SELESAI!")


if __name__ == "__main__":
    CLIConverter().run()