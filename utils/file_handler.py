# utils/file_handler.py
import os
from pathlib import Path
from typing import List, Dict


class FileHandler:
    @staticmethod
    def validate_file_exists(file_path: str) -> Path:
        path = Path(file_path)
        if not path.is_file():
            raise FileNotFoundError(f"File tidak ditemukan: {file_path}")
        return path

    @staticmethod
    def validate_file_extension(file_path: str, allowed: List[str]) -> None:
        ext = Path(file_path).suffix.lower()
        allowed = [a.lower() if a.startswith('.') else f".{a.lower()}" for a in allowed]
        if ext not in allowed:
            raise ValueError(f"Ekstensi tidak didukung: {ext}. Harus: {', '.join(allowed)}")

    @staticmethod
    def validate_file_size(file_path: str, min_size: int = 1) -> None:
        if os.path.getsize(file_path) < min_size:
            raise ValueError("File kosong atau terlalu kecil")

    @staticmethod
    def auto_generate_output_path(input_path: str, conv_type: str) -> str:
        input_file = FileHandler.validate_file_exists(input_path)
        mapping = {
            'doc_to_pdf': '.pdf',
            'pdf_to_docx': '.docx',
            'pdf_to_doc': '.doc'
        }
        if conv_type not in mapping:
            raise ValueError(f"Tipe konversi tidak valid: {conv_type}")
        return str(input_file.with_suffix(mapping[conv_type]))

    @staticmethod
    def get_file_info(file_path: str) -> Dict:
        path = FileHandler.validate_file_exists(file_path)
        return {
            'name': path.name,
            'size': path.stat().st_size,
            'extension': path.suffix.lower(),
            'parent_dir': str(path.parent)
        }

    @staticmethod
    def safe_delete(file_path: str) -> bool:
        try:
            Path(file_path).unlink(missing_ok=True)
            return True
        except:
            return False