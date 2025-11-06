import os
from pathlib import Path


class FileHandler:
    """Utility class untuk menangani operasi file"""
    
    @staticmethod
    def validate_file_exists(file_path: str) -> bool:
        """Validasi bahwa file exists"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File tidak ditemukan: {file_path}")
        return True
    
    @staticmethod
    def validate_file_extension(file_path: str, allowed_extensions: list) -> bool:
        """Validasi ekstensi file"""
        file_extension = Path(file_path).suffix.lower()
        if file_extension not in allowed_extensions:
            raise ValueError(f"Ekstensi file tidak didukung: {file_extension}. Harus: {allowed_extensions}")
        return True