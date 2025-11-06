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
    @staticmethod
    def validate_file_size(file_path: str) -> bool:
        """Validasi ukuran file"""
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            raise ValueError("File kosong (0 bytes)")
        return True
    
    @staticmethod
    def auto_generate_output_path(input_path: str, conversion_type: str) -> str:
        """Generate output path otomatis berdasarkan input path dan tipe konversi"""
        input_file = Path(input_path)
        
        conversion_map = {
            'doc_to_pdf': '.pdf',
            'pdf_to_docx': '.docx',
            'pdf_to_doc': '.doc'
        }
        
        if conversion_type not in conversion_map:
            raise ValueError(f"Tipe konversi tidak valid: {conversion_type}")
        
        output_extension = conversion_map[conversion_type]
        return str(input_file.with_suffix(output_extension))
    
    @staticmethod
    def get_file_info(file_path: str) -> dict:
        """Mendapatkan informasi file"""
        path = Path(file_path)
        return {
            'name': path.name,
            'size': os.path.getsize(file_path),
            'extension': path.suffix.lower(),
            'parent_dir': str(path.parent)
        }
    
    @staticmethod
    def safe_delete(file_path: str) -> bool:
        """Hapus file dengan safety check"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception:
            return False