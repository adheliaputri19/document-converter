import zipfile
import os

class DocumentCompressor:
    """Kelas untuk mengompres satu atau beberapa dokumen ke format ZIP."""

    def compress_documents(self, file_paths, output_zip="compressed_documents.zip"):
        """
        Mengompres daftar dokumen ke satu file ZIP.
        :param file_paths: list path file yang ingin dikompres
        :param output_zip: nama file ZIP output
        """
        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in file_paths:
                if os.path.exists(file_path):
                    zipf.write(file_path, os.path.basename(file_path))
                else:
                    print(f"[!] File tidak ditemukan: {file_path}")
        print(f"[OK] Dokumen berhasil dikompres ke: {output_zip}")
