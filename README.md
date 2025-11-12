# Document Converter & Compressor

Alat **konversi dokumen + kompresi ukuran file** yang powerful, serbaguna, dan mudah digunakan. Mendukung **DOC, DOCX, PDF** dengan **GUI** dan **CLI**.

--- 

## Fitur Utama

### Konversi Dokumen
- **DOC/DOCX → PDF**  
- **PDF → DOCX** (dengan gambar & formatting)  
- **PDF → DOC** *(butuh Microsoft Word)*

### Metode Konversi PDF → DOCX
| Metode      | Kelebihan                  | Kekurangan          |
|-------------|----------------------------|---------------------|
| `auto`      | Pilih otomatis terbaik     | —                   |
| `pdf2docx`  | **Gambar + layout terjaga**| Lebih lambat        |
| `pymupdf`   | Ekstrak gambar + teks      | Layout terbatas     |
| `text_only` | **Super cepat**            | Tanpa gambar        |

### Kompresi Ukuran File
- **Kompres PDF & DOCX** (kecilkan ukuran gambar)
- **Drag & drop di GUI**
- **CLI: `compress` & `compress-folder`**
- **Overwrite dengan `--force`**
- **Struktur folder terjaga**

### Dual Interface
- **GUI** – Tab: *Konversi* | *Kompresi*  
- **CLI** – Untuk scripting & otomasi

--- 

## Mulai Cepat

### 1. Instalasi
```bash
git clone <repo-url>
cd document_converter

# Lengkap (semua fitur + GUI)
pip install -r requirements.txt