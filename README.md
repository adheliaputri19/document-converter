# 📄 Document Converter

Alat konversi dokumen yang powerful dan serbaguna yang mendukung berbagai format termasuk DOC, DOCX, dan PDF dengan antarmuka GUI dan CLI.

## ✨ Fitur

- **Multiple Tipe Konversi**:
  - ✅ DOC/DOCX → PDF
  - ✅ PDF → DOCX (dengan preservasi gambar)
  - ✅ PDF → DOC (membutuhkan Microsoft Word)

- **Multiple Metode Konversi** untuk PDF ke DOCX:
  - 🔄 Auto (Rekomendasi) - Otomatis memilih metode terbaik
  - 🖼️ pdf2docx - Terbaik untuk mempertahankan gambar dan formatting
  - 📝 PyMuPDF - Ekstraksi Text + Gambar
  - ⚡ Text Only - Konversi cepat untuk PDF text-only

- **Dual Interface**:
  - 🖥️ Graphical User Interface (GUI) - Aplikasi desktop user-friendly
  - ⌨️ Command Line Interface (CLI) - Untuk automasi dan scripting

- **Smart Detection**:
  - Auto-detect instalasi Microsoft Word
  - UI adaptif berdasarkan library yang tersedia
  - Metode konversi fallback

## 🚀 Mulai Cepat

### Instalasi

1. **Clone atau download project**:
```bash
git clone <repository-url>
cd document_converter
```

2. **Install dependencies yang diperlukan**:
```bash
# Instalasi minimal (fungsi dasar)
pip install docx2pdf pymupdf python-docx

# Instalasi lengkap (semua fitur)
pip install docx2pdf pdf2docx pymupdf python-docx comtypes
```

### Penggunaan

#### Mode GUI (Direkomendasikan untuk kebanyakan user)
```bash
python main.py
```

#### Mode CLI (Untuk automasi)
```bash
# Convert DOCX ke PDF
python -m cli.cli_converter input.docx output.pdf doc_to_pdf

# Convert PDF ke DOCX dengan metode spesifik
python -m cli.cli_converter input.pdf output.docx pdf_to_docx --method pdf2docx

# Convert PDF ke DOC (butuh MS Word)
python -m cli.cli_converter input.pdf output.doc pdf_to_doc
```

## 📁 Struktur Project

```
document_converter/
│
├── conversion/           # Logic konversi
│   ├── __init__.py
│   ├── strategies.py     # Strategi konversi
│   ├── engine.py         # Mesin konversi
│
├── ui/                   # User interface
│   ├── __init__.py
│   ├── gui_manager.py    # Implementasi GUI
│
├── cli/                  # Command line interface
│   ├── __init__.py
│   ├── cli_converter.py  # Implementasi CLI
│
├── utils/                # Utility functions
│   ├── __init__.py
│   ├── file_handler.py   # Operasi file
│
├── factory.py            # Factory pattern
├── main.py               # Main entry point
└── requirements.txt      # Dependencies
```

## 🛠️ Detail Teknis

### Format yang Didukung

| Konversi | Format Input | Format Output | Requirements |
|----------|--------------|---------------|--------------|
| DOC/DOCX → PDF | .doc, .docx | .pdf | library docx2pdf |
| PDF → DOCX | .pdf | .docx | PyMuPDF atau pdf2docx |
| PDF → DOC | .pdf | .doc | Microsoft Word |

### Metode Konversi untuk PDF ke DOCX

1. **pdf2docx** (Rekomendasi)
   - ✅ Mempertahankan gambar
   - ✅ Mempertahankan formatting
   - ✅ Output kualitas terbaik
   - ❌ Proses lebih lambat

2. **PyMuPDF dengan Gambar**
   - ✅ Mengekstrak gambar
   - ✅ Preservasi text yang baik
   - ⚠️ Formatting terbatas

3. **Text Only**
   - ✅ Konversi tercepat
   - ✅ Ringan
   - ❌ Tanpa gambar
   - ❌ Formatting dasar

### Dependencies

**Dependencies Inti**:
- `docx2pdf` - Konversi DOCX ke PDF
- `pymupdf` - Processing PDF dan ekstraksi text
- `python-docx` - Pembuatan file DOCX

**Dependencies Opsional**:
- `pdf2docx` - Konversi PDF ke DOCX enhanced dengan gambar
- `comtypes` - Integrasi Microsoft Word untuk file .doc

**Built-in**:
- `tkinter` - Framework GUI
- `pathlib` - Penanganan path file

## 🎯 Contoh Penggunaan

### Penggunaan GUI
1. Jalankan aplikasi: `python main.py`
2. Pilih tipe konversi (DOC→PDF, PDF→DOCX, PDF→DOC)
3. Pilih metode konversi untuk file PDF
4. Browse dan pilih file input
5. Tentukan lokasi output
6. Klik "Konversi" untuk memulai konversi

### Contoh Penggunaan CLI

```bash
# Konversi dasar DOCX ke PDF
python -m cli.cli_converter document.docx document.pdf doc_to_pdf

# PDF ke DOCX dengan preservasi gambar
python -m cli.cli_converter report.pdf report.docx pdf_to_docx --method pdf2docx

# PDF ke DOC (butuh MS Word)
python -m cli.cli_converter manual.pdf manual.doc pdf_to_doc

# Konversi text-only cepat
python -m cli.cli_converter article.pdf article.docx pdf_to_docx --method text_only
```

## ⚠️ Batasan & Requirements

### Ketergantungan Microsoft Word
- Konversi **PDF → DOC** membutuhkan instalasi Microsoft Word
- Konversi **DOC → PDF** untuk file .doc membutuhkan Microsoft Word
- WPS Office dan alternatif lain tidak didukung untuk konversi ini

### Requirements Library
- Pastikan semua library yang diperlukan terinstall untuk fungsi yang diinginkan
- Beberapa metode konversi mungkin memiliki dependency tambahan
- Cek status library di GUI untuk informasi ketersediaan

### Batasan Ukuran File
- File PDF yang sangat besar mungkin membutuhkan waktu proses lebih lama
- Penggunaan memory meningkat dengan ukuran file dan konten gambar
- Pertimbangkan menggunakan metode "Text Only" untuk file besar

## 🔧 Troubleshooting

### Masalah Umum

1. **"Library tidak ditemukan"**
   ```bash
   pip install docx2pdf pymupdf python-docx pdf2docx comtypes
   ```

2. **Konversi PDF ke DOC gagal**
   - Install Microsoft Word
   - Pastikan comtypes terinstall: `pip install comtypes`

3. **Gambar hilang dalam konversi PDF ke DOCX**
   - Gunakan metode "pdf2docx" instead of "PyMuPDF"
   - Pastikan pdf2docx terinstall: `pip install pdf2docx`

4. **Konversi terlalu lambat untuk file besar**
   - Gunakan metode "Text Only" untuk konversi lebih cepat
   - Tutup aplikasi lain untuk membebaskan resources system

### Mode Debug
Untuk log konversi detail, cek output console dimana tersedia.

## 📄 Lisensi

Project ini untuk tujuan edukasi sebagai bagian dari latihan kolaborasi tim.

---

**Selamat Mengkonversi!** 🎉