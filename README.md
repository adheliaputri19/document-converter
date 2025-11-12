# Redme Aplikasi Document Converter & Compressor Pro

## 📋 Deskripsi Aplikasi
**Document Converter & Compressor Pro** adalah aplikasi desktop yang memungkinkan pengguna untuk mengkonversi dan mengkompres dokumen dengan antarmuka yang user-friendly. Aplikasi ini mendukung format dokumen populer seperti DOC, DOCX, dan PDF.

## ✨ Fitur Utama

### 🔄 Konversi Dokumen
- **DOC/DOCX ke PDF** - Konversi dokumen Word ke format PDF
- **PDF ke DOCX** - Konversi PDF kembali ke format Word
- **PDF ke DOC** - Konversi PDF ke format DOC lama
- **Multiple Methods** - Mendukung berbagai metode konversi:
  - Auto (otomatis pilih terbaik)
  - pdf2docx
  - pymupdf
  - text_only

### 📦 Kompres Dokumen
- **Kompres PDF & DOCX** - Reduksi ukuran file dokumen
- **Multiple Compression Levels**:
  - Low (kualitas tinggi)
  - Medium (seimbang)
  - High (ukuran kecil)
- **Batch Processing** - Kompres multiple file sekaligus
- **Drag & Drop** - Support drag and drop file

## 🛠️ Teknologi yang Digunakan

### Backend
- **Python 3.8+**
- **python-docx** - Manipulasi dokumen Word
- **PyMuPDF** - Manipulasi PDF
- **pdf2docx** - Konversi PDF ke Word
- **comtypes** - Integrasi dengan MS Word (Windows)

### Frontend
- **Tkinter** - GUI framework
- **TkinterDnD2** - Drag and drop functionality
- **ttk** - Themed widgets untuk tampilan modern

## 📥 Instalasi

### Prerequisites
- Python 3.8 atau lebih baru
- pip (Python package manager)

### Langkah Instalasi

1. **Clone atau download project**
```bash
git clone [repository-url]
cd document-converter-compressor
```

2. **Buat virtual environment (recommended)**
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Jalankan aplikasi**
```bash
python main.py
```

### Dependencies (requirements.txt)
```
tkinterdnd2>=0.3.0
python-docx>=1.1.0
PyMuPDF>=1.23.0
pdf2docx>=0.5.8
comtypes>=1.1.14
pathlib2>=2.3.7
```

## 🎯 Cara Penggunaan

### Konversi Dokumen
1. Buka tab "🔄 Konversi Dokumen"
2. Pilih file input dengan tombol "Browse"
3. Tentukan lokasi output
4. Pilih tipe konversi yang diinginkan
5. Pilih metode konversi (auto recommended)
6. Klik "🚀 MULAI KONVERSI"

### Kompres Dokumen
1. Buka tab "📦 Kompres Ukuran"
2. Tambah file dengan:
   - Tombol "➕ Tambah File"
   - Tombol "📁 Tambah Folder"
   - Drag & drop file ke area daftar
3. Pilih level kompresi
4. Tentukan folder output
5. Klik "📦 MULAI KOMPRES"

## 🏗️ Struktur Project

```
document-converter-compressor/
│
├── main.py                 # Entry point aplikasi
├── requirements.txt        # Dependencies
├── README.md              # Dokumentasi
├── factory.py             # Mengelola objek
│
├── ui/
│   └── gui_manager.py     # Manajemen GUI dan tampilan
│
├── conversion/
│   ├── engine.py          # Engine utama konversi
│   └── compressor.py      # Engine kompresi dokumen
|   └── strategies.py      # Sistem konversi
│
├── utils/
│   └── file_handler.py    # Utilities handling file
│
└── cli/                 
    ├── cli_converter.py
    └── cli_converter.py
```

## 🔧 Konfigurasi

### Supported Formats
- **Input**: .doc, .docx, .pdf
- **Output**: .pdf, .doc, .docx

### System Requirements
- **OS**: Windows 10/11, Linux, macOS
- **RAM**: Minimum 2GB
- **Storage**: 100MB free space
- **MS Office**: Optional (untuk konversi yang lebih baik)

## 🐛 Troubleshooting

### Common Issues

1. **MS Word tidak terdeteksi**
   - Pastikan MS Office terinstall
   - Aplikasi tetap bisa berjalan tanpa MS Word

2. **Konversi PDF ke DOCX gagal**
   - Coba ganti metode konversi
   - Pastikan PDF tidak terproteksi

3. **Drag & drop tidak bekerja**
   - Pastikan menggunakan TkinterDnD2
   - Restart aplikasi

4. **Error permission denied**
   - Run sebagai administrator (Windows)
   - Check folder permissions

### Performance Tips
- Gunakan metode "auto" untuk konversi optimal
- Untuk kompresi, level "medium" memberikan hasil terbaik
- Tutup aplikasi lain saat processing file besar

## 📝 Changelog

### v1.0.0
- ✅ Konversi DOC/DOCX to PDF
- ✅ Konversi PDF to DOCX/DOC
- ✅ Kompres PDF & DOCX
- ✅ Drag & drop support
- ✅ Modern GUI interface

## 🤝 Kontribusi

1. Fork project
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open Pull Request

## 📄 Lisensi
Distributed under MIT License. See `LICENSE` for more information.
