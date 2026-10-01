# RPS UNIROW — Generator Rencana Pembelajaran Semester OBE Kurikulum 2023 Pendidikan Matematika

Modul Artificial Intelligence (custom skill) untuk Google Antigravity, Cursor IDE, dan Claude Code untuk menyusun, mengisi, dan memformat dokumen Rencana Pembelajaran Semester (RPS) sesuai format resmi dan standar mutu Universitas PGRI Ronggolawe (UNIROW) Tuban berbasis Kurikulum 2023 MBKM Program Studi Pendidikan Matematika serta template resmi Wakil Rektor 1 (versi narasumber OBE Pak Samsul).

---

## 🌟 Fitur Utama & Standar Mutu

1. **Format Baku A4 Landscape (100% Sesuai Template Resmi WR 1)**:
   - Orientasi **A4 Landscape** penuh (`w="16838" h="11909"` dxa).
   - Margin halaman presisi: Atas/Bawah `39.6 pt`, Kiri/Kanan `43.2 pt`.
   - Tipografi elegan **Cambria** (Body 10.5 pt, Header 11 pt, Judul 12 pt bold).
   - Garis tabel single-line warna `#B0B5B3` (sz 4 / 0.5 pt) dengan padding sel proporsional.

2. **Katalog 68 Mata Kuliah Lengkap (Semester 1 s.d. 8)**:
   - Terintegrasi penuh dengan Kurikulum 2023 MBKM Prodi Pendidikan Matematika (pemutakhiran 15 Januari 2025).
   - Memetakan 7 butir CPL Prodi Pendidikan Matematika, CPMK 1–4, dan Sub-CPMK 1–14 berlabel taksonomi Bloom.
   - Matriks korelasi bersarang (Nested Table 18 baris) berbobot presisi 100%.
   - Sesi perkuliahan 16 minggu dengan beban belajar SN-Dikti Permendikbudristek No. 53/2023 (PB, PT, KM).
   - Evaluasi Tengah Semester (UTS) di Minggu ke-8 (25%) dan Evaluasi Akhir Semester (UAS) di Minggu ke-16 (25%).
   - Rubrik Penilaian Holistik Matematika (Pemahaman Konsep, Analisis Kritis, Pemodelan, Sistematika Notasi).
   - Konversi nilai akhir Skala 7 resmi UNIROW Tuban (A s.d. E) dan 3-5 paket butir soal UAS HOTS.

---

## 📁 Struktur Berkas Skill

```text
skills/rps/
├── SKILL.md                          # Definisi instruksi skill & panduan operasional AI
├── README.md                         # Dokumentasi modul skill
├── assets/
│   ├── template_rps_wr1_narsum_obe.docx # Master template Word A4 Landscape resmi WR 1
│   ├── Template_RPS_UNIROW.docx      # Master template Word referensi sistem
│   └── logo_unirow.png               # Logo resmi UNIROW Tuban
├── scripts/
│   ├── build_rps.py                  # Engine generator otomatis dokumen Word (.docx)
│   └── generate_all_pmat_rps.py      # Batch generator seluruh 68 mata kuliah PMAT 2023
└── index.html                        # Portal web simulator & live generator interaktif
```

---

## 💻 Cara Penggunaan

### 1. Melalui Chat AI Antigravity
Cukup ketik perintah di obrolan Antigravity:
> `"/rps-unirow buatkan RPS mata kuliah Struktur Aljabar 3 SKS untuk Semester 4 Kurikulum 2023 Pendidikan Matematika UNIROW"`

AI akan secara otomatis memetakan CPL-Prodi PMAT, merumuskan CPMK & Sub-CPMK berlabel Bloom, menyusun silabus 16 minggu, serta menghasilkan dokumen Word `.docx` siap cetak.

### 2. Melalui CLI (Terminal Python Langsung)
```powershell
# Menghasilkan seluruh 68 mata kuliah sekaligus:
python "skills\rps\scripts\generate_all_pmat_rps.py"

# Menghasilkan satu mata kuliah via JSON:
python "skills\rps\scripts\build_rps.py" --data "data_mk.json" --output "RPS_Kalkulus_I_OBE.docx"
```

### 3. Melalui Simulator & Portal Web Interaktif
Buka file `skills/rps/index.html` langsung di peramban (browser) untuk:
- Memilih cepat dari katalog **68 Mata Kuliah** Kurikulum 2023 PMAT.
- Pratinjau interaktif tata letak A4 Landscape.
- Mengunduh langsung dokumen OpenXML `.docx` via client-side JSZip engine.

---

## 📄 Kontributor & Otorisasi
- **Program Studi**: Pendidikan Matematika (PMAT), FKIP Universitas PGRI Ronggolawe Tuban
- **Kaprodi**: Puji Rahayu, M.Pd. (NIDN. 0718018801)
- **Koordinator Kurikulum / RMK**: Rachmalia Vinda Kusuma, M.Pd. (NIDN. 0713058804)
- **Pengembang Skill**: Mario Fahmi Syahrial, M.Pd. & Tim Dosen Pendidikan Matematika
