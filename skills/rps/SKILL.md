---
name: rps-unirow
description: >
  Skill untuk menyusun, mengisi, dan memformat Rencana Pembelajaran Semester
  (RPS) sesuai format resmi dan standar mutu Universitas PGRI Ronggolawe
  (UNIROW) Tuban berbasis Kurikulum 2023 MBKM Program Studi Pendidikan Matematika
  serta mengacu pada template resmi Wakil Rektor 1 (versi narasumber OBE Pak Samsul).
  WAJIB gunakan skill ini setiap kali pengguna menyebut "RPS", "Rencana Pembelajaran Semester",
  "susun RPS", "buatkan RPS", "isi RPS", "format RPS UNIROW/UNIROW Tuban/Ronggolawe", ingin
  membuat dokumen .docx RPS untuk mata kuliah Pendidikan Matematika, melengkapi CPL/CPMK/Sub-CPMK,
  tabel matriks korelasi CPL terhadap Sub-CPMK, rencana pembelajaran mingguan 16 minggu (termasuk
  UTS di Mg 8 dan UAS di Mg 16), rubrik penilaian holistik, konversi nilai akhir skala 7, soal UAS,
  lembar validasi, atau ingin merevisi RPS agar presisi sesuai standar resmi kampus.
---

# RPS UNIROW — Panduan Resmi Penyusunan RPS Kurikulum Pendidikan Matematika 2023 (Presisi 100%)

Dokumen acuan standar mutu: 
1. **Template RPS Resmi WR 1 UNIROW** (Versi Narasumber OBE Pak Samsul): `assets/template_rps_wr1_narsum_obe.docx` / `assets/Template_RPS_UNIROW.docx`.
2. **Kurikulum 2023 MBKM Program Studi Pendidikan Matematika** FKIP Universitas PGRI Ronggolawe Tuban (Pemutakhiran 15 Januari 2025).
Katalog Terintegrasi: **68 Mata Kuliah Kurikulum 2023 PMAT (Semester 1–8)**.
Script Otomasi Generator: `scripts/build_rps.py` & `scripts/generate_all_pmat_rps.py`.
Live Simulator & Generator Web: `skills/rps/index.html`.

Skill ini menghasilkan dokumen `.docx` RPS yang **identik secara format, tata letak, dan tipografi** dengan standar mutu resmi kurikulum OBE UNIROW Tuban.

---

## 1. Spesifikasi Format & Desain Dokumen Acuan (Template WR 1)

| Parameter | Ketentuan Baku |
|---|---|
| **Orientasi Halaman** | **100% A4 Landscape** (1 *Section* penuh, tanpa rotasi ke portrait) |
| **Ukuran Kertas** | A4 Landscape: `w="16838" h="11909"` DXA (`841.7 × 595.45` pt) |
| **Margin Halaman** | Atas: `39.6 pt`, Bawah: `39.6 pt`, Kiri: `43.2 pt`, Kanan: `43.2 pt` |
| **Tipografi Utama** | **Cambria** (Isi tabel/body: 10.5 pt, Sub-header: 11 pt, Judul: 12 pt bold) |
| **Garis Tabel (Borders)**| Single line, warna `#B0B5B3`, ketebalan `sz="4"` (0.5 pt) di seluruh tabel |
| **Bantalan Sel (Padding)**| Top/Bottom: `100 dxa`, Left/Right: `140 dxa` |
| **Palet Warna Shading** | *Slate Gray* elegan: Header `#EAECEE`, Sub-header `#F2F4F4` / `#F4F6F7`, Kartu Validasi `#FAFAFA` |

---

## 2. Struktur Anatomi Dokumen RPS (4 Tabel + Paragraf Evaluasi)

Setiap dokumen RPS UNIROW wajib memuat komponen berurutan berikut:

### I. TABEL 0: Identitas, Otorisasi, Capaian Pembelajaran, & Detail MK (20 Baris)
1. **Baris 0**: Logo resmi UNIROW (`logo_unirow.png`) | KOP Resmi Universitas, Fakultas Keguruan dan Ilmu Pendidikan, dan Program Studi Pendidikan Matematika | Kode Dokumen (`PMAT/[Semester]/PMA/[Rumpun]`).
2. **Baris 1**: Judul `RENCANA PEMBELAJARAN SEMESTER` (Merge 8 kolom, Shading `#EAECEE`).
3. **Baris 2–3**: Identitas MK (Nama MK, Kode MK, Rumpun MK, Bobot SKS, Semester, Tgl Penyusunan: 15 Januari 2025).
4. **Baris 4–5**: Otorisasi (Dosen Pengembang RPS, Koordinator RMK: Rachmalia Vinda Kusuma, M.Pd., Kaprodi: Puji Rahayu, M.Pd.).
5. **Baris 6–7**: CPL-PRODI Pendidikan Matematika yang dibebankan pada MK:
   - **CPL-1 (S1)**: Bertakwa kepada Tuhan Yang Maha Esa dan mampu menunjukkan sikap religius;
   - **CPL-2 (S2)**: Menjunjung tinggi nilai kemanusiaan dalam menjalankan tugas berdasarkan agama, moral, dan etika;
   - **CPL-3 (S3)**: Berkontribusi dalam peningkatan mutu kehidupan bermasyarakat, berbangsa, bernegara, dan kemajuan peradaban berdasarkan Pancasila;
   - **CPL-4 (S8)**: Menginternalisasi nilai, norma, dan etika akademik;
   - **CPL-5 (KU1)**: Mampu menerapkan pemikiran logis, kritis, sistematis, dan inovatif dalam konteks pengembangan atau implementasi IPTEKS;
   - **CPL-6 (P1)**: Menguasai konsep pedagogi-didaktik matematika serta keilmuan matematika untuk merencanakan pembelajaran inovatif berbasis IPTEKS;
   - **CPL-7 (KK1)**: Mampu mengaplikasikan konsep dan prinsip didaktik-pedagogis matematika serta keilmuan matematika untuk merencanakan pembelajaran inovatif dengan memanfaatkan IPTEKS.
6. **Baris 8–9**: Capaian Pembelajaran Mata Kuliah (CPMK 1 s.d. CPMK 4).
7. **Baris 10–11**: Kemampuan akhir tiap tahapan belajar (Sub-CPMK 1 s.d. 14) berlabel taksonomi Bloom, contoh: `[C2, A2]`, `[C3, A3]`, `[C4, A3]`, `[C5, A4]`.
8. **Baris 12–13**: **Matriks Korelasi CPL terhadap Sub-CPMK (Tabel Bersarang / Nested Table 18 Baris)**:
   - Header: `Sub-CPMK / Evaluasi`, kolom masing-masing CPL (`CPL 1 (%)` s.d. `CPL 7 (%)`), dan `Bobot Penilaian (%)`.
   - Baris 1 s.d. 7: Sub-CPMK 1 s.d. 7 (tanda centang `✓` dan bobot 3%–4%).
   - Baris 8: Evaluasi Tengah Semester (`UTS`, Bobot `25%`, Shading `#F4F6F7`).
   - Baris 9 s.d. 15: Sub-CPMK 8 s.d. 14 (tanda centang `✓` dan bobot 3%–4%).
   - Baris 16: Evaluasi Akhir Semester (`UAS`, Bobot `25%`, Shading `#F4F6F7`).
   - Baris 17: `Total` (Nilai `100%` pada tiap CPL dan total bobot `100%`, Shading `#EAECEE`).
9. **Baris 14**: Deskripsi Singkat MK (1 paragraf komprehensif profil kompetensi keilmuan matematika).
10. **Baris 15**: Bahan Kajian: Materi Pembelajaran (dengan kode Bahan Kajian, misal `BK01...` s.d. `BK14...`).
11. **Baris 16–17**: Pustaka Utama (buku teks ber-ISBN) dan Pustaka Pendukung (jurnal bereputasi/SINTA).
12. **Baris 18–19**: Dosen Pengampu aktual Prodi Pendidikan Matematika dan Matakuliah Syarat.

### II. TABEL 1: Rencana Pembelajaran 16 Minggu (19 Baris)
Tabel matriks perkuliahan 16 minggu dengan format 8 kolom:
- **Baris 0–2**: Header 3 baris standar (Nomor kolom `(1)` s.d. `(8)`).
- **Baris 3–9 (Minggu 1 s.d. 7)**: Sesi perkuliahan reguler (Sub-CPMK 1 s.d. 7).
- **Baris 10 (Minggu 8 — UTS)**: Baris evaluasi khusus (`UJIAN TENGAH SEMESTER (UTS)\nEvaluasi penguasaan materi perkuliahan minggu 1 s.d. 7`, Bobot `25%`, Shading `#F4F6F7`).
- **Baris 11–17 (Minggu 9 s.d. 15)**: Sesi perkuliahan reguler lanjutan (Sub-CPMK 8 s.d. 14).
- **Baris 18 (Minggu 16 — UAS)**: Baris evaluasi khusus (`UJIAN AKHIR SEMESTER (UAS)\nEvaluasi komprehensif penguasaan capaian pembelajaran mata kuliah`, Bobot `25%`, Shading `#F4F6F7`).

> **Aturan Waktu Belajar SN-Dikti**:
> - **PB** (Tatap Muka): $sks \times 50$ menit/minggu (3 SKS = 150 menit).
> - **PT** (Terstruktur): $sks \times 60$ menit/minggu (3 SKS = 180 menit).
> - **KM** (Mandiri): $sks \times 60$ menit/minggu (3 SKS = 180 menit).

### III. TABEL 2: Rubrik Penilaian Holistik (5 Baris × 5 Kolom)
Tabel rubrik penilaian hasil belajar mahasiswa matematika:
- Kolom: `Aspek Penilaian`, `Sangat Baik (85-100)`, `Baik (70-84)`, `Cukup (60-69)`, `Bobot`.
- 4 Aspek Penilaian Utama:
  1. **Pemahaman Konsep Matematis (40%)**: Penguasaan definisi, teorema, dan prinsip fundamental tanpa kerancuan konsep.
  2. **Kemampuan Analisis & Penalaran Kritis (30%)**: Ketajaman berpikir logis-deduktif, pembuktian, dan kedalaman penafsiran.
  3. **Pemodelan & Aplikasi Konteks Nyata (20%)**: Presisi formulasi masalah riil ke model matematika berbasis IPTEKS.
  4. **Sistematika & Komunikasi Matematis (10%)**: Kerapian penulisan simbol, notasi, grafik, dan argumen ilmiah.

### IV. BAGIAN EVALUASI, SOAL UAS, & VALIDASI (Paragraf Terstruktur)
1. **Kriteria Kelulusan**: Syarat minimal nilai 60 (C) dengan komposisi:
   - Keaktifan & Partisipasi: `10%`
   - Tugas Terstruktur & Kuis: `15%`
   - UTS: `25%`
   - Proyek / Praktikum: `15%`
   - UAS: `35%`
2. **Konversi Nilai Akhir (Standar Mutu UNIROW Tuban - Skala 7)**:
   - `A  : 85 - 100 (Sangat Baik)`
   - `AB : 80 - 84`
   - `B  : 75 - 79 (Baik)`
   - `BC : 70 - 74`
   - `C  : 65 - 69 (Cukup - Batas Minimal Kelulusan)`
   - `D  : 60 - 64`
   - `E  : < 60 (Tidak Lulus)`
3. **Contoh Paket Soal Ujian Akhir Semester (UAS)**:
   - 3-5 butir soal pemecahan masalah analitis tingkat tinggi (HOTS) sesuai CPMK dan Sub-CPMK.
4. **Tabel 3: Lembar Pengesahan / Validasi (2 Baris × 2 Kolom)**:
   - Menyetujui: Ketua Program Studi Pendidikan Matematika (**Puji Rahayu, M.Pd.** / NIDN. 0718018801).
   - Mengetahui: Unit Jaminan Mutu (UJM) Prodi Pendidikan Matematika (**Rachmalia Vinda Kusuma, M.Pd.** / NIDN. 0713058804).

---

## 3. Eksekusi Generator Otomatis (Metode Presisi)

Gunakan script `build_rps.py` dan `generate_all_pmat_rps.py` yang sudah terintegrasi:

```powershell
# 1. Menghasilkan seluruh 68 RPS Kurikulum PMAT 2023 secara batch:
python "skills\rps\scripts\generate_all_pmat_rps.py"

# 2. Menghasilkan RPS perorangan via data JSON:
python "skills\rps\scripts\build_rps.py" --data "data_mk.json" --output "RPS_Nama_MK_OBE.docx"
```
