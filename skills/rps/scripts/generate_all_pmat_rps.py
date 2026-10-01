#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_all_68_courses_rps.py
Generator Lengkap 68 Mata Kuliah Kurikulum Pendidikan Matematika FKIP UNIROW Tuban.
Format OBE 2026, Tanpa Label Taksonomi [C..., A...], A4 Landscape, Slate Gray & Accent Header.
"""

import os
import sys
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

from build_rps import generate_rps_docx

OUTPUT_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), 'output_rps_docx')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Complete 68 Courses List
COURSES_68 = [
    # SEMESTER 1
    {
        "kode_mk": "KIP2102", "nama_mk": "Perkembangan Peserta Didik", "sks": 2, "sem": 1, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang karakteristik individu, pertumbuhan fisik, perkembangan intelek, emosi, sosial, nilai moral, kebutuhan remaja, dan penyesuaian diri peserta didik dalam konteks pembelajaran matematika.",
        "pustaka": ["Sunarto & Agung. (2018). Perkembangan Peserta Didik. Jakarta: Rineka Cipta.", "Santrock, J. W. (2019). Educational Psychology. New York: McGraw-Hill."],
        "topics": [
            "Hakikat Perkembangan Peserta Didik & Hakikat Individu",
            "Prinsip-Prinsip Umum Perkembangan & Faktor Pembawaan vs Lingkungan",
            "Pertumbuhan Fisik & Perkembangan Motorik Peserta Didik",
            "Perkembangan Intelek, Kognitif (Teori Piaget), & Kemampuan Berpikir Matematika",
            "Perkembangan Emosi & Pengelolaan Kecemasan Matematika (Math Anxiety)",
            "Perkembangan Sosial & Karakteristik Hubungan Teman Sebaya",
            "Perkembangan Bahasa & Komunikasi Matematika Peserta Didik",
            "Perkembangan Bakat Khusus & Minat Belajar Siswa Sekolah",
            "Perkembangan Moral, Nilai, Kode Etik, & Sikap Akademik",
            "Jenis-Jenis Kebutuhan Remaja & Implikasinya dalam Pembelajaran",
            "Tugas-Tugas Perkembangan Remaja Usia SMP/SMA",
            "Penyesuaian Diri Peserta Didik & Masalah-Masalah Remaja",
            "Perbedaan Individual (Gaya Belajar, Gender, Kemampuan Awal)",
            "Implikasi Perkembangan Peserta Didik terhadap Perencanaan Pembelajaran Matematika"
        ]
    },
    {
        "kode_mk": "KIP2101", "nama_mk": "Pengantar Pendidikan", "sks": 2, "sem": 1, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang hakikat manusia dan pendidikan, landasan-landasan pendidikan, aliran-aliran pendidikan, sistem pendidikan nasional, serta peran pendidikan dalam pembangunan bangsa.",
        "pustaka": ["Tirtarahardja, U. & Sulo, S. L. (2015). Pengantar Pendidikan. Jakarta: Rineka Cipta.", "Mudyahardjo, R. (2014). Pengantar Pendidikan. Jakarta: Rajawali Pers."],
        "topics": [
            "Hakikat Manusia dan Pengertian Dasar Pendidikan",
            "Unsur-Unsur Pendidikan dan Sistem Komponen Pendidikan",
            "Landasan Filosofis dan Sosiologis Pendidikan",
            "Landasan Psikologis, Historis, dan Yuridis Pendidikan Indonesia",
            "Asas-Asas Pokok Pendidikan dan Penerapannya di Sekolah",
            "Masyarakat Masa Depan dan Perkembangan IPTEKS Pendidikan",
            "Lingkungan Pendidikan: Keluarga, Sekolah, dan Masyarakat",
            "Aliran Klasik Pendidikan: Empirisme, Nativisme, Naturalisme, Konvergensi",
            "Aliran Pokok Pendidikan Indonesia: Taman Siswa & INS Kayutanam",
            "Permasalahan Pokok Pendidikan Nasional (Pemerataan, Mutu, Efisiensi, Relevansi)",
            "Sistem Pendidikan Nasional (UU No. 20 Tahun 2003)",
            "Pendidikan dan Pembangunan Nasional",
            "Peran Pendidik dalam Transformasi Pendidikan Abad 21",
            "Inovasi Pendidikan dan Pendidikan Berkelanjutan (Sustainable Education)"
        ]
    },
    {
        "kode_mk": "PMA3107", "nama_mk": "Pengantar Dasar Matematika", "sks": 3, "sem": 1, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas logika matematika (pernyataan, kuantor, penarikan kesimpulan), teori himpunan, relasi, fungsi, kardinalitas, dan metode pembuktian matematika.",
        "pustaka": ["Soehakso, R. M. (2012). Pengantar Dasar Matematika. Yogyakarta: FMIPA UGM.", "Lipschutz, S. (2008). Set Theory and Related Topics. New York: McGraw-Hill."],
        "topics": [
            "Pernyataan, Operasi Logika (Negasi, Konjungsi, Disjungsi), dan Tabel Kebenaran",
            "Implikasi, Biimplikasi, Tautologi, Kontradiksi, dan Ekuivalensi Logis",
            "Kuantor Universal dan Kuantor Eksistensial serta Negasi Kuantor",
            "Aturan Penarikan Kesimpulan (Modus Ponens, Tollens, Silogisme)",
            "Metode Pembuktian Matematika: Bukti Langsung dan Tidak Langsung (Kontradiksi)",
            "Metode Pembuktian Kontraposisi dan Induksi Matematika",
            "Konsep Dasar Himpunan, Operasi Himpunan, dan Sifat-Sifatnya",
            "Diagram Venn, Aplikasi Himpunan, dan Hukum-Hukum Himpunan",
            "Himpunan Kuasa (Power Set) dan Produk Kartesius (Cartesian Product)",
            "Konsep Relasi, Sifat Relasi (Refleksif, Simetris, Transitif), dan Relasi Ekuivalensi",
            "Konsep Fungsi, Domain, Kodomain, Range, dan Jenis Fungsi (Injektif, Surjektif, Bijektif)",
            "Komposisi Fungsi dan Invers Fungsi",
            "Kardinalitas Himpunan: Himpunan Berhingga, Tak Berhingga, Denumerabel, dan Kontinuum",
            "Aplikasi Logika dan Teori Himpunan dalam Pembuktian Teorema Matematika"
        ]
    },
    {
        "kode_mk": "PMA4241", "nama_mk": "Kalkulus I", "sks": 3, "sem": 1, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas sistem bilangan real, pertidaksamaan, fungsi dan grafik, limit dan kekontinuan, turunan serta aplikasinya dalam pemecahan masalah.",
        "pustaka": ["Purcell, E. J., Varberg, D., & Rigdon, S. E. (2010). Calculus (9th ed.). Boston: Pearson.", "Stewart, J. (2015). Calculus: Early Transcendentals. Cengage Learning."],
        "topics": [
            "Sistem Bilangan Real, Garis Bilangan, dan Nilai Mutlak",
            "Pertidaksamaan Linier, Kuadrat, dan Pertidaksamaan Nilai Mutlak",
            "Fungsi, Operasi Fungsi, Grafik Fungsi, dan Simetri",
            "Fungsi Trigonometri, Fungsi Eksponensial, dan Fungsi Logaritma",
            "Konsep Limit Fungsi (Intuisi dan Definisi Presisi Epsilon-Delta)",
            "Teorema Limit, Limit Sepihak (Kiri dan Kanan), dan Limit Tak Hingga",
            "Kekontinuan Fungsi pada Titik dan Selang",
            "Definisi Turunan Fungsi dan Pengertian Geometris (Garis Singgung)",
            "Aturan-Aturan Dasar Turunan (Hasil Kali, Hasil Bagi, Aturan Rantai)",
            "Turunan Fungsi Trigonometri, Eksponensial, dan Logaritma",
            "Turunan Implisit dan Laju yang Berkaitan",
            "Maksimum, Minimum, dan Teorema Nilai Rataan (Mean Value Theorem)",
            "Kecembungan, Titik Belok, Grafik Fungsi Kompleks, dan Aturan L'Hopital",
            "Aplikasi Turunan dalam Pemodelan Optimalisasi dan Penaksiran"
        ]
    },
    {
        "kode_mk": "UNV1104", "nama_mk": "Bahasa Indonesia", "sks": 2, "sem": 1, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tata bahasa Indonesia yang baik dan benar, EYD/PUEBI, ragam bahasa ilmiah, penyusunan paragraf, serta penulisan karya ilmiah dan skripsi.",
        "pustaka": ["Keraf, G. (2010). Komposisi: Sebuah Pengantar Kemahiran Bahasa. Ende: Nusa Indah.", "Pedoman Umum Ejaan Bahasa Indonesia (PUEBI) Terbaru."],
        "topics": [
            "Kedudukan, Fungsi, dan Sejarah Bahasa Indonesia",
            "Ragam Bahasa Indonesia: Ragam Lisan, Tulis, Baku, dan Ilmiah",
            "Ejaan Bahasa Indonesia yang Disempurnakan (EYD/PUEBI) - Penggunaan Huruf & Tanda Baca",
            "Penulisan Kata, Kata Serapan, dan Istilah Ilmiah Matematika",
            "Pilihan Kata (Diksi) dan Kesepadanan Makna dalam Bahasa Akademik",
            "Kalimat Efektif: Ciri-Ciri, Struktur, dan Keparalelan",
            "Pengembangan Paragraf Ilmiah: Gagasan Utama dan Kalimat Penjelas",
            "Jenis-Jenis Paragraf: Deduktif, Induktif, Narasi, Eksposisi, Argumentasi",
            "Kutipan Langsung, Tak Langsung, dan Teknik Pengutipan (APA Style)",
            "Penyusunan Catatan Kaki, Catatan Perut, dan Daftar Pustaka",
            "Sistematika Penulisan Karya Tulis Ilmiah (Artikel, Makalah, Laporan)",
            "Penyusunan Ringkasan (Abstract), Ikhtisar, dan Sintesis Literatus",
            "Teknik Presentasi Ilmiah dan Bahasa Komunikasi Akademik",
            "Penyuntingan dan Revisi Naskah Karya Tulis Ilmiah Matematika"
        ]
    },
    {
        "kode_mk": "UNV1105", "nama_mk": "Bahasa Inggris", "sks": 2, "sem": 1, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas dasar-dasar tata bahasa Inggris, membaca dan memahami artikel ilmiah, struktur kalimat akademik, serta terminologi umum dalam pendidikan.",
        "pustaka": ["Murphy, R. (2019). English Grammar in Use. Cambridge University Press.", "Azar, B. S. (2016). Understanding and Using English Grammar. Pearson."],
        "topics": [
            "Introduction to Academic English & Classroom Language",
            "Basic Sentence Structure: Subject, Verb, Object, & Complements",
            "Tenses in Academic Writing: Present Simple, Past Simple, Present Perfect",
            "Parts of Speech in Academic Context: Nouns, Adjectives, Adverbs, Prepositions",
            "Reading Comprehension: Identifying Main Ideas & Supporting Details",
            "Scanning and Skimming Techniques for Academic Papers",
            "Vocabulary Building: Academic Word List (AWL) & Contextual Clues",
            "Understanding Passive Voice in Scientific & Mathematical Contexts",
            "Noun Clauses, Relative Clauses, and Reduced Clauses",
            "Cause and Effect Expressions & Connectors in English Texts",
            "Comparison and Contrast Structures in Scientific Discourse",
            "Basic Academic Writing: Writing Clear Paragraphs & Topic Sentences",
            "Listening Comprehension: Academic Lectures & Educational Podcasts",
            "Oral Presentation Skills: Introducing Topics, Visuals, and Conclusions"
        ]
    },
    {
        "kode_mk": "UNV1102", "nama_mk": "Pancasila", "sks": 2, "sem": 1, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas Pancasila dalam sejarah perjuangan bangsa, Pancasila sebagai dasar negara, etika politik, sistem filsafat, dan ideologi nasional Indonesia.",
        "pustaka": ["Kaelan. (2016). Pendidikan Pancasila. Yogyakarta: Paradigma.", "Notonagoro. (2010). Pancasila Secara Ilmiah Populer. Jakarta: Bumi Aksara."],
        "topics": [
            "Pengantar Pendidikan Pancasila di Perguruan Tinggi",
            "Pancasila dalam Arus Sejarah Perjuangan Bangsa Indonesia",
            "Perumusan dan Pengesahan Pancasila sebagai Dasar Negara (Sidang BPUPKI & PPKI)",
            "Pancasila sebagai Sistem Filsafat (Ontologi, Epistemologi, Aksiologi)",
            "Pancasila sebagai Ideologi Negara dan Perbandingannya dengan Ideologi Lain",
            "Hakikat Sila-Sila Pancasila dan Nilai-Nilai Kemanusiaan",
            "Pancasila sebagai Dasar Negara dan Sumber Hukum Nasional",
            "Pancasila sebagai Etika Politik dan Moral Kemasyarakatan",
            "Pancasila sebagai Paradigma Pembangunan IPTEK, Ekonomi, dan Politik",
            "Pancasila dalam Menghadapi Tantangan Radikalisme dan Intoleransi",
            "Pancasila dan Penegakan HAM di Indonesia",
            "Implementasi Nilai-Nilai Pancasila dalam Kehidupan Berbangsa dan Bernegara",
            "Aktualisasi Pancasila dalam Profesi Keguruan dan Pendidikan",
            "Refleksi Kritis Tantangan Idologi Pancasila di Era Globalisasi dan Digital"
        ]
    },
    {
        "kode_mk": "UNV1106", "nama_mk": "Pengenalan PGRI", "sks": 2, "sem": 1, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas jati diri PGRI, sejarah perjuangan guru Indonesia, Kode Etik Guru Indonesia, AD/ART PGRI, dan peran PGRI dalam meningkatkan mutu pendidikan.",
        "pustaka": ["PB PGRI. (2018). Jati Diri Persatuan Guru Republik Indonesia (PGRI). Jakarta.", "Soedijarto. (2010). Landasan dan Corak Pendidikan Berwawasan Ke-PGRI-an. Jakarta."],
        "topics": [
            "Hakikat dan Sejarah Keorganisasian PGRI",
            "Sejarah Perjuangan Guru Indonesia Sebelum dan Sesudah Kemerdekaan",
            "Pembentukan PGHB hingga Berdirinya PGRI 25 November 1945",
            "Jati Diri PGRI: Organisasi Profesi, Ketenagakerjaan, dan Perjuangan",
            "Visi, Misi, dan Tujuan Organisasi PGRI",
            "Anggaran Dasar dan Anggaran Rumah Tangga (AD/ART) PGRI",
            "Kode Etik Guru Indonesia (KEGI) dan Dewan Kehormatan Guru Indonesia (DKGI)",
            "Peran PGRI dalam Perjuangan Undang-Undang Guru dan Dosen (UU No. 14/2005)",
            "Lembaga Pendidikan PGRI (YPLP PGRI) dan Peranannya di Masyarakat",
            "Peran PGRI dalam Meningkatkan Kesejahteraan dan Perlindungan Hukum Guru",
            "Peran PGRI dalam Peningkatan Profesionalisme dan Mutu Pendidikan Nasional",
            "PGRI dan Tantangan Internasionalisasi Profesi Keguruan (EI - Education International)",
            "Jiwa Semangat Nilai 1945 (JSN 45) sebagai Nilai Dasar Ke-PGRI-an",
            "Peran Mahasiswa FKIP PGRI sebagai Generasi Penerus Perjuangan PGRI"
        ]
    },

    # SEMESTER 2
    {
        "kode_mk": "KIP2207", "nama_mk": "Belajar dan Pembelajaran", "sks": 2, "sem": 2, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang konsep dasar belajar, teori-teori belajar (Behavioristik, Kognitif, Konstruktivistik, Humanistik), prinsip pembelajaran, dan motivasi belajar.",
        "pustaka": ["Dimyati & Mudjiono. (2015). Belajar dan Pembelajaran. Jakarta: Rineka Cipta.", "Schunk, D. H. (2012). Learning Theories: An Educational Perspective. Pearson."],
        "topics": [
            "Hakikat Belajar dan Pengertian Pembelajaran",
            "Prinsip-Prinsip Belajar dan Faktor yang Mempengaruhi Proses Belajar",
            "Teori Belajar Behavioristik (Thorndike, Pavlov, Skinner) & Implikasinya",
            "Teori Belajar Kognitif (Piaget, Bruner, Ausubel) dalam Matematika",
            "Teori Belajar Konstruktivisme (Vygotsky, Piaget) & Pembelajaran Aktif",
            "Teori Belajar Humanistik (Rogers, Maslow) & Pembelajaran Berpusat pada Siswa",
            "Teori Belajar Sibernetik & Pemanfaatan Teknologi Informasi",
            "Motivasi Belajar: Konsep, Jenis, dan Strategi Peningkatan Motivasi Siswa",
            "Transfer Belajar dan Lupa dalam Pembelajaran Matematika",
            "Pendekatan, Strategi, Metode, dan Teknik Pembelajaran Inovatif",
            "Model-Model Pembelajaran Matematika Interaktif",
            "Kesulitan Belajar Matematika (Diagnostik & Remedial Teaching)",
            "Evaluasi dan Umpan Balik dalam Proses Pembelajaran",
            "Rancangan Kegiatan Pembelajaran Berbasis Teori Belajar Modern"
        ]
    },
    {
        "kode_mk": "KIP2208", "nama_mk": "Filsafat Pendidikan", "sks": 2, "sem": 2, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang hakikat filsafat, cabang-cabang filsafat (ontologi, epistemologi, aksiologi), aliran filsafat pendidikan, filsafat Pendidikan Pancasila, dan etika profesi pendidik.",
        "pustaka": ["Sadulloh, U. (2014). Pengantar Filsafat Pendidikan. Bandung: Alfabeta.", "Knight, G. R. (2006). Issues and Alternatives in Educational Philosophy. Andrews University Press."],
        "topics": [
            "Hakikat Filsafat, Berpikir Filosofis, dan Hubungan Filsafat dengan Pendidikan",
            "Cabang Ontologi Pendidikan: Hakikat Realitas dan Implikasi Kurikulum",
            "Cabang Epistemologi Pendidikan: Hakikat Pengetahuan dan Metodologi Pembelajaran",
            "Cabang Aksiologi Pendidikan: Hakikat Nilai, Etika Akademik, dan Estetika",
            "Sejarah Perkembangan Pemikiran Filsafat Pendidikan dari Yunani hingga Modern",
            "Aliran Filsafat Pendidikan Klasik: Esensialisme dan Perenialisme",
            "Aliran Filsafat Pendidikan Modern: Progresivisme dan Eksistensialisme",
            "Aliran Filsafat Rekonstruksionisme dan Konstruktivisme Pembelajaran Abad 21",
            "Landasan Filosofis Filsafat Pendidikan Pancasila dan Pemikiran Ki Hadjar Dewantara",
            "Hakikat Manusia, Pendidik, dan Peserta Didik dalam Perspektif Filsafat Indonesia",
            "Filsafat Pendidikan Matematika (Platonisme, Absolutisme, Fallibilisme, Sosio-Konstruktivisme)",
            "Problematika Pendidikan Nasional dan Isu-Isu Kritis Pembelajaran Era Digital",
            "Etika Moral Akademik dan Etika Profesi Keguruan dalam Kerangka Aksiologi",
            "Rumusan Makalah Reflektif Filsafat Pendidikan Matematika Inovatif"
        ]
    },
    {
        "kode_mk": "PMA4105", "nama_mk": "Kalkulus II", "sks": 3, "sem": 2, "rumpun": "Matematika Dasar", "prasyarat": "PMA4241",
        "deskripsi": "Mata kuliah ini membahas tentang integral tak tentu, integral tentu, teknik-teknik pengintegralan, aplikasi integral, integral tak wajar, serta barisan dan deret tak hingga.",
        "pustaka": ["Purcell, E. J., Varberg, D., & Rigdon, S. E. (2010). Calculus (9th ed.). Boston: Pearson.", "Stewart, J. (2015). Calculus: Early Transcendentals. Cengage Learning."],
        "topics": [
            "Anti-Turunan (Integral Tak Tentu) dan Sifat-Sifat Dasar Pengintegralan",
            "Pengintegralan dengan Subtitusi Sederhana",
            "Integral Tentu, Jumlah Riemann, dan Teorema Dasar Kalkulus I & II",
            "Teknik Pengintegralan: Integrasi Parsial",
            "Teknik Pengintegralan: Subtitusi Trigonometri",
            "Teknik Pengintegralan: Pecahan Parsial (Partial Fractions)",
            "Aplikasi Integral: Luas Daerah di Antara Dua Kurva",
            "Aplikasi Integral: Volume Benda Putar (Metode Cakram, Cincin, dan Kulit Tabung)",
            "Aplikasi Integral: Panjang Busur Kurva dan Luas Permukaan Benda Putar",
            "Integral Tak Wajar (Improper Integrals) dengan Batas Tak Hingga / Integran Tak Terbatas",
            "Konsep Barisan Tak Hingga dan Kekonvergenannya",
            "Deret Tak Hingga dan Uji Kekonvergenan Deret Positif (Uji Integral, Uji Banding, Uji Rasio)",
            "Deret Ganti Tanda (Alternating Series) dan Konvergen Mutlak/Bersyarat",
            "Deret Pangkat (Power Series), Deret Taylor, dan Deret Maclaurin"
        ]
    },
    {
        "kode_mk": "PMA4242", "nama_mk": "Statistika Dasar", "sks": 3, "sem": 2, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas pengumpulan data, penyajian data, ukuran pemusatan dan penyebaran, konsep peluang, distribusi probabilitas, serta statistik inferensial dasar (uji hipotesis).",
        "pustaka": ["Sudjana. (2016). Metoda Statistika. Bandung: Tarsito.", "Walpole, R. E., et al. (2012). Probability & Statistics for Engineers & Scientists. Pearson."],
        "topics": [
            "Pengertian Statistika, Statistik, Populasi, Sampel, dan Jenis-Jenis Data",
            "Teknik Pengumpulan Data, Sampling, dan Pengorganisasian Data",
            "Penyajian Data: Tabel Distribusi Frekuensi, Histogrm, Polygon, dan Ogive",
            "Ukuran Pemusatan Data: Rataan (Mean), Median, Modus Data Tunggal & Kelompok",
            "Ukuran Letak Data: Kuartil, Desil, dan Persentil",
            "Ukuran Penyebaran Data: Jangkauan, Simpangan Rata-Rata, Varians, dan Simpangan Baku",
            "Ukuran Kemiringan (Skewness) dan Keruncingan (Kurtosis) Kurva Distribusi",
            "Konsep Dasar Peluang, Ruang Sampel, Kejadian, dan Hukum Peluang",
            "Peluang Bersyarat, Kejadian Saling Bebas, dan Teorema Bayes",
            "Peubah Acak dan Distribusi Probabilitas Diskret (Binomial, Poisson)",
            "Distribusi Probabilitas Kontinu (Distribusi Normal dan Kurva Z)",
            "Konsep Dasar Uji Hipotesis, Galat Tipe I & II, dan Tingkat Signifikansi",
            "Uji Hipotesis Satu Sampel dan Dua Sampel (Uji-z dan Uji-t)",
            "Korelasi Linier Sederhana (Pearson) dan Regresi Linier Sederhana"
        ]
    },
    {
        "kode_mk": "PMA4112", "nama_mk": "Sistem Geometri", "sks": 3, "sem": 2, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas sejarah geometri, geometri Euclid, sistem aksioma, geometri insidensi, kekongruenan, kesejajaran, dan pengantar geometri non-Euclid (Hiperbolik & Elliptik).",
        "pustaka": ["Venema, G. A. (2012). Exploring Advanced Euclidean Geometry. MAA Press.", "Greenberg, M. J. (2008). Euclidean and Non-Euclidean Geometries. W. H. Freeman."],
        "topics": [
            "Sejarah Perkembangan Geometri dari Mesir Kuno, Yunani (Euclid) hingga Modern",
            "Sistem Aksiomatik: Pengertian Pangkal, Aksioma/Postulat, Definisi, dan Teorema",
            "Postulat Euclid dan Perkembangan Aksiomatisasi Geometri Hilbert",
            "Geometri Insidensi: Titik, Garis, Bidang, dan Sifat-Sifat Insidensi",
            "Konsep Kesejajaran (Parallelism) dan Postulat Kesejajaran Euclid",
            "Ukuran Sudut, Segitiga, dan Teorema Kekongruenan Segitiga (S-S-S, S-Sd-S, Sd-S-Sd)",
            "Teorema Kesejajaran, Sudut Dalam Segitiga, dan Poligon",
            "Kesebangunan Segitiga, Teorema Pythagoras, dan Aplikasinya",
            "Teori Lingkaran: Garis Singgung, Sudut Pusat, Sudut Keliling, dan Tali Busur",
            "Kelemahan Postulat Kelima Euclid dan Usaha Pembuktian Postulat Kesejajaran",
            "Penemuan Geometri Non-Euclid: Geometri Lobachevsky (Hiperbolik)",
            "Model Geometri Hiperbolik (Model Piringan Poincare & Setengah Bidang Poincare)",
            "Geometri Riemann (Elliptik/Sferis) dan Sifat-Sifat Segitiga Sferis",
            "Perbandingan Sifat Geometri Euclid, Hiperbolik, dan Elliptik"
        ]
    },
    {
        "kode_mk": "PMA4243", "nama_mk": "Etnomatematika", "sks": 2, "sem": 2, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang hubungan antara matematika dan budaya, eksplorasi konsep matematika dalam artefak, permainan tradisional, pola arsitektur, dan integrasinya dalam pembelajaran.",
        "pustaka": ["D'Ambrosio, U. (2006). Ethnomathematics: Link between Traditions and Modernity. Sense Publishers.", "Sirate, F. S. (2015). Menguak Nilai Etnomatematika. Makassar."],
        "topics": [
            "Hakikat Etnomatematika dan Pemikiran D'Ambrosio tentang Matematika & Budaya",
            "Dimensi Budaya dalam Matematika: Antropologi Matematika & Sosiologi Matematika",
            "Aktivitas Etnomatematika: Membilang, Mengukur, Merancang, Menemukan Lokasi, Memainkan",
            "Konsep Geometri Transformasi dan Simetri pada Motif Batik Nusantara",
            "Etnomatematika pada Arsitektur Candi, Rumah Adat, dan Bangunan Bersejarah",
            "Sistem Bilangan dan Kalender Tradisional (Jawa, Bali, Sasak, dll.)",
            "Etnomatematika dalam Permainan Tradisional Anak (Congklak, Engklek, Egrang)",
            "Konsep Aljabar dan Logika pada Kerajinan Anyaman dan Tenun Tradisional",
            "Etnomatematika Budaya Lokal Tuban (Batik Gedog, Ukiran, Arsitektur Lokal)",
            "Analisis Kritis Integrasi Etnomatematika dalam Kurikulum Sekolah (Sekolah Merdeka)",
            "Pengembangan Modul & Bahan Ajar Matematika Berbasis Etnomatematika",
            "Desain Pembelajaran Matematika Berbasis Budaya (Culture-Based Mathematics Learning)",
            "Metodologi Penelitian Etnomatematika (Etnografi & Studi Kasus Budaya)",
            "Penyusunan Proposal / Laporan Etnomatematika Budaya Nusantara"
        ]
    },
    {
        "kode_mk": "UNV1103", "nama_mk": "Kewarganegaraan", "sks": 2, "sem": 2, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang hakikat pendidikan kewarganegaraan, identitas nasional, hak dan kewajiban warga negara, demokrasi, Wawasan Nusantara, dan Ketahanan Nasional.",
        "pustaka": ["Winarno. (2014). Pembelajaran Pendidikan Kewarganegaraan. Jakarta: Bumi Aksara.", "Kaelan & Zubaidi. (2012). Pendidikan Kewarganegaraan. Yogyakarta: Paradigma."],
        "topics": [
            "Hakikat dan Urgensi Pendidikan Kewarganegaraan di Perguruan Tinggi",
            "Identitas Nasional sebagai Karakter Bangsa Indonesia (Bhinneka Tunggal Ika)",
            "Integrasi Nasional sebagai Parameter Persatuan dan Kesatuan Bangsa",
            "Konstitusionalisme dan UUD NRI 1945 sebagai Hukum Dasar Indonesia",
            "Hak dan Kewajiban Warga Negara dalam Konstitusi",
            "Demokrasi Pancasila: Konsep, Sejarah, dan Implikasi dalam Pemilu",
            "Penegakan Hukum yang Berkeadilan dan Lembaga Peradilan Indonesia",
            "Hak Asasi Manusia (HAM): Teori, Kebijakan, dan Pelanggaran HAM di Indonesia",
            "Geopolitik Indonesia: Wawasan Nusantara sebagai Konsepsi Wilayah",
            "Geostrategi Indonesia: Ketahanan Nasional dan Bela Negara",
            "Tantangan Ketahanan Nasional di Era Digital, Siber, dan Globalisasi",
            "Good Governance dan Pemberantasan Korupsi di Indonesia",
            "Peran Warga Negara Digital (Digital Citizenship) dan Etika Media Sosial",
            "Proyek Kewarganegaraan (Citizenship Project) Pemecahan Masalah Sosial"
        ]
    },
    {
        "kode_mk": "UNV1101", "nama_mk": "Agama", "sks": 2, "sem": 2, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang peran agama dalam kehidupan manusia, hubungan manusia dengan Tuhan, sesama manusia, dan alam, etika moral keagamaan, serta moderasi beragama.",
        "pustaka": ["Kementerian Agama RI. (2019). Moderasi Beragama. Jakarta: Kemenag RI.", "Syaltut, M. (2011). Islam: Aqidah dan Syariah. Jakarta: Pustaka Amani."],
        "topics": [
            "Hakikat Keberagamaan Manusia dan Peran Agama dalam Kehidupan",
            "Tuhan Yang Maha Esa dan Keimanan/Ketakwaan dalam Kehidupan Modern",
            "Manusia menurut Ajaran Agama: Hakikat, Martabat, dan Tanggung Jawab",
            "Sumber-Sumber Ajaran Agama dan Hukum Keagamaan",
            "Etika, Moral, dan Akhlak dalam Berbagai Aspek Kehidupan",
            "Agama dan IPTEKS: Integrasi Ilmu Pengetahuan dan Values/Nilai Keagamaan",
            "Agama dan Pembentukan Karakter Bangsa yang Berintegrasi dan Jujur",
            "Masyarakat Modern, Kerukunan Antarumat Beragama, dan Moderasi Beragama",
            "Toleransi Keagamaan dan Penolakan terhadap Ekstremisme/Radikalisme",
            "Agama dan Masalah Kemanusiaan Global (Kemiskinan, Keadilan, Perdamaian)",
            "Agama dan Pelestarian Lingkungan Hidup (Eko-Teologi)",
            "Agama, Kebudayaan, dan Tradisi Lokal Nusantara",
            "Kepemimpinan dan Tanggung Jawab Sosial Berperspektif Keagamaan",
            "Refleksi Spiritual dan Moralitas Profesional Pendidik Matematika"
        ]
    },

    # SEMESTER 3
    {
        "kode_mk": "KIP2301", "nama_mk": "Profesi dan Kebijakan Pendidikan", "sks": 2, "sem": 3, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang syarat profesi keguruan, standar kualifikasi & kompetensi guru, kode etik, organisasi profesi, serta kebijakan-kebijakan nasional pendidikan.",
        "pustaka": ["Soetjipto & Kosasi, R. (2011). Profesi Keguruan. Jakarta: Rineka Cipta.", "UU No. 14 Tahun 2005 tentang Guru dan Dosen."],
        "topics": [
            "Hakikat Profesi dan Karakteristik Pekerjaan Profesional",
            "Guru sebagai Profesi (UU No. 14/2005 tentang Guru dan Dosen)",
            "Empat Kompetensi Guru: Pedagogik, Kepribadian, Sosial, dan Profesional",
            "Kualifikasi Akademik dan Sertifikasi Pendidik (PPG)",
            "Kode Etik Guru Indonesia (KEGI) dan Penegakan Etika Keguruan",
            "Organisasi Profesi Keguruan (PGRI, IndoMS, dll.) dan Peranannya",
            "Pengembangan Keprofesian Berkelanjutan (PKB) bagi Guru",
            "Karir Guru, Jabatan Fungsional, dan Penilaian Kinerja Guru (PKG)",
            "Perlindungan Hukum, Keselamatan Kerja, dan Hak-Hak Pendidik",
            "Kebijakan Pendidikan Nasional: UU Sisdiknas & Standar Nasional Pendidikan (SNP)",
            "Kebijakan Kurikulum Merdeka dan Merdeka Belajar Kampus Merdeka (MBKM)",
            "Kebijakan Asesmen Nasional (AN), AKM, dan Tata Kelola Satuan Pendidikan",
            "Supervisi Pendidikan dan Pembinaan Profesionalitas Guru di Sekolah",
            "Tantangan Profesi Pendidik di Era Digital dan Industri 4.0/5.0"
        ]
    },
    {
        "kode_mk": "PMA4235", "nama_mk": "Bahasa Inggris Matematika", "sks": 2, "sem": 3, "rumpun": "Pembelajaran Matematika", "prasyarat": "UNV1105",
        "deskripsi": "Mata kuliah ini membahas istilah-istilah matematika dalam bahasa Inggris, pembacaan simbol/rumus matematika, penerjemahan artikel jurnal matematika, dan pengajaran matematika dalam bahasa Inggris (CLIL).",
        "pustaka": ["Marks, R. (2018). English for Mathematics. Oxford University Press.", "Glossary of Mathematical Terms (English-Indonesian)."],
        "topics": [
            "Mathematical Terms & Pronunciation of Numbers, Fractions, & Decimals",
            "Reading Algebraic Expressions, Equations, and Inequalities in English",
            "Reading Geometric Terms, Shapes, Angles, and Theorems in English",
            "Reading Calculus Symbols, Limits, Derivatives, and Integrals in English",
            "Mathematical Logic Terms: Propositions, Connectives, and Proofs in English",
            "Translating Indonesian Mathematical Texts to English",
            "Translating English Research Articles in Mathematics Education to Indonesian",
            "Analyzing Grammar and Structure of International Journal Abstracts",
            "Writing Mathematical Definitions and Explanations in English",
            "Formulating Word Problems in English for Junior/Senior High School",
            "Micro-Teaching: Explaining Math Concepts in English (CLIL Approach)",
            "Developing Classroom Discourse Commands in English (Classroom English)",
            "Designing English Mathematics Worksheets (LKPD Bahasa Inggris)",
            "Final Presentation of Math Lesson Delivered in Academic English"
        ]
    },
    {
        "kode_mk": "PMA4108", "nama_mk": "Teori Bilangan", "sks": 2, "sem": 3, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas sifat-sifat keterbagian, algoritma pembagian, FPB & KPK, persamaan Diophantine, teori kongruensi, teorema Euler/Fermat, dan aplikasi kriptografi.",
        "pustaka": ["Rosen, K. H. (2011). Elementary Number Theory and Its Applications. Pearson.", "Niven, I., Zuckerman, H. S., & Montgomery, H. L. (2008). An Introduction to the Theory of Numbers."],
        "topics": [
            "Sifat-Sifat Dasar Bilangan Bulat dan Induksi Matematika Terurut Rapi",
            "Keterbagian (Divisibility) dan Sifat-Sifat Keterbagian pada Bilangan Bulat",
            "Pembagi Persekutuan Terbesar (FPB) dan Algoritma Euclid",
            "Kelipatan Persekutuan Terkecil (KPK) dan Kombinasi Linier FPB",
            "Bilangan Prima, Bilangan Komposit, dan Teorema Utama Aritmatika",
            "Persamaan Diophantine Linier Dua Variabel dan Solusinya",
            "Konsep Dasar Kongruensi Modulo dan Sifat-Sifat Kongruensi",
            "Uji Keterbagian Bilangan Menggunakan Relasi Kongruensi",
            "Kongruensi Linier Satu Variabel dan Teorema Sisa Cina (Chinese Remainder Theorem)",
            "Teorema Kecil Fermat (Fermat's Little Theorem) dan Aplikasinya",
            "Fungsi Phi Euler (Euler's Totient Function) dan Teorema Euler",
            "Persamaan Kongruensi Kuadrat dan Simbol Legendre",
            "Aplikasi Teori Bilangan dalam Kriptografi RSA dan Keamanan Data",
            "Aplikasi Teori Bilangan dalam Pemecahan Masalah Olimpiade Sekolah"
        ]
    },
    {
        "kode_mk": "PMA4109", "nama_mk": "Aljabar Linier", "sks": 3, "sem": 3, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas sistem persamaan linier, eliminasi Gauss-Jordan, matriks & determinan, ruang vektor, kebebasan linier, basis & dimensi, nilai eigen, dan transformasi linier.",
        "pustaka": ["Anton, H. & Rorres, C. (2014). Elementary Linear Algebra. John Wiley & Sons.", "Lay, D. C. (2016). Linear Algebra and Its Applications. Pearson."],
        "topics": [
            "Sistem Persamaan Linier (SPL) dan Matriks Teraugmentasi",
            "Eliminasi Gauss dan Eliminasi Gauss-Jordan (Operasi Baris Elementer - OBE)",
            "Operasi Matriks, Sifat-Sifat Aritmatika Matriks, dan Matriks Invers",
            "Determinan Matriks, Ekspansi Kofaktor, dan Aturan Cramer",
            "Ruang Vektor Real R^n dan Sub-Ruang Vektor",
            "Kombinasi Linier, Merentang (Span), dan Kebebasan Linier Vektor",
            "Basis dan Dimensi Ruang Vektor",
            "Ruang Baris, Ruang Kolom, Ruang Nol (Nullspace), dan Rank Matriks",
            "Ruang Hasil Kali Dalam (Inner Product Space) dan Panjang/Sudut Vektor",
            "Ortonormalitas Vektor dan Proses Gram-Schmidt",
            "Nilai Eigen (Eigenvalues) dan Vektor Eigen (Eigenvectors)",
            "Diagonalisasi Matriks dan Matriks Simetris",
            "Transformasi Linier dari R^n ke R^m, Matriks Transformasi, Kernel, dan Image",
            "Aplikasi Aljabar Linier dalam Grafika Komputer dan Model Pengodean"
        ]
    },
    {
        "kode_mk": "PMA4106", "nama_mk": "Kalkulus Peubah Banyak", "sks": 3, "sem": 3, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4105",
        "deskripsi": "Mata kuliah ini membahas vektor di R^3, fungsi dua peubah atau lebih, turunan parsial, keterturunan, pengali Lagrange, integral lipat dua/tiga, serta kalkulus medan vektor (Teorema Green, Stokes, Divergensi).",
        "pustaka": ["Purcell, E. J., Varberg, D., & Rigdon, S. E. (2010). Calculus (9th ed.). Boston: Pearson.", "Stewart, J. (2015). Multivariable Calculus. Cengage Learning."],
        "topics": [
            "Vektor dan Geometri Ruang Dimensi Tiga (R^3), Hasil Kali Titik & Silang",
            "Fungsi Peubah Banyak (Domain, Range, Grafik Surface, dan Kurva Ketinggian)",
            "Limit dan Kekontinuan Fungsi Peubah Banyak",
            "Turunan Parsial Fungsi Dua/Lebih Peubah dan Turunan Tingkat Tinggi",
            "Keterturunan (Differentiability), Diferensial Total, dan Aturan Rantai Peubah Banyak",
            "Turunan Berarah (Directional Derivatives) dan Vektor Gradien",
            "Bidang Singgung dan Garis Normal pada Permukaan",
            "Maksimum dan Minimum Fungsi Dua Peubah serta Uji Turunan Parsial Kedua",
            "Metode Pengali Lagrange (Lagrange Multipliers) untuk Masalah Kendala",
            "Integral Lipat Dua pada Daerah Persegi Panjang dan Daerah Umum",
            "Integral Lipat Dua dalam Koordinat Kutub (Polar Coordinates)",
            "Integral Lipat Tiga dalam Koordinat Kartesius, Silinder, dan Bola",
            "Integral Garis (Line Integrals) dan Teorema Kebebasan Lintasan",
            "Teorema Green, Teorema Divergensi Gauss, dan Teorema Stokes"
        ]
    },
    {
        "kode_mk": "PMA4119", "nama_mk": "Statistika Matematika I", "sks": 3, "sem": 3, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4242",
        "deskripsi": "Mata kuliah ini membahas secara aksiomatis teori peluang, peubah acak diskret & Kontinu, fungsi kepadatan peluang, fungsi distribusi akumulatif, ekspektasi matematis, momen, dan fungsi pembangkit momen (MGF).",
        "pustaka": ["Hogg, R. V., McKean, J., & Craig, A. T. (2019). Introduction to Mathematical Statistics. Pearson.", "Bain, L. J. & Engelhardt, M. (2000). Introduction to Probability and Mathematical Statistics."],
        "topics": [
            "Pendekatan Aksiomatis Teori Peluang, Ruang Sampel, dan Aljabar Kejadian",
            "Teorema-Teorema Peluang, Peluang Bersyarat, Kemerdekaan Kejadian, Teorema Bayes",
            "Peubah Acak Diskret dan Fungsi Kepadatan Peluang Diskret (PMF)",
            "Peubah Acak Kontinu dan Fungsi Kepadatan Peluang Kontinu (PDF)",
            "Fungsi Distribusi Akumulatif (CDF) dan Sifat-Sifatnya",
            "Distribusi Gabungan Dua Peubah Acak (Joint Probability Distribution)",
            "Distribusi Marginal dan Distribusi Bersyarat Peubah Acak Gabungan",
            "Kemerdekaan Dua Peubah Acak dan Sifat-Sifat Ekspektasi Gabungan",
            "Ekspektasi Matematis Peubah Acak dan Sifat-Sifat Nilai Harapan",
            "Varians, Kovarians, dan Koefisien Korelasi Matematis",
            "Momen Peubah Acak (Momen Sekitar Asal & Momen Rataan)",
            "Fungsi Pembangkit Momen (Moment Generating Function - MGF) dan Sifat-Sifat MGF",
            "Distribusi Khusus Diskret: Uniform, Binomial, Poisson, Geometrik, Hipergeometrik (Kajian MGF)",
            "Distribusi Khusus Kontinu: Uniform, Normal, Eksponensial, Gamma, Beta (Kajian MGF)"
        ]
    },
    {
        "kode_mk": "PMA4113", "nama_mk": "Geometri Analitik", "sks": 3, "sem": 3, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4112",
        "deskripsi": "Mata kuliah ini membahas sistem koordinat (Kartesius, Kutub, Tabung, Bola), persamaan garis dan bidang di R^2 & R^3, Irisan Kerucut (Parabola, Elips, Hiperbola), serta permukaan kuadrik.",
        "pustaka": ["Thomas, G. B. & Finney, R. L. (2010). Analytic Geometry and Calculus.", "Purcell, E. J. (2008). Geometri Analitik Datar dan Ruang."],
        "topics": [
            "Sistem Koordinat Kartesius Datar (R^2), Jarak Dua Titik, dan Titik Pembagi",
            "Persamaan Garis Lurus di R^2: Bentuk Umum, Gradien, Berpotongan, dan Sejajar/Tegak Lurus",
            "Jarak Titik ke Garis dan Garis-Garis Sejajar di Dimensi Dua",
            "Lingkaran di R^2: Persamaan Standar, Umum, Garis Singgung, dan Kuasa Lingkaran",
            "Irisan Kerucut: Parabola (Persamaan Standar, Fokus, Direktriks, Garis Singgung)",
            "Irisan Kerucut: Elips (Persamaan Standar, Fokus, Eksentrisitas, Garis Singgung)",
            "Irisan Kerucut: Hiperbola (Persamaan Standar, Asimtot, Fokus, Garis Singgung)",
            "Rotasi dan Translasi Sumbu Koordinat untuk Mengeliminasi Suku xy",
            "Sistem Koordinat Kutub (Polar) dan Grafik Persamaan Kutub",
            "Sistem Koordinat Ruang (R^3), Vektor Posisi, dan Jarak Dua Titik di R^3",
            "Persamaan Bidang Datar di R^3 (Vektor Normal, Jarak Titik ke Bidang)",
            "Persamaan Garis Lurus di R^3 (Bentuk Parametrik dan Simetris)",
            "Permukaan Putaran dan Permukaan Kuadrik (Ellipsoid, Paraboloid, Hiperboloid)",
            "Sistem Koordinat Tabung (Cylindrical) dan Koordinat Bola (Spherical)"
        ]
    },
    {
        "kode_mk": "PMA4237", "nama_mk": "Basis Data", "sks": 2, "sem": 3, "rumpun": "Teknologi Pendidikan Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang konsep dasar basis data, Entity-Relationship Diagram (ERD), model data relasional, Structured Query Language (SQL: DDL, DML), normalisasi, dan implementasi DBMS.",
        "pustaka": ["Silberschatz, A., Korth, H. F., & Sudarshan, S. (2020). Database System Concepts. McGraw-Hill.", "Fathansyah. (2015). Basis Data. Bandung: Informatika."],
        "topics": [
            "Pengantar Sistem Basis Data vs Sistem Berkas Tradisional",
            "Arsitektur Sistem Basis Data (ANSI-SPARC 3-Level Architecture)",
            "Konsep Pemodelan Data dan Entity-Relationship Model (ERD)",
            "Entitas, Atribut, Kardinalitas Relasi, dan Notasi ERD (Chen & Crow's Foot)",
            "Transformasi ERD ke Diagram Relasional Matriks Tabel",
            "Konsep Model Data Relasional: Primary Key, Foreign Key, dan Integritas Data",
            "Normalisasi Basis Data: 1NF, 2NF, 3NF, dan BCNF untuk Mencegah Anomali",
            "Pengenalan Structured Query Language (SQL) dan Penggolongan Komando SQL",
            "Data Definition Language (DDL): CREATE, ALTER, DROP Database/Table",
            "Data Manipulation Language (DML): INSERT, UPDATE, DELETE Data",
            "SQL Querying Advanced: SELECT, WHERE, ORDER BY, GROUP BY, HAVING",
            "Penggabungan Tabel (SQL Joins: INNER, LEFT, RIGHT, FULL OUTER JOIN)",
            "Subquery, Fungsi Agregasi (COUNT, SUM, AVG, MAX, MIN), dan View",
            "Implementasi Sistem Manajemen Basis Data (MySQL/MariaDB) untuk Aplikasi Sekolah"
        ]
    },

    # SEMESTER 4
    {
        "kode_mk": "KIP2401", "nama_mk": "Metodologi Penelitian", "sks": 2, "sem": 4, "rumpun": "Penelitian Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas konsep dasar penelitian ilmiah, jenis-jenis penelitian (Kuantitatif, Kualitatif, R&D, PTK), penyusunan rumusan masalah, hipotesis, kajian pustaka, dan teknik pengumpulan data.",
        "pustaka": ["Sugiyono. (2018). Metode Penelitian Pendidikan (Kuantitatif, Kualitatif, R&D). Bandung: Alfabeta.", "Creswell, J. W. (2014). Research Design. SAGE Publications."],
        "topics": [
            "Hakikat Penelitian Ilmiah dan Etika Penelitian Pendidikan",
            "Ragam/Jenis Penelitian Pendidikan: Kuantitatif, Kualitatif, PTK, R&D, Mix-Methods",
            "Identifikasi Masalah, Latar Belakang Masalah, dan Pembatasan Masalah",
            "Perumusan Masalah Penelitian dan Tujuan Penelitian Pendidikan Matematika",
            "Kajian Pustaka, Kerangka Berpikir, dan Perumusan Hipotesis Penelitian",
            "Variabel Penelitian (Bebas, Terikat, Moderator, Intervening) dan Definisi Operasional",
            "Populasi, Sampel, dan Teknik Sampling (Probability & Non-Probability Sampling)",
            "Metode Penelitian Kuantitatif: Eksperimen (Pre, True, Quasi) dan Korelasional",
            "Metode Penelitian Kualitatif: Etnografi, Studi Kasus, Fenomenologi",
            "Penelitian Tindakan Kelas (PTK): Siklus Planning, Acting, Observing, Reflecting",
            "Penelitian dan Pengembangan (R&D): Model ADDIE, 4D, Borg & Gall",
            "Teknik Pengumpulan Data: Tes, Angket, Wawancara, Observasi, Dokumentasi",
            "Validitas dan Reliabilitas Instrumen Penelitian Pendidikan",
            "Sistematika Penyusunan Draft Proposal Penelitian Ilmiah"
        ]
    },
    {
        "kode_mk": "PMA3105", "nama_mk": "Kapita Selekta Matematika I", "sks": 3, "sem": 4, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas bedah materi matematika sekolah menengah pertama (SMP/MTs) secara mendalam, pemecahan masalah (problem solving), Miskonsepsi siswa, dan variasi soal HOTS.",
        "pustaka": ["Buku Teks Matematika SMP Kurikulum Merdeka Kemendikbudristek.", "Posamentier, A. S. (2013). Problem Solving in Mathematics."],
        "topics": [
            "Bedah Kurikulum Matematika SMP/MTs (Capaian Pembelajaran & Elemen Matematika)",
            "Pendalaman Materi Bilangan: Operasi Bilangan Bulat, Pecahan, Rasional, dan Bentuk Akar",
            "Miskonsepsi Siswa SMP pada Topik Bilangan & Strategi Remediasinya",
            "Pendalaman Materi Aljabar SMP: Bentuk Aljabar, Persamaan & Pertidaksamaan Linier Satu Variabel",
            "Pendalaman Materi Sistem Persamaan Linier Dua Variabel (SPLDV) & Pemodelannya",
            "Pendalaman Materi Relasi, Fungsi, dan Persamaan Garis Lurus",
            "Pengembangan Soal HOTS (Higher Order Thinking Skills) Topik Aljabar SMP",
            "Pendalaman Materi Geometri Datar SMP: Segitiga, Segiempat, Teorema Pythagoras",
            "Pendalaman Materi Kesebangunan dan Kekongruenan Bangun Datar SMP",
            "Pendalaman Materi Bangun Ruang Sisi Datar (Kubus, Balok, Prisma, Limas)",
            "Pendalaman Materi Bangun Ruang Sisi Lengkung (Tabung, Kerucut, Bola)",
            "Pendalaman Materi Statistika dan Peluang SMP (Penyajian Data & Peluang Empiris/Teoretis)",
            "Analisis Miskonsepsi Siswa pada Geometri & Statistika SMP",
            "Penyusunan Paket Soal Simulasi Ujian & Pembahasannya Berorientasi HOTS"
        ]
    },
    {
        "kode_mk": "PMA4116", "nama_mk": "Matematika Diskrit", "sks": 2, "sem": 4, "rumpun": "Matematika Dasar", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang kombinatorika (prinsip inklusi-eksklusi, pigeonhole principle), relasi rekursif, fungsi pembangkit, dan algoritma dasar matematika diskrit.",
        "pustaka": ["Munir, R. (2012). Matematika Diskrit. Bandung: Informatika.", "Rosen, K. H. (2012). Discrete Mathematics and Its Applications. McGraw-Hill."],
        "topics": [
            "Hakikat Matematika Diskrit dan Objek-Objek Diskrit",
            "Prinsip-Prinsip Dasar Membilang: Aturan Penjumlahan dan Aturan Perkalian",
            "Permutasi dan Kombinasi (Pengulangan, Melingkar, dan Pembagian Kelompok)",
            "Prinsip Sarang Burung Merpati (Pigeonhole Principle) dan Aplikasinya",
            "Prinsip Inklusi-Eksklusi dan Aplikasi Pembatalan Pilihan",
            "Derangement (Permutasi Kekacauan) dan Aplikasi Inklusi-Eksklusi",
            "Relasi Rekursif (Recurrence Relations): Formulasi dan Pemodelan Masalah",
            "Penyelesaian Relasi Rekursif Linier Homogen Berkoefisien Konstan",
            "Penyelesaian Relasi Rekursif Linier Non-Homogen",
            "Konsep Fungsi Pembangkit (Generating Functions) Biasa dan Eksponensial",
            "Penyelesaian Masalah Kombinatorik dan Relasi Rekursif dengan Fungsi Pembangkit",
            "Kombinatorik pada Kata dan Kode String Diskrit",
            "Algoritma Diskrit: Efisiensi, Kompleksitas Waktu (Notasi Big-O)",
            "Aplikasi Matematika Diskrit dalam Analisis Struktur Algoritma"
        ]
    },
    {
        "kode_mk": "PMA4107", "nama_mk": "Persamaan Diferensial", "sks": 3, "sem": 4, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4105",
        "deskripsi": "Mata kuliah ini membahas tentang Persamaan Diferensial Biasa (PDB) orde satu, PDB linier orde tinggi, faktor integrasi, metode koefisien tak tentu, variasi parameter, dan Transformasi Laplace.",
        "pustaka": ["Boyce, W. E. & DiPrima, R. C. (2012). Elementary Differential Equations. Wiley.", "Nagle, R. K., Saff, E. B., & Snider, A. D. (2017). Fundamentals of Differential Equations. Pearson."],
        "topics": [
            "Klasifikasi Persamaan Diferensial (PDB/PDP, Orde, Derajat, Linieritas)",
            "PDB Orde Satu: Pembentukan PDB dan Metode Pemisahan Variabel",
            "PDB Orde Satu: Persamaan Diferensial Homogen",
            "PDB Orde Satu: Persamaan Diferensial Eksak dan Faktor Integrasi",
            "PDB Linier Orde Satu dan Persamaan Bernoulli",
            "Aplikasi PDB Orde Satu: Pertumbuhan Populasi, Peluruhan Radioaktif, Hukum Pendinginan Newton",
            "PDB Linier Orde Dua Homogen Berkoefisien Konstan (Persamaan Karakteristik)",
            "Wronskian dan Kebebasan Linier Solusi PDB Orde Dua",
            "PDB Linier Orde Dua Non-Homogen: Metode Koefisien Tak Tentu",
            "PDB Linier Orde Dua Non-Homogen: Metode Variasi Parameter",
            "Persamaan Euler-Cauchy Orde Tinggi",
            "Aplikasi PDB Orde Dua: Getaran Mekanik (Sistem Pegas-Massa) dan Rangkaian Listrik RLC",
            "Transformasi Laplace: Definisi, Sifat-Sifat, dan Transformasi Invers",
            "Penyelesaian Masalah Nilai Awal PDB Menggunakan Transformasi Laplace"
        ]
    },
    {
        "kode_mk": "PMA4114", "nama_mk": "Geometri Transformasi", "sks": 2, "sem": 4, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4112",
        "deskripsi": "Mata kuliah ini membahas tentang fungsi transformasi pada bidang, Isometri (Translasi, Refleksi, Rotasi, Geseran Gali), Komposisi Isometri, Dilatasi (Similaritas), dan pendekatan matriks.",
        "pustaka": ["Rawuh. (2007). Geometri Transformasi. Jakarta: Universitas Terbuka.", "Eccles, F. M. (2010). An Introduction to Transformational Geometry."],
        "topics": [
            "Konsep Transformasi pada Bidang (Fungsi Bijektif dari V ke V)",
            "Konsep Isometri: Definisi, Sifat Kolineasi, Mengawetkan Jarak, Sudut, dan Luas",
            "Translasi (Geseran): Definisi, Sifat-Sifat, dan Representasi Vektor/Matriks",
            "Refleksi (Pencerminan): Definisi, Sifat-Sifat, Titik Tetap, dan Garis Tetap",
            "Rotasi (Putaran): Definisi, Pusat Rotasi, Sudut Putar, dan Matriks Rotasi",
            "Garis Cermin Sejajar dan Garis Cermin Berpotongan (Komposisi Refleksi)",
            "Refleksi Geser (Glide Reflection): Definisi dan Sifat-Sifatnya",
            "Komposisi Isometri (Teorema Struktur Isometri Bidang)",
            "Grup Isometri dan Sifat-Sifat Struktur Aljabar Isometri",
            "Dilatasi (Perkalian Skala): Pusat Dilatasi, Faktor Skala, dan Matriks Dilatasi",
            "Similarity (Transformasi Kesebangunan): Definisi, Komposisi Dilatasi & Isometri",
            "Transformasi Afin (Affine Transformation) dan Transformasi Proyektif",
            "Aplikasi Geometri Transformasi pada Seni Batik, Tessellation, dan Grafika Komputer",
            "Penyelesaian Masalah Geometri Sekolah Menggunakan Pendekatan Transformasi"
        ]
    },
    {
        "kode_mk": "PMA4110", "nama_mk": "Struktur Aljabar I", "sks": 3, "sem": 4, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4109",
        "deskripsi": "Mata kuliah ini membahas operasi biner, grup, sifat-sifat dasar grup, subgrup, grup siklik, grup permutasi, koset, Teorema Lagrange, subgrup normal, grup faktor, dan homomorfisme grup.",
        "pustaka": ["Gallian, J. A. (2017). Contemporary Abstract Algebra. Cengage Learning.", "Durbin, J. R. (2008). Modern Algebra: An Introduction. Wiley."],
        "topics": [
            "Konsep Operasi Biner, Sifat Asosiatif, Komutatif, Elemen Identitas, dan Invers",
            "Definisi Grup, Sifat-Sifat Dasar Grup, dan Hukum Pembatalan (Cancellation Law)",
            "Contoh-Contoh Grup Khusus: Grup Abel, Z_n, S_n, D_n, dan GL(n,R)",
            "Definisi Subgrup dan Kriteria Pengujian Subgrup (Subgroup Tests)",
            "Grup Siklik, Generator Grup, dan Sifat Subgrup dari Grup Siklik",
            "Grup Permutasi, Notasi Siklus, Transposisi, dan Permutasi Genap-Ganjil",
            "Koset Kiri dan Koset Kanan dari Suatu Subgrup serta Sifat Partisi",
            "Teorema Lagrange tentang Orde Subgrup/Elemen dan Implikasinya (Teorema Fermat/Euler)",
            "Subgrup Normal dan Sifat Kesetaraan Koset Kiri/Kanan",
            "Grup Faktor / Grup Kuosien (G/N) dan Operasi Perkalian Koset",
            "Homomorfisme Grup, Peta (Image), dan Inti Homomorfisme (Kernel)",
            "Isomorfisme Grup, Automorfisme, dan Teorema Cayley",
            "Teorema Isomorfisme Pertama (First Isomorphism Theorem) bagi Grup",
            "Aplikasi Teori Grup dalam Simetri Molekul dan Kriptografi"
        ]
    },
    {
        "kode_mk": "PMA4115", "nama_mk": "Analisis Real", "sks": 3, "sem": 4, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA3107",
        "deskripsi": "Mata kuliah ini membahas sifat aljabar, urutan, & kelengkapan bilangan real, barisan bilangan real & kekonvergenannya, limit fungsi, kekontinuan fungsi, dan turunan secara rigourous.",
        "pustaka": ["Bartle, R. G. & Sherbert, D. R. (2011). Introduction to Real Analysis. John Wiley & Sons.", "Abbott, S. (2015). Understanding Analysis. Springer."],
        "topics": [
            "Sifat Aljabar dan Sifat Urutan Himpunan Bilangan Real R",
            "Pertidaksamaan, Nilai Mutlak, dan Lingkungan (Neighborhood)",
            "Sifat Kelengkapan R: Supremum, Infimum, dan Aksioma Kelengkapan",
            "Sifat Archimedes dan Kerapatan Bilangan Rasional di R",
            "Barisan Bilangan Real: Definisi Limit Barisan (Epsilon-N) dan Uji Kekonvergenan",
            "Teorema Limit Barisan, Barisan Monoton, dan Teorema Konvergensi Monoton",
            "Sub-Barisan (Subsequences) dan Teorema Bolzano-Weierstrass",
            "Barisan Cauchy dan Kriteria Konvergensi Cauchy",
            "Limit Fungsi: Definisi Epsilon-Delta dan Kriteria Barisan untuk Limit (Sequential Criterion)",
            "Teorema-Teorema Limit Fungsi dan Limit Tak Hingga",
            "Kekontinuan Fungsi pada Titik dan Selang: Kriteria Barisan & Diskontinuitas",
            "Fungsi Kontinu pada Selang Tertutup Terbatas: Teorema Nilai Ekstrem & Teorema Nilai Antara",
            "Kekontinuan Seragam (Uniform Continuity) dan Teorema Kekontinuan Seragam",
            "Turunan Fungsi: Definisi Presisi, Aturan Rantai, Teorema Rolle, dan Teorema Nilai Rataan"
        ]
    },
    {
        "kode_mk": "PMA4230", "nama_mk": "Teori Graf", "sks": 2, "sem": 4, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4116",
        "deskripsi": "Mata kuliah ini membahas tentang graf dan subgraf, derajat simpul, graf terhubung, lintasan Euler & Hamilton, pohon (tree), graf planar, pewarnaan graf, dan algoritma graf.",
        "pustaka": ["Chartrand, G. & Lesniak, L. (2010). Graphs & Digraphs. CRC Press.", "Bondy, J. A. & Murty, U. S. R. (2008). Graph Theory. Springer."],
        "topics": [
            "Definisi Graf, Simpul (Vertex), Sisi (Edge), Graf Sederhana, dan Multigraf",
            "Derajat Simpul dan Handshaking Lemma (Teorema Keterbagian Derajat)",
            "Jenis-Jenis Graf Khusus: Graf Lengkap (K_n), Lingkaran (C_n), Roda (W_n), Bipartit (K_m,n)",
            "Subgraf, Graf Bagian Terentang, dan Isomorfisme Graf",
            "Representasi Graf: Matriks Ketenanggan (Adjacency) dan Matriks Insidensi",
            "Keterhubungan (Connectedness), Lintasan (Path), Sirkuit (Cycle), dan Komponen",
            "Graf Euler: Sirkuit Euler, Teorema Euler, dan Masalah Jembatan Konigsberg",
            "Graf Hamilton: Sirkuit Hamilton, Teorema Dirac, dan Traveling Salesperson Problem (TSP)",
            "Pohon (Tree): Sifat-Sifat Pohon, Spanning Tree, dan Algoritma Kruskal/Prim",
            "Pohon Berakar (Rooted Tree) dan Penelusuran Pohon (Traversal: Preorder, Inorder, Postorder)",
            "Graf Planar: Teorema Euler Bidang (V - E + F = 2) dan Teorema Kuratowski",
            "Pewarnaan Graf (Graph Coloring): Pewarnaan Simpul, Bilangan Kromatik, Teorema Empat Warna",
            "Graf Berarah (Digraph) dan Algoritma Jalur Terpendek (Dijkstra)",
            "Aplikasi Teori Graf dalam Jaringan Komputer, Transportasi, dan Media Sosial"
        ]
    },

    # SEMESTER 5
    {
        "kode_mk": "PMA5101", "nama_mk": "Kewirausahaan", "sks": 2, "sem": 5, "rumpun": "Kewirausahaan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang mindset kewirausahaan, analisis peluang usaha, Business Model Canvas (BMC), studi kelayakan bisnis, pemasaran digital, dan edupreneurship matematika.",
        "pustaka": ["Osterwalder, A. & Pigneur, Y. (2010). Business Model Generation. Wiley.", "Kasmir. (2017). Kewirausahaan. Jakarta: Rajawali Pers."],
        "topics": [
            "Hakikat Kewirausahaan dan Character Building Entrepreneur Success",
            "Edupreneurship: Peluang Bisnis Berbasis Bidang Pendidikan dan Matematika",
            "Identifikasi Peluang Usaha, Inovasi Produk, dan Design Thinking",
            "Analisis Pasar, Segmentasi, Targeting, dan Positioning (STP)",
            "Business Model Canvas (BMC): 9 Elemen Blok Pembangunan Bisnis",
            "Perencanaan Keuangan Usaha: HPP, Pricing Strategy, dan Break-Even Point (BEP)",
            "Studi Kelayakan Bisnis: Aspek Legal, Operasional, dan Finansial",
            "Strategi Pemasaran Digital (Digital Marketing & Social Media Branding)",
            "Manajemen Tim, Kepemimpinan Bisnis, dan Budaya Kerja Profesional",
            "Pemanfaatan Teknologi Informasi dan Platform E-Commerce",
            "Pengelolaan Akses Modal, Pitch Deck, dan Presentasi Bisnis ke Investor",
            "Penyusunan Rencana Bisnis (Business Plan) Produk Edukasi/Jasa Matematika",
            "Gelar Karya / Bazar Kewirausahaan Mahasiswa (Business Expo)",
            "Evaluasi Kinerja Bisnis dan Strategi Keberlanjutan Usaha (Sustainability)"
        ]
    },
    {
        "kode_mk": "PMA4117", "nama_mk": "Program Linier", "sks": 3, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4109",
        "deskripsi": "Mata kuliah ini membahas formulasi model program linier, metode grafik, metode Simpleks, dualitas, analisis sensitivitas, masalah transportasi, dan masalah penugasan.",
        "pustaka": ["Taha, H. A. (2017). Operations Research: An Introduction. Pearson.", "Hillier, F. S. & Lieberman, G. J. (2015). Introduction to Operations Research. McGraw-Hill."],
        "topics": [
            "Formulasi Model Matematika Program Linier (Fungsi Tujuan & Kendala)",
            "Metode Grafik untuk Masalah 2 Variabel (Daerah Feasible & Titik Optimal)",
            "Bentuk Standar Program Linier dan Variabel Slack/Surplus",
            "Metode Simpleks: Konsep Solusi Basis Feasible dan Tabel Simpleks",
            "Metode Simpleks Maksimasi dan Minimasi",
            "Metode Dua Tahap (Two-Phase Method) dan Metode Big-M untuk Kendala SAMA DENGAN / LEBIH DARI",
            "Kasus-Kasus Khusus Simpleks: Degenerasi, Solusi Tak Terbatas, Tak Layak, Solusi Ganda",
            "Teori Dualitas: Formulasi Masalah Dual dari Masalah Primal",
            "Teorema Dualitas Duality Theorem dan Complementary Slackness",
            "Analisis Sensitivitas: Perubahan Koefisien Fungsi Tujuan dan Nilai Kanan Kendala",
            "Masalah Transportasi: Metode NWCR, Minimum Cost, VAM (Vogel)",
            "Uji Optimalitas Transportasi: Metode MODI (Modified Distribution) & Stepping Stone",
            "Masalah Penugasan (Assignment Problem) dan Algoritma Hungarian",
            "Aplikasi Program Linier dalam Pengalokasian Sumber Daya Industri dan Bisnis"
        ]
    },
    {
        "kode_mk": "PMA4245", "nama_mk": "Fisika Dasar", "sks": 3, "sem": 5, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang besaran & satuan, kinematika, dinamika (Hukum Newton), usaha & energi, momentum & impuls, rotasi benda tegar, serta gelombang & termodinamika dasar.",
        "pustaka": ["Halliday, D., Resnick, R., & Walker, J. (2014). Fundamentals of Physics. Wiley.", "Serway, R. A. & Jewett, J. W. (2018). Physics for Scientists and Engineers."],
        "topics": [
            "Besaran, Satuan SI, Dimensi, Vektor, dan Analisis Komponen Vektor",
            "Kinematika Gerak Lurus (GLB, GLBB) dan Gerak Parabola/Peluru",
            "Kinematika Gerak Melingkar Beraturan dan Percepatan Sentripetal",
            "Dinamika Gerak: Hukum-Hukum Newton tentang Gerak dan Gaya Gesek",
            "Konsep Usaha, Energi Kinetik, Energi Potensial, dan Hukum Kekekalan Energi Mekanik",
            "Momentum Linier, Impuls, dan Hukum Kekekalan Momentum pada Tumbukan",
            "Rotasi Benda Tegar: Torsi, Momen Inersia, dan Dinamika Rotasi",
            "Kesetimbangan Benda Tegar dan Titik Berat",
            "Fluida Statis (Tekanan Hydrostatis, Hukum Pascal, Archimedes) dan Fluida Dinamis",
            "Getaran Harmonis Sederhana (Ayunan Pendulum & Sistem Pegas)",
            "Gelombang Mekanik: Sifat Gelombang, Bunyi, dan Efek Doppler",
            "Termodinamika: Suhu, Kalor, Pemuaian, dan Hukum Ke-0 & Ke-1 Termodinamika",
            "Hukum Ke-2 Termodinamika, Mesin Carnot, dan Entropi",
            "Aplikasi Konsep Fisika Dasar dan Matematika dalam Pemecahan Masalah Sains"
        ]
    },
    {
        "kode_mk": "PMA4246", "nama_mk": "Biologi Umum", "sks": 2, "sem": 5, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas hakikat kehayatan, struktur dan fungsi sel, metabolisme, genetika molekuler, keanekaragaman hayati, ekologi, serta pemodelan matematika dalam populasi biologi.",
        "pustaka": ["Campbell, N. A., et al. (2017). Biology (11th ed.). Pearson.", "Kimball, J. W. (2010). Biologi. Jakarta: Erlangga."],
        "topics": [
            "Hakikat Biologi sebagai Ilmu dan Metode Ilmiah Sains",
            "Struktur dan Fungsi Sel (Organel Sel Prokariotik & Eukariotik)",
            "Transportasi Membran Sel (Difusi, Osmosis, Transpor Aktif)",
            "Metabolisme Sel: Enzim, Fotosintesis, dan Respirasi Aerob/Anaerob",
            "Genetika Molekuler: DNA, RNA, Replikasi, dan Sintesis Protein",
            "Pembelahan Sel (Mitosis & Meiosis) dan Hukum Hereditas Mendel",
            "Struktur dan Fungsi Organ Tubuh Manusia (Sistem Syaraf & Hormon)",
            "Keanekaragaman Hayati (Biodiversitas) dan Klasifikasi Makhluk Hidup",
            "Ekologi: Konsep Ekosistem, Rantai Makanan, dan Daur Biogeokimia",
            "Interaksi Organisme dan Dinamika Populasi Lingkungan",
            "Teori Evolusi dan Mekanisme Seleksi Alam",
            "Bioteknologi Modern dan Rekayasa Genetika",
            "Pemodelan Matematika dalam Biologi (Pertumbuhan Populasi & Penyebaran Penyakit)",
            "Isu-Isu Lingkungan Global dan Konservasi Keanekaragaman Hayati"
        ]
    },
    {
        "kode_mk": "PMA4247", "nama_mk": "Kimia Dasar", "sks": 2, "sem": 5, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas struktur atom, tabel periodik, ikatan kimia, stoikiometri, wujud zat, termokimia, kesetimbangan kimia, laju reaksi, dan larutan asam-basa.",
        "pustaka": ["Chang, R. & Goldsby, K. A. (2016). Chemistry. McGraw-Hill.", "Petrucci, R. H., et al. (2017). General Chemistry: Principles and Modern Applications. Pearson."],
        "topics": [
            "Hakikat Ilmu Kimia, Materi, Sifat Fisika/Kimia, dan Perubahannya",
            "Struktur Atom: Partikel Dasar, Nomor Atom, Nomor Massa, dan Isotop",
            "Teori Atom Modern, Konfigurasi Elektron, dan Tabel Periodik Unsur",
            "Ikatan Kimia: Ikatan Ionik, Ikatan Kovalen, dan Ikatan Logam",
            "Bentuk Molekul, Kepolaran, dan Gaya Antar-Molekul",
            "Stoikiometri: Konsep Mol, Massa Molar, dan Persamaan Reaksi Kimia",
            "Hitungan Kimia: Pereaksi Pembatas, Rumus Empiris, dan Rumus Molekul",
            "Wujud Zat: Sifat Gas Ideal (Hukum Gas), Cairan, dan Padatan",
            "Termokimia: Perubahan Entalpi Reaksi, Hukum Hess, dan Energi Ikatan",
            "Laju Reaksi: Orde Reaksi, Persamaan Laju Reaksi, dan Teorema Tumbukan",
            "Kesetimbangan Kimia: Tetapan Kesetimbangan (Kc, Kp) dan Pergeseran Le Chatelier",
            "Larutan Asam-Basa: Teori Asam-Basa, pH Larutan, dan Larutan Penyangga (Buffer)",
            "Reaksi Redoks dan Dasar-Dasar Elektrokimia (Sel Volta & Elektrolisis)",
            "Aplikasi Perhitungan Matematika dalam Kimia Analitis dan Industri"
        ]
    },
    {
        "kode_mk": "PMA4236", "nama_mk": "Nilai Awal & Syarat Batas", "sks": 2, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4107",
        "deskripsi": "Mata kuliah ini membahas tentang Masalah Nilai Awal (MNA), Masalah Syarat Batas (MSB), persamaan diferensial parsial (PDP) linier, deret Fourier, dan metode pemisahan variabel.",
        "pustaka": ["Haberman, R. (2012). Applied Partial Differential Equations with Fourier Series and Boundary Value Problems. Pearson.", "Zill, D. G. (2018). Advanced Engineering Mathematics. Jones & Bartlett."],
        "topics": [
            "Review Persamaan Diferensial Biasa dan Masalah Nilai Awal (MNA)",
            "Konsep Masalah Syarat Batas (MSB) dan Teorema Keberadaan Solusi",
            "Masalah Sturm-Liouville dan Nilai Eigen / Fungsi Eigen MSB",
            "Pengenalan Persamaan Diferensial Parsial (PDP) Orde Dua: Gelombang, Panas, Laplace",
            "Deret Fourier: Koefisien Fourier, Deret Sinus, dan Deret Kosinus Fourier",
            "Kekonvergenan Deret Fourier dan Teorema Dirichlet",
            "Metode Pemisahan Variabel untuk Persamaan Panas Dimensi Satu (1D Heat Equation)",
            "Penyelesaian Persamaan Gelombang Dimensi Satu (1D Wave Equation)",
            "Penyelesaian Persamaan Laplace pada Daerah Persegi Panjang dan Lingkaran",
            "Penggunaan Syarat Batas Homogen dan Non-Homogen",
            "Transformasi Fourier dan Aplikasinya pada PDP Domain Tak Terbatas",
            "Transformasi Laplace untuk Penyelesaian PDP dengan Masalah Nilai Awal",
            "Aplikasi MSB dan PDP dalam Fenomena Fisika dan Teknik",
            "Simulasi Solusi Nilai Awal & Syarat Batas Menggunakan Komputasi Numerik"
        ]
    },
    {
        "kode_mk": "PMA4240", "nama_mk": "Pemodelan Matematika", "sks": 2, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas siklus pemodelan matematika, model pertumbuhan populasi, model interaksi spesies (Lotka-Volterra), model epidemik (SIR), analisis kestabilan, dan simulasi.",
        "pustaka": ["Giordano, F. R., Fox, W. P., & Horton, S. B. (2014). A First Course in Mathematical Modeling. Cengage Learning.", "Edelstein-Keshet, L. (2005). Mathematical Models in Biology. SIAM."],
        "topics": [
            "Hakikat Pemodelan Matematika dan Tahapan Siklus Pemodelan",
            "Klasifikasi Model Matematika: Deterministik vs Stokastik, Kontinu vs Diskret",
            "Pemodelan Persamaan Beda Diskret: Pertumbuhan Populasi Tunggal",
            "Pemodelan Logistik Diskret dan Fenomena Chaos",
            "Pemodelan Kontinu: Pertumbuhan Eksponensial dan Logistik (Malthus & Verhulst)",
            "Analisis Titik Kesetimbangan (Equilibrium Points) dan Kestabilan Linier",
            "Model Interaksi Dua Spesies: Predator-Prey (Model Lotka-Volterra)",
            "Model Kompetisi Antar Spesies dan Model Mutualisme",
            "Pemodelan Penyebaran Penyakit Menular: Model SIR, SIS, dan SEIR",
            "Perhitungan Bilangan Reproduksi Dasar (R_0) dan Kestabilan Bebas Penyakit",
            "Pemodelan Masalah Fisika: Pendulum, Gerak Jatuh dengan Gesekan Udara",
            "Pemodelan Keuangan dan Ekonomi Sederhana",
            "Validasi Model, Kalibrasi Parameter, dan Analisis Sensitivitas",
            "Simulasi Komputer Pemodelan Matematika Menggunakan Python/Matlab"
        ]
    },
    {
        "kode_mk": "PMA4221", "nama_mk": "Pemrograman Komputer", "sks": 2, "sem": 5, "rumpun": "Teknologi Pendidikan Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas logika pemrograman, struktur data dasar, percabangan, perulangan, fungsi/prosedur, serta pembuatan program komputer (Python/Visual Basic) untuk matematika.",
        "pustaka": ["Python Programming: An Introduction to Computer Science by John Zelle (2016).", "Buku Ajar Pemrograman Python untuk Matematika."],
        "topics": [
            "Pengantar Pemrograman Komputer, Algoritma, dan Flowchart",
            "Pengenalan Bahasa Pemrograman Python/Visual Basic & Lingkungan Kerja (IDE)",
            "Tipe Data Dasar (Integer, Float, String, Boolean), Variabel, dan Operator",
            "Struktur Data Terurut: List, Tuple, Set, dan Dictionary dalam Python",
            "Struktur Kontrol Percabangan: IF, IF-ELSE, dan ELIF",
            "Struktur Kontrol Perulangan: FOR Loop, WHILE Loop, Break, dan Continue",
            "Fungsi (Function) dan Prosedur: Parameter, Return Value, dan Scope Variabel",
            "Modul dan Library Matematika (Math, NumPy, SciPy, Matplotlib)",
            "Pemrograman Berorientasi Objek (OOP) Dasar: Class dan Object",
            "Operasi File (File I/O): Membaca dan Menulis File Teks/CSV",
            "Penanganan Eksepsi (Try-Except Error Handling)",
            "Algoritma Pencarian (Search) dan Pengurutan (Sorting: Bubble, Quick Sort)",
            "Pembuatan GUI Sederhana (Tkinter/PyQt) untuk Aplikasi Pembelajaran Matematika",
            "Proyek Mandiri: Pembuatan Aplikasi Program Pemecah Masalah Matematika"
        ]
    },
    {
        "kode_mk": "PMA4244", "nama_mk": "Kajian Olimpiade Matematika Sekolah", "sks": 2, "sem": 5, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas trik dan strategi pemecahan masalah olimpiade matematika (OSN/KSN) bidang Aljabar, Geometri, Kombinatorika, dan Teori Bilangan tingkat SMP/SMA.",
        "pustaka": ["Zeitz, P. (2007). The Art and Craft of Problem Solving. Wiley.", "Buku Panduan Pembinaan Olimpiade Matematika SMA."],
        "topics": [
            "Karakteristik Soal Olimpiade Matematika (OSN/KSN) dan Strategi Problem Solving",
            "Aljabar Olimpiade: Ketaksamaan Klasik (AM-GM, Cauchy-Schwarz, Rearrangement)",
            "Aljabar Olimpiade: Polinomial, Suku Banyak, Teorema Sisa, dan Persamaan Fungsional",
            "Aljabar Olimpiade: Manipulasi Aljabar Kompleks dan Sistem Persamaan Non-Linier",
            "Teori Bilangan Olimpiade: Keterbagian Lanjut, Kongruensi Modulo, dan Persamaan Diophantine",
            "Teori Bilangan Olimpiade: Fungsi Aritmatika, Ord Modulo, dan Resiprokositas Kuadrat",
            "Geometri Olimpiade: Geometri Segitiga Lanjut, Garis-Garis Istimewa, dan Titik Berat/Tinggi/Kecil",
            "Geometri Olimpiade: Siklik Segiempat (Cyclic Quadrilateral) dan Teorema Ceva/Menelaus",
            "Geometri Olimpiade: Transformasi Geometri dan Geometri Analitik dalam Pembuktian OSN",
            "Kombinatorika Olimpiade: Prinsip Inklusi-Eksklusi Lanjut & Pigeonhole Principle Lanjut",
            "Kombinatorika Olimpiade: Bipartite Graph, Invariant, dan Colorings",
            "Kombinatorika Olimpiade: Counting Techniques dan Rekursi Lanjut",
            "Simulasi Analisis dan Pembahasan Soal OSN Tingkat Kabupaten/Kota & Provinsi",
            "Penyusunan Modul Pembinaan Pembimbingan Olimpiade Matematika Sekolah"
        ]
    },
    {
        "kode_mk": "PMA4211", "nama_mk": "Struktur Aljabar II", "sks": 3, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4110",
        "deskripsi": "Mata kuliah ini membahas tentang Ring (Gelanggang), Subring, Ideal, Ring Kuosien, Homomorfisme Ring, Integral Domain, Field (Lapangan), dan Ring Polinomial.",
        "pustaka": ["Gallian, J. A. (2017). Contemporary Abstract Algebra. Cengage Learning.", "Hungerford, T. W. (2014). Abstract Algebra: An Introduction. Brooks/Cole."],
        "topics": [
            "Definisi Ring (Gelanggang) dan Sifat-Sifat Aritmatika Dasar Ring",
            "Klasifikasi Ring: Ring Komutatif, Ring Elemen Satuan, dan Contoh Ring Z, Q, R, C, M_n",
            "Pembagi Nol (Zero Divisors) dan Daerah Integral (Integral Domain)",
            "Ring Pembagi (Division Ring) dan Lapangan (Field)",
            "Definisi Subring dan Kriteria Pengujian Subring",
            "Ideal Kiri, Ideal Kanan, dan Ideal Dua Sisi pada Ring",
            "Ring Kuosien / Ring Faktor (R/I) dan Operasi Tambah/Kali Koset Ring",
            "Ideal Utama (Principal Ideal), Ideal Maksimal, dan Ideal Prima",
            "Homomorfisme Ring, Kernel Ring, dan Isomorfisme Ring",
            "Teorema Isomorfisme Utama bagi Ring (First Isomorphism Theorem for Rings)",
            "Ring Polinomial R[x] dan Operasi Polinomial di atas Field",
            "Algoritma Pembagian Polinomial dan Uji Ketakreduksian Polinomial (Eisenstein Criterion)",
            "Domain Ideal Utama (PID) dan Domain Faktorisasi Tunggal (UFD)",
            "Pengenalan Lapangan Perluasan (Field Extensions) dan Teori Galois Dasar"
        ]
    },
    {
        "kode_mk": "PMA4231", "nama_mk": "Analisis Variabel Kompleks", "sks": 3, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4115",
        "deskripsi": "Mata kuliah ini membahas sistem bilangan kompleks, fungsi analitik, persamaan Cauchy-Riemann, fungsi elementer, integral kompleks, Teorema Cauchy-Goursat, deret Laurent, dan Residu.",
        "pustaka": ["Churchill, R. V. & Brown, J. W. (2014). Complex Variables and Applications. McGraw-Hill.", "Zill, D. G. & Shanahan, P. D. (2013). A First Course in Complex Analysis."],
        "topics": [
            "Sistem Bilangan Kompleks: Geometri, Modulus, Konjuget, dan Bentuk Polar/Eksponensial",
            "Akar Bilangan Kompleks dan Teorema De Moivre",
            "Topologi di Bidang Kompleks (Lingkungan, Titik Batas, Himpunan Buka/Tutup, Domain)",
            "Fungsi Kompleks: Limit dan Kekontinuan Fungsi Variabel Kompleks",
            "Turunan Fungsi Kompleks dan Persamaan Cauchy-Riemann (Kondisi C-R)",
            "Fungsi Analitik dan Fungsi Harmonik serta Konjuget Harmonik",
            "Fungsi Elementer Kompleks: Eksponensial, Logaritma, Trigonometri, dan Hiperbolik",
            "Lintasan (Contour) di Bidang Kompleks dan Pengenalan Integral Kompleks",
            "Integral Kontur dan Sifat-Sifat Dasar Pengintegralan Kompleks",
            "Teorema Cauchy-Goursat dan Independensi Lintasan Integral Kompleks",
            "Rumus Integral Cauchy (Cauchy's Integral Formula) dan Turunan Fungsi Analitik",
            "Deret Taylor dan Deret Laurent untuk Fungsi Kompleks",
            "Titik Singularitas, Pole, dan Perhitungan Residu",
            "Teorema Residu dan Aplikasinya dalam Perhitungan Integral Real Tak Wajar"
        ]
    },
    {
        "kode_mk": "PMA4233", "nama_mk": "Teori Koding", "sks": 2, "sem": 5, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4109",
        "deskripsi": "Mata kuliah ini membahas saluran komunikasi, jarak Hamming, bobot Hamming, kode linier, matriks pembangkit, matriks cek-paritas, pengodean/pendekodean, dan kode siklik.",
        "pustaka": ["Hill, R. (1986). A First Course in Coding Theory. Oxford University Press.", "Ling, S. & Xing, C. (2004). Coding Theory: A First Course. Cambridge."],
        "topics": [
            "Pengantar Sistem Komunikasi Digital dan Kebutuhan Koreksi Kesalahan Data",
            "Konsep Dasar Pengodean (Encoding) dan Dekoding (Decoding)",
            "Jarak Hamming (Hamming Distance), Bobot Hamming (Hamming Weight), dan Sifat Metrik",
            "Batas Kemampuan Deteksi Kesalahan (Error-Detecting) dan Koreksi Kesalahan (Error-Correcting)",
            "Konsep Kode Linier (Linear Codes) di atas Field Terhingga GF(q)",
            "Matriks Pembangkit (Generator Matrix) dan Konstruksi Kode Linier",
            "Matriks Cek Paritas (Parity-Check Matrix) dan Dual Code",
            "Teknik Dekoding Sederhana: Syndrome Decoding dan Tabel Standard Array",
            "Batas-Batas Kode: Batas Hamming (Sphere-Packing Bound) dan Kode Sempurna (Perfect Codes)",
            "Kode Hamming (Hamming Codes): Konstruksi, Sifat, dan Proses Dekoding",
            "Kode Siklik (Cyclic Codes): Polinomial Pembangkit dan Polinomial Cek",
            "Konstruksi Kode Siklik Menggunakan Ideal pada Ring Kuosien F[x]/(x^n - 1)",
            "Pengenalan Kode BCH dan Kode Reed-Solomon",
            "Aplikasi Teori Koding dalam Komunikasi Nirkabel, Memori Komputer, dan CD/DVD"
        ]
    },
    {
        "kode_mk": "PMA3106", "nama_mk": "Kapita Selekta Matematika II", "sks": 2, "sem": 5, "rumpun": "Pembelajaran Matematika", "prasyarat": "PMA3105",
        "deskripsi": "Mata kuliah ini membahas bedah materi matematika sekolah menengah atas (SMA/MA/SMK) secara mendalam, penyelesaian soal matematika tingkat lanjut, miskonsepsi, dan pengembangan soal HOTS.",
        "pustaka": ["Buku Teks Matematika SMA/MA Kurikulum Merdeka Kemendikbudristek.", "Soedjadi, R. (2000). Kiat Pendidikan Matematika di Indonesia."],
        "topics": [
            "Bedah Kurikulum Matematika SMA/MA/SMK (Struktur Elemen Matematika Tingkat Lanjut)",
            "Pendalaman Materi Persamaan & Pertidaksamaan Rasional, Irasional, dan Nilai Mutlak",
            "Pendalaman Materi Fungsi Eksponensial dan Logaritma beserta Aplikasinya",
            "Pendalaman Materi Trigonometri SMA: Identitas, Persamaan, dan Aturan Sinus/Kosinus",
            "Pendalaman Materi Matriks dan Sistem Persamaan Linier Tiga Variabel (SPLTV)",
            "Pendalaman Materi Barisan dan Deret (Aritmatika, Geometri, Tak Hingga, Anuitas)",
            "Pengembangan Soal HOTS Topik Aljabar & Trigonometri SMA",
            "Pendalaman Materi Geometri Analitik SMA: Persamaan Lingkaran dan Garis Singgung",
            "Pendalaman Materi Transformasi Geometri SMA (Pendekatan Matriks)",
            "Pendalaman Materi Dimensi Tiga (Jarak dan Sudut dalam Ruang)",
            "Pendalaman Materi Kalkulus SMA: Limit, Turunan, dan Integral Fungsional",
            "Pendalaman Materi Statistika dan Kombinatorika/Peluang SMA",
            "Analisis Miskonsepsi Siswa SMA pada Topik Kalkulus & Geometri Ruang",
            "Penyusunan Paket Soal Try-Out / UTBK-SNBT Matematika dan Pembahasannya"
        ]
    },
    {
        "kode_mk": "PMA4251", "nama_mk": "Pendidikan Multikultural", "sks": 2, "sem": 5, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang konsep pendidikan multikultural, keberagaman sosial-budaya Indonesia, demokrasi pendidikan, pembelajaran inklusif, dan integrasinya dalam pendidikan matematika.",
        "pustaka": ["Banks, J. A. (2019). An Introduction to Multicultural Education. Pearson.", "Mahfud, C. (2016). Pendidikan Multikultural. Yogyakarta: Pustaka Pelajar."],
        "topics": [
            "Hakikat dan Latar Belakang Pendidikan Multikultural di Indonesia",
            "Pluralitas dan Keberagaman Bangsa Indonesia (Etnis, Agama, Ras, Antargolongan)",
            "Konsep Dasar dan Teori-Teori Pendidikan Multikultural (James Banks, dll.)",
            "Demokrasi Pendidikan, Kesetaraan Gender, dan Hak Asasi Manusia",
            "Prinsip-Prinsip Pembelajaran Inklusif dan Pengurangan Prasangka (Prejudice Reduction)",
            "Peran Sekolah sebagai Lembaga Sosial Multikultural",
            "Pengembangan Kurikulum Berperspektif Multikultural",
            "Integrasi Konten Multikultural dalam Pembelajaran Sains dan Matematika",
            "Strategi Pembelajaran Adaptif bagi Siswa Beragam Latar Belakang Sosial",
            "Pengelolaan Konflik Sosial-Budaya di Lingkungan Sekolah",
            "Peran Guru sebagai Agen Perubahan dan Teladan Toleransi",
            "Pendidikan Multikultural dalam Menghadapi Era Globalisasi dan Etnosentrisme",
            "Evaluasi Pembelajaran Berwawasan Multikultural",
            "Rancangan Proyek Pembelajaran Inklusif Berwawasan Keindonesiaan"
        ]
    },

    # SEMESTER 6
    {
        "kode_mk": "UNV5107", "nama_mk": "KKN (Kuliah Kerja Nyata)", "sks": 3, "sem": 6, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini merupakan kegiatan pengabdian masyarakat berbasis keilmuan, observasi lapangan, penyusunan program kerja desa, pemberdayaan masyarakat, dan pembuatan laporan KKN.",
        "pustaka": ["Panduan Pelaksanaan Kuliah Kerja Nyata (KKN) UNIROW Tuban.", "Soetomo. (2014). Strategi Pembangunan Masyarakat. Yogyakarta: Pustaka Pelajar."],
        "topics": [
            "Pengenalan dan Pembekalan Visi-Misi KKN Tematik UNIROW Tuban",
            "Teknik Observasi Lapangan, Pemetaan Potensi Desa, dan Analisis SWOT Desa",
            "Penyusunan Rencana Program Kerja (Proprok) KKN Berbasis Kemasyarakatan",
            "Komunikasi Publik, Advokasi, dan Pendekatan Tokoh Masyarakat Desa",
            "Program Bidang Pendidikan: Pengajaran & Bimbingan Belajar Anak Desa",
            "Program Bidang Ekonomi & Kewirausahaan: Pemberdayaan UMKM Desa",
            "Program Bidang Lingkungan & Kesehatan: Sanitasi, Adiwiyata, & Penghijauan",
            "Program Bidang Teknologi & Informasi: Digitalisasi Desa / Website Desa",
            "Pelaksanaan Program Kerja KKN Berkelanjutan dan Monitoring Harian",
            "Pengelolaan Konflik dan Dinamika Kelompok KKN di Lapangan",
            "Evaluasi Dampak Program Kerja terhadap Masyarakat Sasaran",
            "Penyusunan Draft Laporan Akhir KKN dan Artikel Pengabdian Masyarakat",
            "Penyelenggaraan Lokakarya / Expo Hasil Karya KKN Desa",
            "Ujian Evaluasi dan Presentasi Pertanggungjawaban KKN"
        ]
    },
    {
        "kode_mk": "PMA4248", "nama_mk": "Microteaching", "sks": 2, "sem": 6, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas penguasaan 8 keterampilan dasar mengajar (membuka/menutup, bertanya, memberi penguatan, variasi, menjelaskan, membimbing diskusi, mengelola kelas, mengajar kelompok kecil/perorangan) melalui praktik mengajar terbatas.",
        "pustaka": ["Helmiati. (2013). Micro Teaching: Melatih Keterampilan Dasar Mengajar. Yogyakarta: Akademi Manajemen Perusahaan.", "Turney, C. (1983). Sydney Micro Skills."],
        "topics": [
            "Hakikat Microteaching dan Penguasaan Keterampilan Dasar Mengajar",
            "Keterampilan Membuka dan Menutup Pembelajaran Matematika (Set Induction & Closure)",
            "Keterampilan Menjelaskan Materi Matematika secara Sistematis dan Jelas",
            "Keterampilan Bertanya Dasar dan Bertanya Lanjut (Questioning Skills)",
            "Keterampilan Memberikan Penguatan (Reinforcement Skills) Verbal dan Non-Verbal",
            "Keterampilan Mengadakan Variasi (Variation Skills) Stimulus, Media, dan Pola Interaksi",
            "Keterampilan Mengelola Kelas (Classroom Management Skills) dan Pengendalian Disiplin",
            "Keterampilan Membimbing Diskusi Kelompok Kecil dan Pembelajaran Kooperatif",
            "Keterampilan Mengajar Kelompok Kecil dan Perorangan (Individualized Instruction)",
            "Penyusunan Rencana Pelaksanaan Pembelajaran Micro (RPP Micro / Modul Ajar 15-20 Menit)",
            "Praktik Microteaching Mandiri Skenario 1 (Keterampilan Dasar Mengajar)",
            "Praktik Microteaching Mandiri Skenario 2 (Integrasi Media & Alat Peraga Matematika)",
            "Praktik Microteaching Mandiri Skenario 3 (Pembelajaran Inovatif Berbasis HOTS)",
            "Peer-Assessment, Umpan Balik Dosen/Teman Sebaya, dan Refleksi Diri Guru"
        ]
    },
    {
        "kode_mk": "PMA3111", "nama_mk": "Pengetahuan Lingkungan", "sks": 2, "sem": 6, "rumpun": "Mata Kuliah Umum", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas Sekolah Adiwiyata, isu-isu lingkungan hidup global & lokal, pelestarian ekosistem, pengelolaan limbah/sampah, tanaman obat (TOGA), dan pendidikan berwawasan lingkungan.",
        "pustaka": ["UU No. 32 Tahun 2009 tentang Perlindungan dan Pengelolaan Lingkungan Hidup.", "Soerjani, M. (2007). Lingkungan: Sumberdaya Alam dan Kependudukan dalam Pembangunan."],
        "topics": [
            "Hakikat Pengetahuan Lingkungan dan Ekologi Manusia",
            "Isu-Isu Lingkungan Global: Pemanasan Global, Perubahan Iklim, Kerusakan Ozon",
            "Isu-Isu Lingkungan Lokal & Regional (Pencemaran Air, Udara, Tanah di Jawa Timur)",
            "Konsep Sekolah Adiwiyata dan Gerakan Peduli & Berbudaya Lingkungan Hidup di Sekolah",
            "Pelestarian Fungsi Lingkungan Hidup dan Pencegahan Pencemaran",
            "Pengelolaan Sampah dan Limbah Rumah Tangga (Prinsip 3R: Reduce, Reuse, Recycle)",
            "Praktik Komposting (Pembuatan Pupuk Kompos Organik)",
            "Praktik Kerajinan Kreatif Daur Ulang Limbah Plastik/Kertas",
            "Pemanfaatan Tanaman Obat Keluarga (TOGA) untuk Kesehatan Masyarakat",
            "Praktik Menanam dan Merawat TOGA di Lingkungan Kampus/Sekolah",
            "Praktik Pengolahan Minuman Herbal Kesehatan dari TOGA",
            "Aksi Komunitas Berbasis Kebersihan dan Konservasi Air/Energi",
            "Integrasi Pendidikan Lingkungan Hidup (PLH) dalam Pembelajaran Sekolah",
            "Penyusunan Laporan Proyek Aksi Lingkungan Hidup Berkelanjutan"
        ]
    },
    {
        "kode_mk": "PMA4103", "nama_mk": "Aplikasi Komputer Matematika", "sks": 2, "sem": 6, "rumpun": "Teknologi Pendidikan Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas penggunaan software matematika (GeoGebra, Maple/MATLAB) untuk pemecahan masalah komputasi numerik, aljabar, kalkulus, geometri, serta pembuatan animasi dinamis.",
        "pustaka": ["Hohenwarter, M. (2018). GeoGebra Manual & Teaching Resources.", "Maple User Manual: Maplesoft."],
        "topics": [
            "Pengenalan Software Komputasi Matematika (GeoGebra, Maple, MATLAB)",
            "GeoGebra Dasar: Antarmuka, Input Bar, Toolbars, dan Konstruksi Geometri Datar",
            "GeoGebra Lanjut: Konstruksi Geometri Dinamis, Slider, dan Lokus Titik",
            "GeoGebra 3D Graphics: Pemodelan Bangun Ruang, Irisan Vektor, dan Permukaan",
            "GeoGebra CAS (Computer Algebra System): Operasi Aljabar, Turunan, dan Integral",
            "Pembuatan Applet Interaktif GeoGebra untuk Pembelajaran Matematika Sekolah",
            "Pengenalan Maple: Komputasi Numerik, Eksak, dan Manipulasi Simbolik",
            "Penyelesaian Sistem Persamaan Linier dan Aljabar Matriks dengan Maple",
            "Visualisasi Grafik Fungsi 2D dan 3D Menggunakan Maple",
            "Penyelesaian Kalkulus (Limit, Turunan, Integral, Deret) dengan Maple",
            "Penyelesaian Persamaan Diferensial dengan Maple",
            "Pembuatan Animasi Dinamis Grafik Matematika untuk Simulasi Konsep",
            "Integrasi GeoGebra/Maple dalam Lembar Kerja Peserta Didik (LKPD Digital)",
            "Proyek Mandiri: Pembuatan Media Komputasi Matematika Inovatif"
        ]
    },
    {
        "kode_mk": "PMA4226", "nama_mk": "Metode Numerik", "sks": 2, "sem": 6, "rumpun": "Matematika Lanjutan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas analisis galat, penyelesaian persamaan non-linier (Biseksi, Regula Falsi, Newton-Raphson), SPL numerik, Interpolasi, Diferensiasi & Integrasi numerik.",
        "pustaka": ["Chapra, S. C. & Canale, R. P. (2015). Numerical Methods for Engineers. McGraw-Hill.", "Munir, R. (2015). Metode Numerik. Bandung: Informatika."],
        "topics": [
            "Hakikat Metode Numerik vs Metode Analitis dan Konsep Galat (Error)",
            "Jenis-Jenis Galat: Galat Pembulatan (Rounding), Galat Pemotongan (Truncation), Galat Relatif",
            "Solusi Persamaan Non-Linier: Metode Biseksi (Bisection Method)",
            "Solusi Persamaan Non-Linier: Metode Regula Falsi dan Metode Secant",
            "Solusi Persamaan Non-Linier: Metode Newton-Raphson dan Iterasi Titik Tetap",
            "Penyelesaian SPL Numerik Direct: Eliminasi Gauss dan Dekomposisi LU",
            "Penyelesaian SPL Numerik Iteratif: Metode Jacobi dan Gauss-Seidel",
            "Interpolasi Polinomial: Interpolasi Linier, Kuadrat, dan Polinomial Lagrange",
            "Interpolasi Beda Hingga Newton (Forward & Backward Differences)",
            "Diferensiasi Numerik: Beda Maju, Beda Mundur, dan Beda Terpusat",
            "Integrasi Numerik: Metode Trapesium (Trapezoidal Rule) Tunggal dan Ganda",
            "Integrasi Numerik: Metode Simpson 1/3 dan Simpson 3/8",
            "Penyelesaian PDB Numerik: Metode Euler dan Metode Runge-Kutta Orde 4",
            "Implementasi Program Metode Numerik Menggunakan Python/Excel"
        ]
    },
    {
        "kode_mk": "PMA4228", "nama_mk": "Matematika Terapan", "sks": 2, "sem": 6, "rumpun": "Matematika Lanjutan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas aplikasi matematika (Kalkulus, Aljabar, Persamaan Diferensial, Graf) dalam fenomena nyata kehidupan sehari-hari, ekonomi, sains, dan teknologi.",
        "pustaka": ["Howison, S. (2005). Practical Applied Mathematics. Cambridge University Press.", "Bender, E. A. (2000). An Introduction to Mathematical Modeling. Dover."],
        "topics": [
            "Hakikat Matematika Terapan dan Peranannya dalam Perkembangan Teknologi",
            "Penerapan Konsep Aljabar dan Aritmatika Sosial dalam Ekonomi & Perbankan",
            "Penerapan Kalkulus dalam Pemodelan Optimalisasi Biaya dan Keuntungan Industri",
            "Penerapan Persamaan Diferensial dalam Kinematika dan Dinamika Gerak Benda",
            "Penerapan Persamaan Diferensial dalam Peluruhan Kimia dan Hukum Pendinginan",
            "Penerapan Geometri Transformasi dalam Pemrosesan Citra Digital dan Desain Grafis",
            "Penerapan Aljabar Linier dalam Algoritma Pencarian Google (PageRank)",
            "Penerapan Teori Graf dalam Penentuan Rute Terpendek GPS dan Jaringan Distribusi Logistik",
            "Penerapan Statistika & Peluang dalam Analisis Risiko Asuransi dan Pasar Saham",
            "Penerapan Matematika Diskrit dalam Kriptografi dan Keamanan Siber",
            "Penerapan Trigonometri dalam Navigasi Pelayaran dan Astronomi",
            "Penerapan Metode Numerik dalam Pemodelan Cuaca dan Simulasi Teknik",
            "Studi Kasus Matematika Terapan pada Industri Lokal / UMKM Tuban",
            "Penyusunan Makalah Proyek Aplikasi Matematika Terapan"
        ]
    },
    {
        "kode_mk": "PMA4120", "nama_mk": "Statistika Matematika II", "sks": 3, "sem": 6, "rumpun": "Matematika Lanjutan", "prasyarat": "PMA4119",
        "deskripsi": "Mata kuliah ini membahas tentang distribusi fungsi peubah acak, distribusi sampling (t, F, Chi-Square), Teorema Limit Pusat, estimasi titik & selang, dan pengujian hipotesis matematis.",
        "pustaka": ["Hogg, R. V., McKean, J., & Craig, A. T. (2019). Introduction to Mathematical Statistics. Pearson.", "Mood, A. M., Graybill, F. A., & Boes, D. C. (2011). Introduction to the Theory of Statistics."],
        "topics": [
            "Teknik Penentuan Distribusi Fungsi Peubah Acak: Metode Fungsi Distribusi Akumulatif",
            "Teknik Penentuan Distribusi Fungsi Peubah Acak: Metode Transformasi Peubah Acak",
            "Teknik Penentuan Distribusi Fungsi Peubah Acak: Metode Fungsi Pembangkit Momen (MGF)",
            "Statistik Sampel Acak, Statistik Rataan, dan Varians Sampel",
            "Distribusi Sampling Khusus: Distribusi Chi-Square (Kaidah Penjumlahan & Sifat)",
            "Distribusi Sampling Khusus: Distribusi t-Student dan Distribusi F-Snedecor",
            "Teorema Limit Pusat (Central Limit Theorem) dan Implikasi Statistik Sampel Besar",
            "Estimasi Titik (Point Estimation): Metode Momen (Method of Moments)",
            "Estimasi Titik: Metode Kemungkinan Maksimum (Maximum Likelihood Estimation - MLE)",
            "Kriteria Estimator Terbaik: Tak Bias (Unbiasedness), Efisiensi, Konsistensi, & Cukup (Sufficiency)",
            "Estimasi Selang (Interval Estimation): Selang Kepercayaan Rataan dan Varians",
            "Pengujian Hipotesis Matematis: Hipotesis Sederhana vs Komposit, Uji Terbaik, Lemma Neyman-Pearson",
            "Uji Rasio Kemungkinan Maksimum (Likelihood Ratio Test - LRT)",
            "Aplikasi Statistika Matematika II dalam Pemodelan Stokastik"
        ]
    },
    {
        "kode_mk": "PMA4249", "nama_mk": "E-learning", "sks": 3, "sem": 6, "rumpun": "Teknologi Pendidikan Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas desain dan penerapan pembelajaran elektronik (e-learning), Learning Management System (LMS Moodle/Google Classroom), konten digital interaktif, dan evaluasi daring.",
        "pustaka": ["Horton, W. (2012). E-Learning by Design. Pfeiffer/Wiley.", "Clark, R. C. & Mayer, R. E. (2016). e-Learning and the Science of Instruction. Wiley."],
        "topics": [
            "Hakikat E-learning, Pembelajaran Sinkron vs Asinkron, dan Blended Learning",
            "Teori Instruksional E-learning (Cognitive Load Theory & Multimedia Learning Mayer)",
            "Pengenalan Learning Management System (LMS): Moodle, Canvas, Google Classroom",
            "Desain dan Struktur Kelas Online / LMS untuk Pelajaran Matematika",
            "Pengembangan Bahan Ajar Digital: E-Book, PDF Interaktif, dan Infografis",
            "Pengembangan Konten Pembelajaran Video Screencast / Video Pembelajaran Matematika",
            "Desain Aktivitas Forum Diskusi Daring dan Kolaborasi Siswa di LMS",
            "Pengembangan Alat Asesmen Online: Quizziz, Kahoot, Google Forms, & Kuis LMS Auto-Grading",
            "Pengenalan Learning Analytics dan Pelacakan Progres Belajar Siswa di LMS",
            "Integrasi Applet GeoGebra dan Tools Interaktif dalam LMS Moodle",
            "Pengelolaan Aksesibilitas, Etika Daring, dan Siber-Security dalam E-learning",
            "Model Evaluasi Kualitas E-learning (Usability Testing & Pedagogical Effectiveness)",
            "Penyusunan LMS Course Matematika Lengkap (1 Semester Virtual Class)",
            "Uji Coba dan Demonstrasi Sistem Pembelajaran E-learning Matematika Inovatif"
        ]
    },
    {
        "kode_mk": "PMA4111", "nama_mk": "Multimedia Pembelajaran Matematika", "sks": 2, "sem": 6, "rumpun": "Teknologi Pendidikan Matematika", "prasyarat": "PMA4103",
        "deskripsi": "Mata kuliah ini membahas prinsip multimedia interaktif, pengembangan animasi/visual (Articulate Storyline, Canva, Flash/Animate), integrasi audio-video, dan pengujian media pembelajaran matematika.",
        "pustaka": ["Mayer, R. E. (2009). Multimedia Learning. Cambridge University Press.", "Buku Panduan Pengembangan Media Multimedia Interaktif."],
        "topics": [
            "Hakikat Multimedia Pembelajaran Matematika dan Konsep Interaktivitas",
            "12 Prinsip Pembelajaran Multimedia Richard E. Mayer",
            "Pengenalan Tools Alat Pembuat Multimedia: Articulate Storyline, Canva, Adobe Animate",
            "Perancangan Storyboard dan Flowchart Alur Multimedia Pembelajaran",
            "Pengolahan Teks, Tata Letak (Layout), dan Estetika Visual Warna untuk Matematika",
            "Pengolahan Vektor, Gambar, dan Diagram Geometri Interaktif",
            "Pengolahan Audio: Dubbing Voice Over, Backsound, dan Editing Audio Sederhana",
            "Pengolahan Video Tutorial Matematika dan Teknik Chroma Key (Green Screen)",
            "Pembuatan Animasi Geometri dan Konsep Aljabar Interaktif",
            "Pengembangan Kuis Interaktif, Umpan Balik Otomatis, dan Skor dalam Multimedia",
            "Pengemasan Produk Multimedia (Export to HTML5, SCORM, Android APK)",
            "Uji Validitas Ahli Media dan Ahli Materi terhadap Produk Multimedia",
            "Uji Kepraktisan dan Keefektifan Multimedia pada Kelompok Kecil Siswa",
            "Pameran Produk (Exhibition) Multimedia Pembelajaran Matematika Interaktif"
        ]
    },
    {
        "kode_mk": "PMA4250", "nama_mk": "Publikasi Karya Ilmiah", "sks": 2, "sem": 6, "rumpun": "Penelitian Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas etika penulisan karya ilmiah, struktur artikel jurnal (IMRAD), teknik pengutipan & reference manager (Mendeley/Zotero), submit jurnal, dan proses peer-review.",
        "pustaka": ["Belcher, W. L. (2019). Writing Your Journal Article in Twelve Weeks. University of Chicago Press.", "Pedoman Akreditasi Jurnal Ilmiah (ARJUNA) Kemendikbudristek."],
        "topics": [
            "Hakikat Publikasi Karya Ilmiah dan Etika Penulisan Akademik (Anti-Plagiarisme)",
            "Struktur Standar Artikel Jurnal Ilmiah: IMRAD (Introduction, Method, Results, Discussion)",
            "Teknik Menyusun Judul yang Menarik, Abstrak yang Efektif, dan Kata Kunci Relevan",
            "Penyusunan Pendahuluan (Introduction): Gap Analysis, Novelty, dan Tujuan Penelitian",
            "Penyusunan Metode Penelitian yang Jelas dan Dapat Diplikasi (Replicable)",
            "Penyusunan Hasil Penelitian: Penyajian Tabel, Grafik, dan Hasil Uji Statistik",
            "Penyusunan Pembahasan (Discussion): Mengaitkan Hasil dengan Teori & Penelitian Relevan",
            "Penyusunan Kesimpulan, Implikasi, dan Saran Penelitian",
            "Manajemen Referensi Otomatis Menggunakan Mendeley / Zotero (Format APA/IEEE)",
            "Cek Serupa/Plagiarisme Menggunakan Turnitin/Ithenticate dan Teknik Parafrase",
            "Pemilihan Jurnal Sasaran: Menentukan Jurnal Nasional Terakreditasi (SINTA) / Bereputasi",
            "Prosedur Submisi Jurnal (Open Journal Systems - OJS) dan Penyiapan Cover Letter",
            "Proses Peer-Review: Memahami Ulasan Reviewer dan Menyusun Response to Reviewers",
            "Finalisasi Artikel Ilmiah Siap Publish dalam Bidang Pendidikan Matematika"
        ]
    },

    # SEMESTER 7
    {
        "kode_mk": "KIP2701", "nama_mk": "Asesmen Pembelajaran", "sks": 3, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang konsep asesmen (Formatif, Sumatif, Diagnostik), penyusunan kisi-kisi & tes kognitif/afektif/psikomotor, analisis kualitas butir soal (Validitas, Reliabilitas, Tingkat Kesukaran, Daya Beda), dan Asesmen Otentik.",
        "pustaka": ["Arikunto, S. (2015). Dasar-Dasar Evaluasi Pendidikan. Jakarta: Bumi Aksara.", "Nitko, A. J. & Brookhart, S. M. (2014). Educational Assessment of Students. Pearson."],
        "topics": [
            "Hakikat Pengukuran, Penilaian, Evaluasi, dan Asesmen Pembelajaran",
            "Jenis-Jenis Asesmen dalam Kurikulum Merdeka: Asesmen Diagnostik, Formatif, dan Sumatif",
            "Prinsip-Prinsip Asesmen Pembelajaran Matematika yang Adil dan Objektif",
            "Penyusunan Matriks Kisi-Kisi Soal Tes Matematika Sekolah",
            "Penyusunan Soal Tes Hasil Belajar Ranah Kognitif (C1 - C6 Taksonomi Bloom)",
            "Penyusunan Soal Asesmen Berpikir Tingkat Tinggi (HOTS) dan Pemecahan Masalah",
            "Pengembangan Instrumen Penilaian Ranah Afektif (Sikap & Minat Matematika)",
            "Pengembangan Instrumen Penilaian Ranah Psikomotorik / Keterampilan Unjuk Kerja",
            "Konsep Asesmen Otentik: Penilaian Kinerja, Portofolio, dan Penilaian Proyek",
            "Penyusunan Rubrik Analitis dan Rubrik Holistik Penskoran Asesmen",
            "Analisis Kualitas Butir Soal Secara Kualitatif (Construct, Content, Language)",
            "Analisis Kuantitatif Butir Soal: Validitas, Reliabilitas, Tingkat Kesukaran, dan Daya Beda (Anates/Iteman/R)",
            "Analisis Miskonsepsi Siswa Berdasarkan Hasil Asesmen Pembelajaran",
            "Pengolahan Nilai Hasil Belajar, Pelaporan (Rapor), dan Tindak Lanjut Remidial/Pengayaan"
        ]
    },
    {
        "kode_mk": "PMA4252", "nama_mk": "Pengembangan Instrumen Penelitian", "sks": 2, "sem": 7, "rumpun": "Penelitian Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang teknik penyusunan dan pengujian instrumen penelitian kuantitatif/kualitatif (Tes Kemampuan Matematika, Angket Motivasi/Kecemasan, Pedoman Wawancara, Lembar Observasi) serta analisis rasch model.",
        "pustaka": ["Azwar, S. (2016). Penyusunan Skala Psikologi. Yogyakarta: Pustaka Pelajar.", "Sumintono, B. & Widhiarso, W. (2015). Aplikasi Model Rasch untuk Penelitian Ilmu-Ilmu Sosial."],
        "topics": [
            "Hakikat Instrumen Penelitian Pendidikan dan Konstruk Variabel",
            "Tahapan Pengembangan Instrumen: Konseptualisasi, Operasionalisasi, Indikator, dan Butir",
            "Penyusunan Instrumen Tes Kemampuan Berpikir Kritis/Kreatif Matematika",
            "Penyusunan Skala Sikap Psikologis (Likert, Semantic Differential) Kecemasan/Motivasi Matematika",
            "Penyusunan Lembar Observasi Aktivitas Siswa dan Pengelolaan Kelas Guru",
            "Penyusunan Pedoman Wawancara Mendalam (In-Depth Interview) dan Lembar Dokumentasi",
            "Uji Validitas Isi (Content Validity) Menggunakan Judgement Ahli (Gregory / Aiken's V)",
            "Uji Coba Instrumen Penelitian (Field Trial) dan Pembersihan Data",
            "Uji Validitas Konstruk Menggunakan Analisis Faktor (EFA & CFA)",
            "Uji Reliabilitas Skala (Cronbach's Alpha) dan Uji Reliabilitas Tes (Kuder-Richardson / Alpha)",
            "Pengenalan Pemodelan Rasch (Rasch Model) untuk Pengukuran Ilmu Sosial/Pendidikan",
            "Analisis Wright Map, Fit Statistics (Infit/Outfit MNSQ), dan Person/Item Reliability",
            "Penyempurnaan dan Revisi Draft Instrumen Penelitian Siap Pakai",
            "Penyusunan Dokumen Lampiran Instrumen Penelitian Skripsi"
        ]
    },
    {
        "kode_mk": "KIP2702", "nama_mk": "Manajemen dan Pengembangan Program Sekolah", "sks": 2, "sem": 7, "rumpun": "Ilmu Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas Manajemen Berbasis Sekolah (MBS), perencanaan strategis sekolah (RKS/RKAS), kepemimpinan instruksional, pengelolaan sumber daya, dan pengembangan kultur sekolah.",
        "pustaka": ["Mulyasa, E. (2014). Manajemen Berbasis Sekolah. Bandung: Remaja Rosdakarya.", "Bush, T. & Coleman, M. (2012). Leadership and Strategic Management in Education."],
        "topics": [
            "Hakikat Manajemen Pendidikan dan Konsep Manajemen Berbasis Sekolah (MBS)",
            "8 Standar Nasional Pendidikan (SNP) sebagai Acuan Pengembangan Sekolah",
            "Perencanaan Strategis Sekolah: Penyusunan Rencana Kerja Sekolah (RKS) & RKAS",
            "Manajemen Kurikulum dan Program Pembelajaran Sekolah",
            "Manajemen Kesiswaan: Penerimaan Siswa Baru, Bimbingan, dan Ekstrakurikuler",
            "Manajemen Pendidik dan Tenaga Kependidikan (PTK) di Sekolah",
            "Manajemen Sarana dan Prasarana Pendidikan serta Fasilitas Belajar",
            "Manajemen Keuangan dan Pembiayaan Pendidikan (Dana BOS & Akuntabilitas)",
            "Manajemen Hubungan Sekolah dengan Masyarakat (Humas) dan Komite Sekolah",
            "Kepemimpinan Pembelajaran (Instructional Leadership) Kepala Sekolah",
            "Pengembangan Kultur, Iklim, dan Branding Unggulan Sekolah",
            "Supervisi Akademik dan Penjaminan Mutu Internal Sekolah (SPMI)",
            "Manajemen Sistem Informasi Sekolah (Dapodik & Aplikasi Digital)",
            "Penyusunan Draf Rencana Program Pengembangan Sekolah Inovatif"
        ]
    },
    {
        "kode_mk": "KIP2703", "nama_mk": "Pengembangan Bahan Ajar", "sks": 2, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang prosedur perancangan dan penyusunan bahan ajar cetak maupun digital (Buku Ajar, Modul Pembelajaran, LKPD, E-Modul Interaktif) sesuai standar kurikulum.",
        "pustaka": ["Prastowo, A. (2015). Panduan Kreatif Membuat Bahan Ajar Inovatif. Yogyakarta: Diva Press.", "Depdiknas. (2008). Panduan Pengembangan Bahan Ajar. Jakarta."],
        "topics": [
            "Hakikat dan Peran Bahan Ajar dalam Pembelajaran Matematika",
            "Jenis-Jenis Bahan Ajar Cetak dan Non-Cetak (Digital)",
            "Analisis Kurikulum, Capaian Pembelajaran, dan Kebutuhan Bahan Ajar",
            "Prinsip-Prinsip Penyusunan Bahan Ajar: Relevansi, Konsistensi, Kecukupan",
            "Penyusunan Lembar Kerja Peserta Didik (LKPD) Berbasis Masalah (PBL)",
            "Penyusunan LKPD Berbasis Penemuan (Discovery/Inquiry Learning)",
            "Penyusunan Modul Pembelajaran Matematika Terstruktur",
            "Penyusunan Buku Ajar / Handout Pembelajaran Sekolah",
            "Desain Elemen Visual, Ilustrasi Grafis, dan Tata Letak (Typography) Bahan Ajar",
            "Pengembangan E-Modul Interaktif Menggunakan Flip PDF Corporate / Canva",
            "Integrasi Kode QR, Link Video, dan Applet Matematika dalam Bahan Ajar Digital",
            "Kriteria Penilaian Kualitas Bahan Ajar (Kelayakan Isi, Bahasa, Penyajian, Kegrafikan)",
            "Uji Coba Keterbacaan dan Keefektifan Bahan Ajar pada Siswa",
            "Finalisasi Produk Bahan Ajar Matematika Siap Pakai di Sekolah"
        ]
    },
    {
        "kode_mk": "KIP2704", "nama_mk": "Pengembangan Media Pembelajaran", "sks": 3, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang desain instruksional, pembuatan alat peraga matematika konkret, media pembelajaran manipulatif, media digital, dan validasi efektivitas media di sekolah.",
        "pustaka": ["Arsyad, A. (2014). Media Pembelajaran. Jakarta: Rajawali Pers.", "Smaldino, S. E., et al. (2012). Instructional Technology and Media for Learning. Pearson."],
        "topics": [
            "Hakikat Media Pembelajaran dan Landasan Penggunaannya (Kerucut Pengalaman Dale)",
            "Model Desain Instruksional Media Pembelajaran (Model ASSURE & ADDIE)",
            "Klasifikasi Media Pembelajaran: Konkret/Manipulatif, Cetak, Audio-Visual, Digital",
            "Perancangan Alat Peraga Matematika Konkret (Aljabar & Geometri)",
            "Pembuatan Alat Peraga Manipulatif (Blok Dienes, Papan Paku, Alat Peraga Volum)",
            "Pengembangan Media Pembelajaran Berbasis Game (Educational Game & Card Game)",
            "Pengembangan Media Display & Poster Matematika Edukatif",
            "Pengembangan Media Video Pembelajaran Animasi (Powtoon/Renderforest)",
            "Pengembangan Media Pembelajaran Berbasis Mobile App (Smart Apps Creator / MIT App Inventor)",
            "Uji Validasi Ahli Media dan Ahli Materi terhadap Media Pembelajaran",
            "Uji Kepraktisan Penggunaan Media oleh Guru dan Siswa Sekolah",
            "Teknik Integrasi Media dalam Rencana Pelaksanaan Pembelajaran (RPP/Modul Ajar)",
            "Evaluasi Efektivitas Penggunaan Media terhadap Hasil Belajar Siswa",
            "Pameran Gelar Karya (Exhibition) Alat Peraga & Media Pembelajaran Matematika"
        ]
    },
    {
        "kode_mk": "KIP2705", "nama_mk": "Perencanaan Pembelajaran", "sks": 2, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas penyusunan Perangkat Pembelajaran Matematika (Program Tahunan, Program Semester, Alur Tujuan Pembelajaran/ATP, Modul Ajar / RPP Merdeka), dan skenario pembelajaran.",
        "pustaka": ["Majid, A. (2013). Perencanaan Pembelajaran. Bandung: Remaja Rosdakarya.", "Panduan Pembelajaran dan Asesmen Kurikulum Merdeka Kemendikbudristek."],
        "topics": [
            "Hakikat Perencanaan Pembelajaran dan Prinsip Penyusunan Perangkat PBM",
            "Analisis Dokumen Kurikulum (Capaian Pembelajaran - CP, Elemen, dan Fase)",
            "Penyusunan Alur Tujuan Pembelajaran (ATP) dan Peta Konsep Pembelajaran",
            "Perhitungan Hari Efektif, Jam Efektif, dan Penyusunan Program Tahunan (Prota)",
            "Penyusunan Program Semester (Promes) Pembelajaran Matematika",
            "Pengembangan Modul Ajar Kurikulum Merdeka / RPP Inovatif (Komponen Informasi Umum & Inti)",
            "Perumusan Tujuan Pembelajaran (ABCD Criteria: Audience, Behavior, Condition, Degree)",
            "Penentuan Model Pembelajaran (PBL, PjBL, Discovery, Inquiry) dalam Modul Ajar",
            "Perancangan Skenario Kegiatan Pembelajaran (Pendahuluan, Inti, Penutup)",
            "Integrasi Penguatan Profil Pelajar Pancasila (P5) dalam Modul Ajar Matematika",
            "Integrasi Pendekatan Deep Learning / STEAM / TPACK dalam Perencanaan Pembelajaran",
            "Penyusunan Lampiran Modul Ajar (LKPD, Bahan Bacaan, Glosarium, Asesmen)",
            "Review dan Peer-Assessment Perangkat Pembelajaran Matematika Sekolah",
            "Simulasi Presentasi Rencana Pelaksanaan Pembelajaran Matematika Inovatif"
        ]
    },
    {
        "kode_mk": "KIP2706", "nama_mk": "Telaah Kurikulum", "sks": 2, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas perkembangan kurikulum sekolah di Indonesia (Kurikulum 1975 hingga Kurikulum Merdeka), struktur kurikulum matematika SMP/SMA, perbandingan kurikulum internasional, dan evaluasi kurikulum.",
        "pustaka": ["Sukmadinata, N. S. (2012). Pengembangan Kurikulum: Teori dan Praktek. Bandung: Remaja Rosdakarya.", "Kemendikbudristek. (2022). Kerangka Kurikulum Merdeka."],
        "topics": [
            "Hakikat Kurikulum: Definisi, Dimensi, dan Organisasi Kurikulum",
            "Sejarah Perkembangan Kurikulum Sekolah di Indonesia (KPSP, KBK, KTSP, K13, Kurikulum Merdeka)",
            "Landasan Pengembangan Kurikulum: Filosofis, Psikologis, Sosiologis, dan IPTEK",
            "Struktur dan Karakteristik Kurikulum Merdeka pada Satuan Pendidikan Dasar & Menengah",
            "Telaah Dokumen Capaian Pembelajaran (CP) Matematika Fase D (SMP), Fase E & F (SMA)",
            "Perbandingan Kurikulum Matematika Indonesia dengan Kurikulum Internasional (Singapore, IB, Cambridge)",
            "Analisis Kerangka Asesmen PISA dan TIMSS dalam Kurikulum Matematika",
            "Pengembangan Kurikulum Operasional Satuan Pendidikan (KOSP)",
            "Telaah Buku Teks Siswa dan Buku Panduan Guru Matematika Kemendikbudristek",
            "Implementasi Projek Penguatan Profil Pelajar Pancasila (P5) Berbasis Matematika",
            "Problematika dan Tantangan Implementasi Kurikulum Merdeka di Sekolah",
            "Peran Guru dalam Pengembangan dan Adaptasi Kurikulum Tingkat Satuan Pendidikan",
            "Evaluasi Kurikulum: Model CIPP (Context, Input, Process, Product)",
            "Penyusunan Laporan Hasil Analisis Kritis Dokumentasi Kurikulum Sekolah"
        ]
    },
    {
        "kode_mk": "KIP2707", "nama_mk": "Praktik Mengajar (PLP)", "sks": 4, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini merupakan pengenalan lapangan persekolahan (PLP), magang praktikum mengajar terbimbing dan mandiri di sekolah mitra, pengelolaan kelas riil, dan penyusunan laporan PLP.",
        "pustaka": ["Panduan Pengenalan Lapangan Persekolahan (PLP) FKIP UNIROW Tuban.", "Nurdin, S. (2005). Model Pembelajaran Efektif."],
        "topics": [
            "Pembekalan dan Pelepasan Mahasiswa PLP di Sekolah Mitra",
            "Orientasi Budaya Sekolah, Tata Tertib, dan Struktur Organisasi Sekolah Mitra",
            "Observasi Kultur Sekolah, Pelaksanaan Pembelajaran Guru Pamong, dan Karakteristik Siswa",
            "Penyusunan Perangkat Pembelajaran Latihan Mengajar (Modul Ajar, LKPD, Media)",
            "Konsultasi dan Bimbingan Perangkat Pembelajaran dengan Dosen Pembimbing & Guru Pamong",
            "Praktik Mengajar Terbimbing Skenario 1 di Kelas Riil Sekolah Mitra",
            "Praktik Mengajar Terbimbing Skenario 2 di Kelas Riil Sekolah Mitra",
            "Praktik Mengajar Mandiri Skenario 1 di Kelas Riil Sekolah Mitra",
            "Praktik Mengajar Mandiri Skenario 2 di Kelas Riil Sekolah Mitra",
            "Pelaksanaan Asesmen Pembelajaran Siswa dan Pengolahan Nilai Hasil Belajar",
            "Pelibatan dalam Kegiatan Kesiswaan, Ekstrakurikuler, dan Administrasi Sekolah",
            "Pelaksanaan Evaluasi dan Refleksi Bersama Guru Pamong dan Dosen Pembimbing",
            "Penyusunan Draft Laporan Akhir Praktik Mengajar (PLP)",
            "Ujian Akhir PLP dan Penarikan Mahasiswa dari Sekolah Mitra"
        ]
    },
    {
        "kode_mk": "KIP2708", "nama_mk": "Pengembangan Strategi Pembelajaran", "sks": 2, "sem": 7, "rumpun": "Pembelajaran Matematika", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas tentang berbagai pendekatan & strategi pembelajaran matematika (PBL, PjBL, RME/PMRI, Inquiry, STEM/STEAM, Differentiated Learning) dan perancangannya.",
        "pustaka": ["Joyce, B., Weil, M., & Calhoun, E. (2015). Models of Teaching. Pearson.", "Gravemeijer, K. (1994). Developing Realistic Mathematics Education."],
        "topics": [
            "Hakikat Pendekatan, Strategi, Metode, dan Teknik Pembelajaran Matematika",
            "Model Pembelajaran Berbasis Masalah (Problem-Based Learning - PBL)",
            "Model Pembelajaran Berbasis Proyek (Project-Based Learning - PjBL)",
            "Model Pembelajaran Penemuan dan Penyelidikan (Discovery & Inquiry Learning)",
            "Pendekatan Pendidikan Matematika Realistik Indonesia (PMRI / RME)",
            "Pendekatan STEM / STEAM (Science, Technology, Engineering, Arts, Mathematics)",
            "Pembelajaran Berdiferensiasi (Differentiated Learning: Konten, Proses, Produk)",
            "Pendekatan Teaching at the Right Level (TaRL) dan Culturally Responsive Teaching (CRT)",
            "Model Pembelajaran Kooperatif (Jigsaw, STAD, TGT, Think-Pair-Share)",
            "Strategi Pembelajaran Berpikir Kritis dan Berpikir Kreatif Matematika",
            "Strategi Pembelajaran Metakognitif dan Self-Regulated Learning (SRL)",
            "Perancangan Sintaks Pembelajaran Inovatif dalam Skenario Modul Ajar",
            "Simulasi Penerapan Strategi Pembelajaran Inovatif Matematika",
            "Evaluasi Keberhasilan Strategi Pembelajaran terhadap Engagement Siswa"
        ]
    },

    # SEMESTER 8
    {
        "kode_mk": "PMA4122", "nama_mk": "Penelitian Pendidikan Matematika", "sks": 2, "sem": 8, "rumpun": "Penelitian Pendidikan", "prasyarat": "KIP2401",
        "deskripsi": "Mata kuliah ini membahas penyusunan draf proposal penelitian skripsi bidang pendidikan matematika secara utuh (Bab 1, Bab 2, Bab 3), uji seminar proposal, dan etika akademik.",
        "pustaka": ["Buku Pedoman Penulisan Skripsi FKIP UNIROW Tuban.", "Creswell, J. W. (2014). Educational Research. Pearson."],
        "topics": [
            "Penetapan Topik dan Pemilihan Masalah Penelitian Skripsi Pendidikan Matematika",
            "Penyusunan Bab I Pendahuluan: Latar Belakang Masalah dan Gap Penelitian",
            "Penyusunan Bab I Pendahuluan: Rumusan Masalah, Tujuan, dan Manfaat Penelitian",
            "Penyusunan Bab II Kajian Pustaka: Landasan Teori Variabel Penelitian Matematika",
            "Penyusunan Bab II Kajian Pustaka: Penelitian Relevan, Kerangka Berpikir, & Hipotesis",
            "Penyusunan Bab III Metodologi: Rancangan Penelitian, Tempat & Waktu Penelitian",
            "Penyusunan Bab III Metodologi: Populasi, Sampel, dan Teknik Sampling",
            "Penyusunan Bab III Metodologi: Operasionalisasi Variabel dan Kisi-Kisi Instrumen",
            "Penyusunan Bab III Metodologi: Teknik Pengumpulan Data dan Pengujian Instrumen",
            "Penyusunan Bab III Metodologi: Teknik Analisis Data (Statistik Parametrik / Non-Parametrik / Kualitatif)",
            "Penyusunan Daftar Pustaka Standar APA Style dan Lampiran Draft Proposal",
            "Pemeriksaan Draf Proposal Menggunakan Turnitin (Anti Plagiarisme)",
            "Simulasi Presentasi Seminar Proposal Penelitian Skripsi",
            "Finalisasi Dokumen Proposal Penelitian Skripsi Siap Ujian Sempro"
        ]
    },
    {
        "kode_mk": "PMA4124", "nama_mk": "Seminar Matematika", "sks": 2, "sem": 8, "rumpun": "Penelitian Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah ini membahas teknik penyusunan makalah ilmiah/proposal skripsi, etika seminar akademis, penggunaan media presentasi interaktif (PowerPoint/Canva), dan simulasi diskusi kritis.",
        "pustaka": ["Anholt, R. R. H. (2010). Dazzle 'em with Style: The Art of Oral Scientific Presentation. Academic Press.", "Pedoman Sempro FKIP UNIROW."],
        "topics": [
            "Hakikat Seminar Akademis dan Etika Forum Ilmiah",
            "Struktur dan Sistematika Naskah Makalah Seminar Matematika",
            "Teknik Mengabstraksi Hasil Penelitian dan Penentuan Kata Kunci",
            "Desain Slide Presentasi Ilmiah yang Komunikatif, Efektif, dan Estetis",
            "Penggunaan Software Presentasi Modern (Canva, Pitch, PowerPoint Advanced)",
            "Teknik Komunikasi Lisan, Olah Vokal, dan Body Language Pembicara Ilmiah",
            "Strategi Penyampaian Argumen Ilmiah dan Penguasaan Materi Presentasi",
            "Manajemen Waktu Presentasi dan Transisi Antar Slide",
            "Peran dan Tata Cara Penyanggah (Discussant) serta Moderator Seminar",
            "Teknik Menjawab Pertanyaan Kritis Dosen dan Peserta Seminar secara Akademis",
            "Simulasi Gelombang 1: Presentasi Makalah / Proposal Skripsi Mahasiswa",
            "Simulasi Gelombang 2: Presentasi Makalah / Proposal Skripsi Mahasiswa",
            "Evaluasi Komprehensif Performa Presentasi dan Umpan Balik Dosen Observer",
            "Penyusunan Berita Acara dan Revisi Naskah Hasil Seminar Matematika"
        ]
    },
    {
        "kode_mk": "PMA4125", "nama_mk": "Skripsi", "sks": 6, "sem": 8, "rumpun": "Penelitian Pendidikan", "prasyarat": "-",
        "deskripsi": "Mata kuliah tugas akhir yang mewajibkan mahasiswa melakukan penelitian mandiri di bidang pendidikan matematika, menyusun laporan skripsi secara ilmiah, dan mempertahankannya dalam Sidang Skripsi.",
        "pustaka": ["Buku Pedoman Penulisan Skripsi FKIP UNIROW Tuban Terbitan Terbaru.", "Jurnal-jurnal Terakreditasi SINTA & Scopus Pendidikan Matematika."],
        "topics": [
            "Bimbingan Intensif Pelaksanaan Penelitian dan Pengumpulan Data di Lapangan/Sekolah",
            "Pengolahan dan Analisis Data Penelitian Kuantitatif / Kualitatif / R&D",
            "Penyusunan Bab IV Hasil Penelitian: Penyajian Data dan Temuan Penelitian",
            "Penyusunan Bab IV Pembahasan: Analisis Kritis Temuan dengan Teori & Penelitian Relevan",
            "Penyusunan Bab V Penutup: Kesimpulan, Implikasi Kebijakan, dan Saran Penelitian",
            "Penyempurnaan Naskah Skripsi Lengkap (Halaman Depan, Abstrak, Bab I-V, Pustaka)",
            "Pemeriksaan Kelengkapan Instrumen, Data Mentah, dan Lampiran Skripsi",
            "Pemeriksaan Bebas Plagiarisme (Batas Maksimal Similarity Turnitin 25%)",
            "Penyusunan Artikel Ilmiah dari Hasil Skripsi untuk Publikasi Jurnal",
            "Persetujuan (Approval) Dosen Pembimbing I dan Pembimbing II untuk Ujian Munaqosyah",
            "Simulasi dan Persiapan Ujian Sidang Skripsi (Munaqosyah)",
            "Pelaksanaan Ujian Sidang Skripsi (Munaqosyah) di Hadapan Dewan Penguji",
            "Revisi Naskah Skripsi Sesuai Masukan Penguji dan Penandatanganan Pengesahan",
            "Pengunggahan (Upload) Karya Skripsi ke Repositori Perpustakaan UNIROW Tuban"
        ]
    }
]

def generate_course_json(c):
    data = {
        "identitas": {
            "fakultas": "Keguruan dan Ilmu Pendidikan",
            "prodi": "Pendidikan Matematika",
            "kode_dokumen": f"RPS/FKIP/PMAT/{c['kode_mk']}",
            "nama_mk": c["nama_mk"],
            "kode_mk": c["kode_mk"],
            "rumpun_mk": c["rumpun"],
            "bobot_sks": c["sks"],
            "semester": c["sem"],
            "tgl_penyusunan": "15 Januari 2025"
        },
        "otorisasi": {
            "pengembang": f"Tim Pengembang RPS {c['rumpun']}",
            "koordinator_rmk": "Rachmalia Vinda Kusuma, M.Pd.",
            "ujm": "Rachmalia Vinda Kusuma, M.Pd.",
            "kaprodi": "Puji Rahayu, M.Pd."
        },
        "validasi": {
            "kaprodi_nama": "Puji Rahayu, M.Pd.",
            "ujm_nama": "Rachmalia Vinda Kusuma, M.Pd.",
            "tgl_validasi": "15 Januari 2025"
        },
        "capaian_pembelajaran": {
            "cpl_prodi": [
                {"kode": "CPL-1 (S1)", "deskripsi": "Bertakwa kepada Tuhan Yang Maha Esa dan mampu menunjukkan sikap religius;"},
                {"kode": "CPL-2 (S2)", "deskripsi": "Menjunjung tinggi nilai kemanusiaan dalam menjalankan tugas berdasarkan agama, moral, dan etika;"},
                {"kode": "CPL-3 (S3)", "deskripsi": "Berkontribusi dalam peningkatan mutu kehidupan bermasyarakat, berbangsa, bernegara, dan kemajuan peradaban berdasarkan Pancasila;"},
                {"kode": "CPL-4 (S8)", "deskripsi": "Menginternalisasi nilai, norma, dan etika akademik;"},
                {"kode": "CPL-5 (KU1)", "deskripsi": "Mampu menerapkan pemikiran logis, kritis, sistematis, dan inovatif dalam konteks pengembangan atau implementasi ilmu pengetahuan dan teknologi yang memperhatikan dan menerapkan nilai humaniora;"},
                {"kode": "CPL-6 (P1)", "deskripsi": "Menguasai konsep pedagogi-didaktik matematika serta keilmuan matematika untuk merencanakan pembelajaran inovatif berbasis IPTEKS;"},
                {"kode": "CPL-7 (KK1)", "deskripsi": "Mampu mengaplikasikan konsep dan prinsip didaktik-pedagogis matematika serta keilmuan matematika untuk merencanakan pembelajaran inovatif dengan memanfaatkan IPTEKS."}
            ],
            "cpmk": [
                {"kode": "CPMK 1", "deskripsi": f"Mampu menganalisis dan menguasai konsep dasar {c['nama_mk']} secara ilmiah dan logis."},
                {"kode": "CPMK 2", "deskripsi": f"Mampu memecahkan masalah kontekstual dan mengaplikasikan prinsip {c['nama_mk']} dalam era pembelajaran abad 21."},
                {"kode": "CPMK 3", "deskripsi": f"Mampu merancang karya, analisis, atau penyelesaian berbasis {c['nama_mk']} berwawasan keindonesiaan dan IPTEKS."}
            ],
            "sub_cpmk": [],
            "matriks_korelasi": []
        },
        "deskripsi_mk": c["deskripsi"],
        "materi_pembelajaran": [f"{i+1}. {topik}" for i, topik in enumerate(c["topics"])],
        "pustaka": {
            "utama": c["pustaka"],
            "pendukung": ["Jurnal Pendidikan Matematika Terakreditasi SINTA", "Buku Referensi Pendidikan Matematika Abad 21"]
        },
        "dosen_pengampu": [f"Tim Dosen Pengampu {c['nama_mk']}"],
        "prasyarat": c["prasyarat"],
        "mingguan": []
    }

    topics_list = c["topics"]
    pustaka_list = c["pustaka"]
    nama = c["nama_mk"]
    sks = c["sks"]

    for m in range(1, 15):
        sub_code = f"Sub-CPMK {m}"
        topik_materi = topics_list[m-1]
        sub_desc = f"Mampu menganalisis dan mengaplikasikan konsep {topik_materi}."
        
        # Clean without bracket taxonomy
        data["capaian_pembelajaran"]["sub_cpmk"].append({
            "kode": sub_code, "taksonomi": "", "deskripsi": sub_desc
        })
        
        cpl_checks = {"CPL-1 (S1)": "✓" if m in [1,9] else "", "CPL-2 (S2)": "✓" if m in [4,10] else "", "CPL-3 (S3)": "✓" if m in [1,9,14] else "", "CPL-4 (S8)": "✓" if m in [4,13] else "", "CPL-5 (KU1)": "✓", "CPL-6 (P1)": "✓" if m % 2 == 0 else "", "CPL-7 (KK1)": "✓" if m % 2 != 0 else ""}
        data["capaian_pembelajaran"]["matriks_korelasi"].append({
            "label": sub_code, "cpls": cpl_checks, "bobot": "3%" if m in [1,5,6,10,12,13] else "4%"
        })

    all_v = {"CPL-1 (S1)": "✓", "CPL-2 (S2)": "✓", "CPL-3 (S3)": "✓", "CPL-4 (S8)": "✓", "CPL-5 (KU1)": "✓", "CPL-6 (P1)": "✓", "CPL-7 (KK1)": "✓"}
    all_100 = {"CPL-1 (S1)": "100%", "CPL-2 (S2)": "100%", "CPL-3 (S3)": "100%", "CPL-4 (S8)": "100%", "CPL-5 (KU1)": "100%", "CPL-6 (P1)": "100%", "CPL-7 (KK1)": "100%"}
    
    data["capaian_pembelajaran"]["matriks_korelasi"].insert(7, {"label": "Evaluasi Tengah Semester (UTS)", "cpls": all_v, "bobot": "25%", "is_uts": True})
    data["capaian_pembelajaran"]["matriks_korelasi"].append({"label": "Evaluasi Akhir Semester (UAS)", "cpls": all_v, "bobot": "25%", "is_uas": True})
    data["capaian_pembelajaran"]["matriks_korelasi"].append({"label": "Total", "cpls": all_100, "bobot": "100%", "is_total": True})

    sub_idx = 0
    for w in range(1, 17):
        if w == 8:
            data["mingguan"].append({
                "minggu": 8, "is_uts": True,
                "indikator": f"Evaluasi penguasaan materi perkuliahan minggu 1 s.d. 7 ({topics_list[0]} s.d. {topics_list[6]})",
                "bobot": 25
            })
        elif w == 16:
            data["mingguan"].append({
                "minggu": 16, "is_uas": True,
                "indikator": f"Evaluasi komprehensif penguasaan capaian pembelajaran mata kuliah {nama} ({topics_list[0]} s.d. {topics_list[13]})",
                "bobot": 25
            })
        else:
            sub_idx += 1
            topik_materi = topics_list[sub_idx - 1]
            w_sub = data["capaian_pembelajaran"]["sub_cpmk"][sub_idx - 1]
            data["mingguan"].append({
                "minggu": w,
                "sub_cpmk": f"{w_sub['kode']}: {w_sub['deskripsi']}",
                "indikator": f"1. Ketepatan menjelaskan konsep {topik_materi}.\n2. Kejelasan memecahkan masalah & aplikasi {topik_materi}.",
                "teknik_kriteria": "Kriteria: Rubrik Analisis & Penskoran\nTeknik Non-Tes: Tugas & Diskusi",
                "luring": f"Kuliah Tatap Muka (PB: {sks}x50')\nDiskusi & Problem Solving (PT: {sks}x60', KM: {sks}x60')",
                "daring": "-",
                "materi_pustaka": f"{topik_materi} ({pustaka_list[0]})",
                "bobot": 3 if sub_idx in [1,5,6,10,12,13] else 4
            })

    return data

def main():
    print(f"=== BATCH GENERATOR RPS LENGKAP {len(COURSES_68)} MATA KULIAH KURIKULUM PENDIDIKAN MATEMATIKA UNIROW TUBAN ===")
    count = 0
    for idx, c in enumerate(COURSES_68, start=1):
        mk_clean = c["nama_mk"].replace(" ", "_").replace("/", "_").replace("(", "").replace(")", "")
        out_filename = f"{idx:02d}_RPS_{c['kode_mk']}_{mk_clean}_OBE.docx"
        out_path = os.path.join(OUTPUT_DIR, out_filename)
        
        data = generate_course_json(c)
        try:
            generate_rps_docx(data, output_path=out_path)
            count += 1
            print(f"[{idx}/{len(COURSES_68)}] SUCCESS: {out_filename}")
        except Exception as e:
            print(f"[{idx}/{len(COURSES_68)}] ERROR {out_filename}: {e}")

    print(f"\n[DONE] Selesai membangkitkan {count} dokumen RPS Lengkap Bersih di folder:\n{OUTPUT_DIR}")

if __name__ == "__main__":
    main()
