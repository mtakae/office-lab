
'use strict';
/* Studi kasus seluruhnya fiktif, tidak berisi NIM/nama/rekening riil. */
const CAMPUS_CASES=[
 {
  "id": "keuangan",
  "name": "Keuangan",
  "emoji": "Rp",
  "theme": "#0d956b",
  "tagline": "Pengendalian pagu dan realisasi",
  "unit": "Subbagian Keuangan",
  "kind": "Rupiah",
  "description": "Penyusunan rekap penggunaan anggaran, daftar selisih dan catatan telaah realisasi.",
  "scenario": "Rapat monitoring bulanan meminta daftar enam kegiatan, total pagu, total realisasi, identifikasi kegiatan melebihi anggaran, serta memo tindak lanjut.",
  "kpi": "Pagu, realisasi, selisih dan penanganan deviasi",
  "action": "Memverifikasi realisasi yang melampaui pagu dan melengkapi bukti sumber.",
  "files": [
   "Dataset_keuangan.csv",
   "Dokumen_keuangan.html",
   "Paparan_keuangan.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 25500000,
   "actual": 23100000,
   "gap": 2400000
  }
 },
 {
  "id": "kemahasiswaan",
  "name": "Kemahasiswaan",
  "emoji": "🎓",
  "theme": "#326ed0",
  "tagline": "Pemantauan permohonan dan beasiswa",
  "unit": "Layanan Kemahasiswaan",
  "kind": "Layanan",
  "description": "Rekap pengajuan beasiswa, permohonan mahasiswa, status berkas dan waktu tindak lanjut.",
  "scenario": "Kepala unit meminta ringkasan jumlah permohonan layanan mahasiswa per jenis layanan dan ringkasan backlog, tanpa identitas mahasiswa riil.",
  "kpi": "Jumlah layanan masuk, selesai, dan selisih terhadap target",
  "action": "Menyusun prioritas verifikasi berkas berdasarkan data dan antrean.",
  "files": [
   "Dataset_kemahasiswaan.csv",
   "Dokumen_kemahasiswaan.html",
   "Paparan_kemahasiswaan.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 710,
   "actual": 658,
   "gap": 52
  }
 },
 {
  "id": "perencanaan",
  "name": "Perencanaan",
  "emoji": "◎",
  "theme": "#a66b23",
  "tagline": "Monitoring target dan capaian program",
  "unit": "Bagian Perencanaan",
  "kind": "Program",
  "description": "Penyusunan matriks program kerja, capaian indikator, timeline, serta opsi tindak lanjut.",
  "scenario": "Unit menggabungkan indikator capaian program dari beberapa bagian menjadi dashboard dan paparan rapat evaluasi.",
  "kpi": "Target program, capaian, variansi dan prioritas perbaikan",
  "action": "Mengidentifikasi hambatan serta mengonfirmasi pemilik tugas dan tenggat.",
  "files": [
   "Dataset_perencanaan.csv",
   "Dokumen_perencanaan.html",
   "Paparan_perencanaan.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 58,
   "actual": 50,
   "gap": 8
  }
 },
 {
  "id": "arsip",
  "name": "Persuratan & Arsip",
  "emoji": "▤",
  "theme": "#9264be",
  "tagline": "Register surat dan metadata arsip",
  "unit": "Unit Kearsipan",
  "kind": "Dokumen",
  "description": "Register surat masuk/keluar, indeks arsip, metadata, klasifikasi dan akses sesuai pedoman institusi.",
  "scenario": "Petugas memerlukan rekap volume dokumen, daftar arsip yang masih perlu verifikasi metadata, dan memo pengendalian klasifikasi.",
  "kpi": "Volume dokumen, verifikasi metadata, dan ketepatan register",
  "action": "Mengonfirmasi kode klasifikasi dan akses kepada pengelola arsip berwenang.",
  "files": [
   "Dataset_arsip.csv",
   "Dokumen_arsip.html",
   "Paparan_arsip.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 481,
   "actual": 441,
   "gap": 40
  }
 },
 {
  "id": "kepegawaian",
  "name": "Kepegawaian",
  "emoji": "👥",
  "theme": "#c36a62",
  "tagline": "Rekap pelatihan dan layanan pegawai",
  "unit": "Bagian Kepegawaian",
  "kind": "Peserta",
  "description": "Rekap kehadiran, kompetensi pelatihan, jadwal pembinaan, dan evaluasi kegiatan pegawai.",
  "scenario": "Bagian kepegawaian perlu merekap pelatihan internal serta membuat undangan dan paparan hasil program pengembangan kompetensi.",
  "kpi": "Target peserta, hadir, rasio kehadiran dan tindak lanjut",
  "action": "Menetapkan topik pendampingan berikut berdasarkan hasil latihan dan umpan balik.",
  "files": [
   "Dataset_kepegawaian.csv",
   "Dokumen_kepegawaian.html",
   "Paparan_kepegawaian.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 215,
   "actual": 195,
   "gap": 20
  }
 },
 {
  "id": "akademik",
  "name": "Layanan Akademik",
  "emoji": "📚",
  "theme": "#3382a1",
  "tagline": "Administrasi surat dan registrasi akademik",
  "unit": "Pelayanan Akademik",
  "kind": "Permohonan",
  "description": "Register surat, verifikasi permohonan, rekap layanan akademik, dan penyusunan laporan bulanan.",
  "scenario": "Layanan akademik melakukan evaluasi permohonan dan menyiapkan surat keterangan latihan serta ringkasan capaian.",
  "kpi": "Permohonan masuk, selesai, backlog dan waktu proses",
  "action": "Memperjelas panduan kelengkapan berkas tanpa mengungkap identitas mahasiswa.",
  "files": [
   "Dataset_akademik.csv",
   "Dokumen_akademik.html",
   "Paparan_akademik.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 515,
   "actual": 481,
   "gap": 34
  }
 },
 {
  "id": "aset",
  "name": "Pengadaan & Aset",
  "emoji": "▣",
  "theme": "#4d79b1",
  "tagline": "Inventaris dan penggunaan sarana",
  "unit": "Pengadaan dan Aset",
  "kind": "Barang",
  "description": "Register inventaris, berita acara pemeriksaan, rekap kebutuhan barang, dan paparan kondisi aset.",
  "scenario": "Subbagian aset memeriksa stok inventaris dan mempresentasikan kebutuhan pemeliharaan berdasarkan data terverifikasi.",
  "kpi": "Aset tercatat, penyimpangan, dan kebutuhan tindak lanjut",
  "action": "Menyiapkan daftar barang yang perlu diverifikasi bersama pengelola aset.",
  "files": [
   "Dataset_aset.csv",
   "Dokumen_aset.html",
   "Paparan_aset.html"
  ],
  "tasks": [
   {
    "id": "01",
    "text": "Unduh data kasus CSV dan identifikasi header, kategori, serta satuannya."
   },
   {
    "id": "02",
    "text": "Buat rekap Excel: selisih, jumlah, indikator kategori, dan pemeriksaan silang."
   },
   {
    "id": "03",
    "text": "Catat sedikitnya satu anomali serta kebutuhan validasi data sumber."
   },
   {
    "id": "04",
    "text": "Muat template Word, lengkapi temuan dan tabel tindak lanjut dengan data simulasi."
   },
   {
    "id": "05",
    "text": "Periksa bahasa, penomoran, sumber, versi, dan batas akses draf Word."
   },
   {
    "id": "06",
    "text": "Muat template PowerPoint, isi angka KPI dan satuan dari Excel."
   },
   {
    "id": "07",
    "text": "Presentasikan minimal empat slide dengan catatan pembicara."
   },
   {
    "id": "08",
    "text": "Lakukan audit akhir: konsistensi angka Excel–Word–PowerPoint dan simpan bukti."
   }
  ],
  "rubric": [
   {
    "name": "Ketepatan data",
    "weight": 35,
    "desc": "Hasil dan rumus dapat diperiksa dari sumber sintetis."
   },
   {
    "name": "Kualitas dokumen",
    "weight": 25,
    "desc": "Surat/laporan lengkap, konsisten, dan tidak berisi data sensitif asli."
   },
   {
    "name": "Kejelasan paparan",
    "weight": 25,
    "desc": "Judul temuan, indikator, sumber, dan tindak lanjut dapat dijelaskan."
   },
   {
    "name": "Kontrol & kolaborasi",
    "weight": 15,
    "desc": "Nama file, versi, pemeriksaan, dan dokumentasi keluaran sesuai."
   }
  ],
  "metricSummary": {
   "target": 290,
   "actual": 268,
   "gap": 22
  }
 }
];
const EXCEL_CASES={
 "keuangan": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Pagu (Rp)",
   "D1": "Realisasi (Rp)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "ATK Layanan",
   "B2": "Operasional",
   "C2": "5000000",
   "D2": "4200000",
   "E2": "=C2-D2",
   "A3": "Rapat Koordinasi",
   "B3": "Operasional",
   "C3": "3000000",
   "D3": "2750000",
   "E3": "=C3-D3",
   "A4": "Pemeliharaan Sistem",
   "B4": "Digital",
   "C4": "8000000",
   "D4": "6800000",
   "E4": "=C4-D4",
   "A5": "Konsumsi Pelatihan",
   "B5": "Operasional",
   "C5": "2000000",
   "D5": "1800000",
   "E5": "=C5-D5",
   "A6": "Lisensi Aplikasi",
   "B6": "Digital",
   "C6": "6000000",
   "D6": "6200000",
   "E6": "=C6-D6",
   "A7": "Dokumentasi",
   "B7": "Operasional",
   "C7": "1500000",
   "D7": "1350000",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Keuangan",
  "metric": "Rupiah",
  "criterion": "Operasional",
  "lookup": "Rapat Koordinasi",
  "c": "Pagu (Rp)",
  "d": "Realisasi (Rp)",
  "description": "Rapat monitoring bulanan meminta daftar enam kegiatan, total pagu, total realisasi, identifikasi kegiatan melebihi anggaran, serta memo tindak lanjut."
 },
 "kemahasiswaan": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (berkas)",
   "D1": "Selesai (berkas)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Pengajuan Beasiswa",
   "B2": "Layanan",
   "C2": "120",
   "D2": "108",
   "E2": "=C2-D2",
   "A3": "Surat Aktif Kuliah",
   "B3": "Layanan",
   "C3": "180",
   "D3": "169",
   "E3": "=C3-D3",
   "A4": "Verifikasi Berkas",
   "B4": "Verifikasi",
   "C4": "150",
   "D4": "144",
   "E4": "=C4-D4",
   "A5": "Aduan Mahasiswa",
   "B5": "Layanan",
   "C5": "90",
   "D5": "81",
   "E5": "=C5-D5",
   "A6": "Validasi Persyaratan",
   "B6": "Verifikasi",
   "C6": "100",
   "D6": "92",
   "E6": "=C6-D6",
   "A7": "Sosialisasi Beasiswa",
   "B7": "Edukasi",
   "C7": "70",
   "D7": "64",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Kemahasiswaan",
  "metric": "Layanan",
  "criterion": "Layanan",
  "lookup": "Surat Aktif Kuliah",
  "c": "Target (berkas)",
  "d": "Selesai (berkas)",
  "description": "Kepala unit meminta ringkasan jumlah permohonan layanan mahasiswa per jenis layanan dan ringkasan backlog, tanpa identitas mahasiswa riil."
 },
 "perencanaan": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (unit)",
   "D1": "Capaian (unit)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Rapat Koordinasi",
   "B2": "Prioritas",
   "C2": "12",
   "D2": "10",
   "E2": "=C2-D2",
   "A3": "Penyusunan Rencana Kerja",
   "B3": "Prioritas",
   "C3": "8",
   "D3": "7",
   "E3": "=C3-D3",
   "A4": "Monitoring Unit",
   "B4": "Rutin",
   "C4": "18",
   "D4": "15",
   "E4": "=C4-D4",
   "A5": "Evaluasi Indikator",
   "B5": "Prioritas",
   "C5": "10",
   "D5": "9",
   "E5": "=C5-D5",
   "A6": "Sosialisasi Program",
   "B6": "Rutin",
   "C6": "6",
   "D6": "6",
   "E6": "=C6-D6",
   "A7": "Laporan Tahunan",
   "B7": "Prioritas",
   "C7": "4",
   "D7": "3",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Perencanaan",
  "metric": "Program",
  "criterion": "Prioritas",
  "lookup": "Rapat Koordinasi",
  "c": "Target (unit)",
  "d": "Capaian (unit)",
  "description": "Unit menggabungkan indikator capaian program dari beberapa bagian menjadi dashboard dan paparan rapat evaluasi."
 },
 "arsip": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (berkas)",
   "D1": "Diproses (berkas)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Penerimaan Surat",
   "B2": "Surat Masuk",
   "C2": "85",
   "D2": "80",
   "E2": "=C2-D2",
   "A3": "Registrasi Disposisi",
   "B3": "Surat Masuk",
   "C3": "72",
   "D3": "69",
   "E3": "=C3-D3",
   "A4": "Verifikasi Metadata",
   "B4": "Surat Masuk",
   "C4": "90",
   "D4": "76",
   "E4": "=C4-D4",
   "A5": "Penerbitan Surat",
   "B5": "Surat Keluar",
   "C5": "64",
   "D5": "61",
   "E5": "=C5-D5",
   "A6": "Pengiriman Surat",
   "B6": "Surat Keluar",
   "C6": "60",
   "D6": "57",
   "E6": "=C6-D6",
   "A7": "Penataan Berkas",
   "B7": "Arsip",
   "C7": "110",
   "D7": "98",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Persuratan & Arsip",
  "metric": "Dokumen",
  "criterion": "Surat Masuk",
  "lookup": "Verifikasi Metadata",
  "c": "Target (berkas)",
  "d": "Diproses (berkas)",
  "description": "Petugas memerlukan rekap volume dokumen, daftar arsip yang masih perlu verifikasi metadata, dan memo pengendalian klasifikasi."
 },
 "kepegawaian": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (orang)",
   "D1": "Realisasi (orang)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Pelatihan Office",
   "B2": "Pelatihan",
   "C2": "40",
   "D2": "37",
   "E2": "=C2-D2",
   "A3": "Pelatihan Arsip",
   "B3": "Pelatihan",
   "C3": "32",
   "D3": "29",
   "E3": "=C3-D3",
   "A4": "Sosialisasi SOP",
   "B4": "Sosialisasi",
   "C4": "50",
   "D4": "45",
   "E4": "=C4-D4",
   "A5": "Workshop Layanan",
   "B5": "Pelatihan",
   "C5": "38",
   "D5": "35",
   "E5": "=C5-D5",
   "A6": "Orientasi Staf",
   "B6": "Sosialisasi",
   "C6": "20",
   "D6": "19",
   "E6": "=C6-D6",
   "A7": "Evaluasi Kompetensi",
   "B7": "Evaluasi",
   "C7": "35",
   "D7": "30",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Kepegawaian",
  "metric": "Peserta",
  "criterion": "Pelatihan",
  "lookup": "Pelatihan Office",
  "c": "Target (orang)",
  "d": "Realisasi (orang)",
  "description": "Bagian kepegawaian perlu merekap pelatihan internal serta membuat undangan dan paparan hasil program pengembangan kompetensi."
 },
 "akademik": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (berkas)",
   "D1": "Selesai (berkas)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Surat Keterangan",
   "B2": "Surat",
   "C2": "110",
   "D2": "102",
   "E2": "=C2-D2",
   "A3": "Surat Aktif",
   "B3": "Surat",
   "C3": "95",
   "D3": "90",
   "E3": "=C3-D3",
   "A4": "Perbaikan Biodata",
   "B4": "Registrasi",
   "C4": "80",
   "D4": "73",
   "E4": "=C4-D4",
   "A5": "Surat Cuti",
   "B5": "Surat",
   "C5": "40",
   "D5": "35",
   "E5": "=C5-D5",
   "A6": "Registrasi Layanan",
   "B6": "Registrasi",
   "C6": "120",
   "D6": "116",
   "E6": "=C6-D6",
   "A7": "Legalitas Dokumen",
   "B7": "Verifikasi",
   "C7": "70",
   "D7": "65",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Layanan Akademik",
  "metric": "Permohonan",
  "criterion": "Surat",
  "lookup": "Surat Keterangan",
  "c": "Target (berkas)",
  "d": "Selesai (berkas)",
  "description": "Layanan akademik melakukan evaluasi permohonan dan menyiapkan surat keterangan latihan serta ringkasan capaian."
 },
 "aset": {
  "cells": {
   "A1": "Kegiatan / Layanan",
   "B1": "Kategori",
   "C1": "Target (unit)",
   "D1": "Tercatat (unit)",
   "E1": "Selisih",
   "F1": "Status",
   "G1": "Latihan",
   "A2": "Pemeriksaan Ruang",
   "B2": "Inventaris",
   "C2": "50",
   "D2": "47",
   "E2": "=C2-D2",
   "A3": "Rekap Peralatan",
   "B3": "Inventaris",
   "C3": "65",
   "D3": "62",
   "E3": "=C3-D3",
   "A4": "Inventaris Komputer",
   "B4": "Inventaris",
   "C4": "40",
   "D4": "39",
   "E4": "=C4-D4",
   "A5": "Pengajuan Barang",
   "B5": "Pengadaan",
   "C5": "30",
   "D5": "26",
   "E5": "=C5-D5",
   "A6": "Pemeliharaan Printer",
   "B6": "Pemeliharaan",
   "C6": "25",
   "D6": "19",
   "E6": "=C6-D6",
   "A7": "Penataan Gudang",
   "B7": "Inventaris",
   "C7": "80",
   "D7": "75",
   "E7": "=C7-D7",
   "A8": "TOTAL",
   "A9": "INDIKATOR"
  },
  "title": "Pengadaan & Aset",
  "metric": "Barang",
  "criterion": "Inventaris",
  "lookup": "Pemeriksaan Ruang",
  "c": "Target (unit)",
  "d": "Tercatat (unit)",
  "description": "Subbagian aset memeriksa stok inventaris dan mempresentasikan kebutuhan pemeliharaan berdasarkan data terverifikasi."
 }
};
const WORD_TEMPLATES={
 "keuangan": {
  "html": "<h1>MEMO PENELAAHAN REALISASI ANGGARAN</h1><p><b>Unit:</b> Subbagian Keuangan (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Berdasarkan rekap fiktif triwulan berjalan, unit meninjau realisasi belanja dan perbedaan terhadap pagu untuk enam kegiatan.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Pagu, realisasi, selisih dan penanganan deviasi</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Memverifikasi realisasi yang melampaui pagu dan melengkapi bukti sumber.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "MEMO PENELAAHAN REALISASI ANGGARAN",
  "desc": "Berdasarkan rekap fiktif triwulan berjalan, unit meninjau realisasi belanja dan perbedaan terhadap pagu untuk enam kegiatan."
 },
 "kemahasiswaan": {
  "html": "<h1>LAPORAN RINGKAS LAYANAN MAHASISWA</h1><p><b>Unit:</b> Layanan Kemahasiswaan (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Unit mendokumentasikan jumlah permohonan layanan dengan kode fiktif dan status tindak lanjut yang dapat diverifikasi.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Jumlah layanan masuk, selesai, dan selisih terhadap target</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Menyusun prioritas verifikasi berkas berdasarkan data dan antrean.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "LAPORAN RINGKAS LAYANAN MAHASISWA",
  "desc": "Unit mendokumentasikan jumlah permohonan layanan dengan kode fiktif dan status tindak lanjut yang dapat diverifikasi."
 },
 "perencanaan": {
  "html": "<h1>RINGKASAN MONITORING PROGRAM KERJA</h1><p><b>Unit:</b> Bagian Perencanaan (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Bagian perencanaan merangkum capaian indikator program tanpa mengubah definisi target dari unit sumber.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Target program, capaian, variansi dan prioritas perbaikan</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Mengidentifikasi hambatan serta mengonfirmasi pemilik tugas dan tenggat.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "RINGKASAN MONITORING PROGRAM KERJA",
  "desc": "Bagian perencanaan merangkum capaian indikator program tanpa mengubah definisi target dari unit sumber."
 },
 "arsip": {
  "html": "<h1>MEMO PENATAAN REGISTER DAN ARSIP</h1><p><b>Unit:</b> Unit Kearsipan (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Unit menguji konsistensi metadata dokumen simulasi dan melakukan pengelompokan sesuai klasifikasi yang disahkan institusi.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Volume dokumen, verifikasi metadata, dan ketepatan register</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Mengonfirmasi kode klasifikasi dan akses kepada pengelola arsip berwenang.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "MEMO PENATAAN REGISTER DAN ARSIP",
  "desc": "Unit menguji konsistensi metadata dokumen simulasi dan melakukan pengelompokan sesuai klasifikasi yang disahkan institusi."
 },
 "kepegawaian": {
  "html": "<h1>LAPORAN EVALUASI PELATIHAN PEGAWAI</h1><p><b>Unit:</b> Bagian Kepegawaian (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Data sintetik peserta dan kehadiran digunakan untuk menghitung capaian kegiatan tanpa memuat data personal staf.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Target peserta, hadir, rasio kehadiran dan tindak lanjut</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Menetapkan topik pendampingan berikut berdasarkan hasil latihan dan umpan balik.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "LAPORAN EVALUASI PELATIHAN PEGAWAI",
  "desc": "Data sintetik peserta dan kehadiran digunakan untuk menghitung capaian kegiatan tanpa memuat data personal staf."
 },
 "akademik": {
  "html": "<h1>REKAP PELAYANAN ADMINISTRASI AKADEMIK</h1><p><b>Unit:</b> Pelayanan Akademik (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Petugas merekap proses permohonan sintetis dan melaporkan kendala berkas yang memerlukan tindak lanjut.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Permohonan masuk, selesai, backlog dan waktu proses</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Memperjelas panduan kelengkapan berkas tanpa mengungkap identitas mahasiswa.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "REKAP PELAYANAN ADMINISTRASI AKADEMIK",
  "desc": "Petugas merekap proses permohonan sintetis dan melaporkan kendala berkas yang memerlukan tindak lanjut."
 },
 "aset": {
  "html": "<h1>BERITA ACARA PEMERIKSAAN INVENTARIS</h1><p><b>Unit:</b> Pengadaan dan Aset (simulasi) · <b>Periode:</b> November 2026</p><p><b>Jenis:</b> Draf pelatihan, bukan dokumen resmi.</p><h2>Latar Belakang</h2><p>Pemeriksaan aset latihan membedakan barang tercatat, kondisi, dan kebutuhan tindak lanjut tanpa menyatakan hasil pemeriksaan nyata.</p><h2>Ringkasan Data</h2><table><tbody><tr><th>Indikator</th><th>Hasil / Catatan</th></tr><tr><td>Aset tercatat, penyimpangan, dan kebutuhan tindak lanjut</td><td>[Isi dari hasil Excel]</td></tr><tr><td>Tanggal pemeriksaan</td><td>[Tanggal latihan]</td></tr></tbody></table><h2>Temuan dan Tindak Lanjut</h2><ul><li>Menyiapkan daftar barang yang perlu diverifikasi bersama pengelola aset.</li><li>Memeriksa sumber data dan mencatat hasil pengecekan.</li></ul><p><b>Disusun oleh:</b> Peserta pelatihan · <b>Status:</b> Draf untuk latihan.</p>",
  "title": "BERITA ACARA PEMERIKSAAN INVENTARIS",
  "desc": "Pemeriksaan aset latihan membedakan barang tercatat, kondisi, dan kebutuhan tindak lanjut tanpa menyatakan hasil pemeriksaan nyata."
 }
};
const PPT_TEMPLATES={
 "keuangan": {
  "slides": [
   {
    "title": "Paparan Keuangan",
    "body": "Subbagian Keuangan · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Penyusunan rekap penggunaan anggaran, daftar selisih dan catatan telaah realisasi.\nRapat monitoring bulanan meminta daftar enam kegiatan, total pagu, total realisasi, identifikasi kegiatan melebihi anggaran, serta memo tindak lanjut.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Pagu, realisasi, selisih dan penanganan deviasi\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Memverifikasi realisasi yang melampaui pagu dan melengkapi bukti sumber.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Keuangan"
 },
 "kemahasiswaan": {
  "slides": [
   {
    "title": "Paparan Kemahasiswaan",
    "body": "Layanan Kemahasiswaan · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Rekap pengajuan beasiswa, permohonan mahasiswa, status berkas dan waktu tindak lanjut.\nKepala unit meminta ringkasan jumlah permohonan layanan mahasiswa per jenis layanan dan ringkasan backlog, tanpa identitas mahasiswa riil.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Jumlah layanan masuk, selesai, dan selisih terhadap target\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Menyusun prioritas verifikasi berkas berdasarkan data dan antrean.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Kemahasiswaan"
 },
 "perencanaan": {
  "slides": [
   {
    "title": "Paparan Perencanaan",
    "body": "Bagian Perencanaan · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Penyusunan matriks program kerja, capaian indikator, timeline, serta opsi tindak lanjut.\nUnit menggabungkan indikator capaian program dari beberapa bagian menjadi dashboard dan paparan rapat evaluasi.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Target program, capaian, variansi dan prioritas perbaikan\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Mengidentifikasi hambatan serta mengonfirmasi pemilik tugas dan tenggat.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Perencanaan"
 },
 "arsip": {
  "slides": [
   {
    "title": "Paparan Persuratan & Arsip",
    "body": "Unit Kearsipan · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Register surat masuk/keluar, indeks arsip, metadata, klasifikasi dan akses sesuai pedoman institusi.\nPetugas memerlukan rekap volume dokumen, daftar arsip yang masih perlu verifikasi metadata, dan memo pengendalian klasifikasi.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Volume dokumen, verifikasi metadata, dan ketepatan register\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Mengonfirmasi kode klasifikasi dan akses kepada pengelola arsip berwenang.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Persuratan & Arsip"
 },
 "kepegawaian": {
  "slides": [
   {
    "title": "Paparan Kepegawaian",
    "body": "Bagian Kepegawaian · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Rekap kehadiran, kompetensi pelatihan, jadwal pembinaan, dan evaluasi kegiatan pegawai.\nBagian kepegawaian perlu merekap pelatihan internal serta membuat undangan dan paparan hasil program pengembangan kompetensi.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Target peserta, hadir, rasio kehadiran dan tindak lanjut\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Menetapkan topik pendampingan berikut berdasarkan hasil latihan dan umpan balik.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Kepegawaian"
 },
 "akademik": {
  "slides": [
   {
    "title": "Paparan Layanan Akademik",
    "body": "Pelayanan Akademik · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Register surat, verifikasi permohonan, rekap layanan akademik, dan penyusunan laporan bulanan.\nLayanan akademik melakukan evaluasi permohonan dan menyiapkan surat keterangan latihan serta ringkasan capaian.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Permohonan masuk, selesai, backlog dan waktu proses\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Memperjelas panduan kelengkapan berkas tanpa mengungkap identitas mahasiswa.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Layanan Akademik"
 },
 "aset": {
  "slides": [
   {
    "title": "Paparan Pengadaan & Aset",
    "body": "Pengadaan dan Aset · Pelatihan Administrasi Kampus\nDATA SINTETIS / LATIHAN",
    "layout": "cover",
    "notes": "Jelaskan tujuan rapat dan bahwa seluruh data contoh adalah fiktif."
   },
   {
    "title": "Konteks dan Tujuan",
    "body": "Register inventaris, berita acara pemeriksaan, rekap kebutuhan barang, dan paparan kondisi aset.\nSubbagian aset memeriksa stok inventaris dan mempresentasikan kebutuhan pemeliharaan berdasarkan data terverifikasi.",
    "layout": "content",
    "notes": "Terangkan pertanyaan evaluasi yang ingin dijawab dengan data."
   },
   {
    "title": "Indikator Kinerja Utama",
    "body": "Aset tercatat, penyimpangan, dan kebutuhan tindak lanjut\n[Masukkan angka setelah memeriksa dataset Excel]",
    "layout": "split",
    "notes": "Pisahkan definisi indikator, sumber, periode, serta catatan verifikasi."
   },
   {
    "title": "Usulan Tindak Lanjut",
    "body": "Menyiapkan daftar barang yang perlu diverifikasi bersama pengelola aset.\nPIC: [unit penanggung jawab]\nTenggat: [sesuai pembahasan]",
    "layout": "content",
    "notes": "Semua tindakan dan tenggat dalam template ini harus ditinjau pihak berwenang."
   }
  ],
  "title": "Paparan Pengadaan & Aset"
 }
};
const CASE_BY_ID=Object.fromEntries(CAMPUS_CASES.map(x=>[x.id,x]));
