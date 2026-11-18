
'use strict';
/* OfficeLab 2.1 — 75 modul administrasi kampus, latihan dengan data fiktif. */
const LEVELS=[{"id": "Basic", "name": "Basic", "label": "01", "caption": "Fondasi administrasi digital", "description": "Menguasai dokumen, data, dan navigasi dasar.", "color": "#64748b"}, {"id": "Beginner", "name": "Beginner", "label": "02", "caption": "Tugas rutin unit kampus", "description": "Menyelesaikan pekerjaan rutin secara rapi dan terverifikasi.", "color": "#159568"}, {"id": "Intermediate", "name": "Intermediate", "label": "03", "caption": "Analisis dan dokumen terstruktur", "description": "Mengolah data dan menghasilkan keluaran lintas fitur.", "color": "#3976c8"}, {"id": "Advanced", "name": "Advanced", "label": "04", "caption": "Kontrol, kolaborasi, dan otomasi", "description": "Mengelola skenario kompleks dan pemeriksaan mutu.", "color": "#c18836"}, {"id": "Expert", "name": "Expert", "label": "05", "caption": "Solusi dan proyek akhir unit", "description": "Merancang keluaran terintegrasi untuk pengambilan keputusan.", "color": "#885ac5"}];
const V1_INDEX_TO_V2={excel:[],word:[],powerpoint:[]};
const CURRICULUM={
 "excel": [
  {
   "title": "Mengenal workbook, sheet, dan struktur data kampus",
   "objective": "Mengidentifikasi elemen worksheet dan membangun satu tabel data layanan yang konsisten.",
   "points": [
    "Workbook adalah berkas Excel; worksheet adalah lembar di dalamnya.",
    "Header menjelaskan makna setiap kolom dan harus unik.",
    "Satu baris mewakili satu transaksi, permohonan, atau kegiatan.",
    "Kode layanan disimpan terpisah dari nama, tanggal, dan nilai.",
    "Hindari menggabung sel di area data yang akan difilter."
   ],
   "steps": [
    "Buka berkas latihan kosong dan beri nama sheet Register_Layanan.",
    "Tulis header No, Kode, Tanggal, Unit, Uraian, Status.",
    "Isi tiga baris data contoh dengan jenis data yang konsisten.",
    "Pilih sel D3 dan amati Name Box serta Formula Bar.",
    "Simpan berkas sebagai Register_Layanan_v01.xlsx."
   ],
   "example": "A1=Kode; B1=Tanggal; C1=Unit; D1=Uraian; A2=LAY-001; C2=Keuangan",
   "exercise": "Buat tabel 6 kolom berisi 5 permohonan fiktif dari dua unit, tanpa baris kosong di tengah.",
   "check": {
    "question": "Apa struktur yang paling aman untuk rekap layanan?",
    "options": [
     "Satu baris per permohonan dan satu header per kolom",
     "Satu sel berisi semua data permohonan",
     "Header berulang setiap dua baris",
     "Gabung semua baris agar rapi"
    ],
    "answer": 0,
    "why": "Tabel datar dengan header tunggal mendukung sort, filter, rumus, dan audit."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Tabel register lima permohonan berformat XLSX",
   "cautions": [
    "Jangan memasukkan nomor identitas mahasiswa asli pada data latihan."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Mengidentifikasi elemen worksheet dan membangun satu tabel data layanan yang konsisten. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-01-01",
   "duration": 14
  },
  {
   "title": "Memasukkan kode, tanggal, angka, dan data fiktif",
   "objective": "Menghindari perubahan format otomatis yang merusak kode atau tanggal.",
   "points": [
    "Nomor surat, NIM contoh, dan kode akun dapat diperlakukan sebagai teks.",
    "Tanggal harus berupa nilai tanggal, bukan hanya tampilan teks.",
    "Nilai rupiah disimpan sebagai angka tanpa mengetik \"Rp\" di dalam nilai sel.",
    "Angka yang tersimpan sebagai teks dapat mengganggu SUM.",
    "Gunakan contoh sintetis, bukan daftar mahasiswa atau rekening riil."
   ],
   "steps": [
    "Buat kolom Kode, Tanggal, Nominal, dan Status.",
    "Ketik kode 00127 sebagai teks agar nol di depan tetap ada.",
    "Masukkan tanggal lalu pilih format tanggal Indonesia.",
    "Isi nominal 1500000 sebagai angka, bukan \"Rp 1.500.000\".",
    "Uji SUM pada dua nilai nominal."
   ],
   "example": "Kode: 00127 (teks) · Tanggal: 18/11/2026 · Nominal: 1500000 (angka)",
   "exercise": "Buat tiga baris data pembayaran fiktif; tunjukkan perbedaan angka dan kode teks.",
   "check": {
    "question": "Mengapa kode 00127 perlu perhatian khusus?",
    "options": [
     "Excel dapat menghilangkan nol depan bila dianggap angka",
     "Kode selalu berupa tanggal",
     "Semua kode harus dihitung dengan SUM",
     "Tanggal wajib digabung dalam sel kode"
    ],
    "answer": 0,
    "why": "Kode identifikasi adalah label; format teks membantu menjaga nol di depan."
   },
   "roles": [
    "keuangan",
    "kemahasiswaan",
    "akademik"
   ],
   "requiresOffice": false,
   "output": "Dataset sintetis tiga transaksi dengan tipe data benar",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menghindari perubahan format otomatis yang merusak kode atau tanggal. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-01-02",
   "duration": 14
  },
  {
   "title": "Format Rupiah, persen, tanggal, dan keterbacaan",
   "objective": "Menerapkan Number Format tanpa mengubah nilai sumber perhitungan.",
   "points": [
    "Format mata uang menambah lambang dan pemisah ribuan secara visual.",
    "Persentase 0,25 dapat ditampilkan sebagai 25%.",
    "Tanggal perlu disajikan konsisten di seluruh kolom.",
    "Gunakan judul dan satuan rupiah atau orang dalam header.",
    "Atur lebar kolom agar nomor dan nilai tidak terpotong."
   ],
   "steps": [
    "Pilih kolom nominal pada tabel latihan.",
    "Terapkan format Number atau Accounting dengan simbol rupiah.",
    "Masukkan 0,85 pada kolom capaian lalu terapkan Percentage.",
    "Atur format tanggal secara konsisten.",
    "Periksa bahwa Formula Bar tetap memuat nilai numerik."
   ],
   "example": "1500000 → Rp1.500.000 (tampilan); 0,85 → 85%",
   "exercise": "Format tabel rekap pagu–realisasi dengan kolom Rp dan persentase capaian.",
   "check": {
    "question": "Jika 0,85 diformat sebagai persen, hasil tampil adalah ...",
    "options": [
     "85%",
     "0,85%",
     "850%",
     "8,5%"
    ],
    "answer": 0,
    "why": "Format Percentage mengalikan tampilan pecahan dengan 100."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": true,
   "output": "Tabel dua nilai dengan format rupiah dan persentase",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menerapkan Number Format tanpa mengubah nilai sumber perhitungan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-01-03",
   "duration": 14
  },
  {
   "title": "Navigasi, seleksi rentang, dan shortcut kerja",
   "objective": "Memakai shortcut dasar untuk memilih, mencari, dan menyimpan data administratif.",
   "points": [
    "Name Box memudahkan navigasi ke sel yang diketahui.",
    "Ctrl+F mencari nilai pada worksheet.",
    "Ctrl+S menyimpan perubahan.",
    "Ctrl+Z membatalkan tindakan; Ctrl+C dan Ctrl+V menyalin.",
    "Ctrl+Shift+L dapat mengaktifkan filter pada rentang dengan header."
   ],
   "steps": [
    "Buka workbook Register_Layanan.",
    "Pilih A1:F6 menggunakan Name Box atau seleksi mouse.",
    "Tekan Ctrl+F dan cari kode LAY-003.",
    "Salin satu sel contoh ke baris lain lalu batalkan dengan Ctrl+Z.",
    "Simpan perubahan dengan Ctrl+S."
   ],
   "example": "Navigasi: A1:F6 → Ctrl+F → LAY-003 → Esc → Ctrl+S",
   "exercise": "Temukan dua kode layanan pada file simulasi dan tulis alamat selnya.",
   "check": {
    "question": "Shortcut untuk menyimpan workbook ialah ...",
    "options": [
     "Ctrl+S",
     "Ctrl+P",
     "Alt+F5",
     "Ctrl+M"
    ],
    "answer": 0,
    "why": "Ctrl+S adalah shortcut simpan yang umum digunakan."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Daftar alamat dua sel hasil pencarian",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Memakai shortcut dasar untuk memilih, mencari, dan menyimpan data administratif. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-01-04",
   "duration": 14
  },
  {
   "title": "Penamaan berkas, folder, dan versi kerja",
   "objective": "Mengelola berkas latihan untuk menghindari salah versi atau kehilangan sumber.",
   "points": [
    "Pisahkan folder Sumber, Proses, dan Final sesuai SOP lokal.",
    "Nama berkas berisi unit, periode, jenis, dan versi.",
    "Simpan sumber mentah dalam keadaan tidak berubah.",
    "CSV cocok untuk transfer tabel sederhana tetapi tidak mempertahankan rumus/format.",
    "Pemberian akses mengikuti ketentuan unit, bukan sekadar kemampuan teknis."
   ],
   "steps": [
    "Buat folder Latihan/01_Sumber, 02_Proses, 03_Final.",
    "Simpan file latihan sebagai Rekap_Layanan_2026_v01.xlsx.",
    "Buat satu salinan versi v02 setelah revisi.",
    "Ekspor satu sheet menjadi CSV untuk pertukaran data.",
    "Periksa apakah CSV kehilangan formula dan tampilan format."
   ],
   "example": "2026_Kemahasiswaan_RegisterBeasiswa_v02.xlsx",
   "exercise": "Rancang struktur folder semester berjalan dengan berkas sumber dan final terpisah.",
   "check": {
    "question": "Untuk menghindari kehilangan rumus saat berbagi, format yang lebih tepat adalah ...",
    "options": [
     "XLSX",
     "CSV",
     "TXT",
     "JPG"
    ],
    "answer": 0,
    "why": "XLSX mempertahankan struktur workbook, formula, dan pemformatan."
   },
   "roles": [
    "arsip",
    "semua"
   ],
   "requiresOffice": false,
   "output": "Struktur folder serta skema penamaan versi kerja",
   "cautions": [
    "Retensi dan akses arsip harus mengikuti ketentuan institusi."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Mengelola berkas latihan untuk menghindari salah versi atau kehilangan sumber. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-01-05",
   "duration": 14
  },
  {
   "title": "Perhitungan pagu, realisasi, dan selisih",
   "objective": "Menghitung selisih anggaran menggunakan referensi sel sehingga dapat diperiksa ulang.",
   "points": [
    "Rumus dimulai dengan tanda sama dengan (=).",
    "Selisih anggaran = pagu dikurangi realisasi.",
    "Gunakan referensi C2 dan D2, bukan menulis nominal ulang.",
    "Salin rumus ke baris berikut untuk mengurangi kesalahan ketik.",
    "Selisih negatif perlu diverifikasi sebagai kemungkinan melebihi pagu, bukan otomatis kesalahan."
   ],
   "steps": [
    "Buka kasus Keuangan pada simulator Excel.",
    "Perhatikan C2=Pagu dan D2=Realisasi.",
    "Isi E2 dengan =C2-D2.",
    "Salin logika ke E3 sampai E7 bila perlu.",
    "Periksa selisih dengan satu perhitungan manual."
   ],
   "example": "C2=5000000; D2=4200000; E2 =C2-D2 → 800000",
   "exercise": "Buat kolom Selisih untuk enam kegiatan, lalu identifikasi nilai negatif.",
   "check": {
    "question": "Rumus selisih pagu minus realisasi adalah ...",
    "options": [
     "=C2-D2",
     "=C2+D2",
     "=C2*D2",
     "=D2/C2"
    ],
    "answer": 0,
    "why": "Pengurangan nilai anggaran terhadap realisasi menghasilkan sisa alokasi."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": false,
   "output": "Rekap enam kegiatan dengan selisih tepat",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menghitung selisih anggaran menggunakan referensi sel sehingga dapat diperiksa ulang. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-02-01",
   "duration": 18
  },
  {
   "title": "Rekap total, rata-rata, maksimum, dan hitungan baris",
   "objective": "Merangkum volume layanan dan penggunaan dana tanpa mengetik penjumlahan manual.",
   "points": [
    "SUM menjumlahkan angka pada rentang.",
    "AVERAGE menghitung rata-rata angka.",
    "MAX dan MIN mencari nilai tertinggi/terendah.",
    "COUNT menghitung sel numerik; COUNTA menghitung sel tak kosong.",
    "Sel total diberi label jelas dan tidak dicampur ke rentang data berikutnya."
   ],
   "steps": [
    "Perhatikan enam nilai realisasi D2:D7.",
    "Hitung =SUM(D2:D7) pada D8.",
    "Hitung =AVERAGE(D2:D7) pada C9.",
    "Bandingkan =MAX(D2:D7) dengan angka terbesar.",
    "Periksa hasil dengan kalkulator sebagai kontrol."
   ],
   "example": "=SUM(D2:D7) · =AVERAGE(D2:D7) · =MAX(D2:D7)",
   "exercise": "Hitung total, rata-rata, dan realisasi tertinggi dari dataset latihan.",
   "check": {
    "question": "Fungsi untuk mendapatkan nilai rata-rata adalah ...",
    "options": [
     "AVERAGE",
     "COUNTIF",
     "LEFT",
     "VLOOKUP"
    ],
    "answer": 0,
    "why": "AVERAGE menghitung rerata numerik dari rentang."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Tiga indikator agregat dengan nilai sesuai sumber",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Merangkum volume layanan dan penggunaan dana tanpa mengetik penjumlahan manual. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-02-02",
   "duration": 18
  },
  {
   "title": "Status layanan menggunakan IF",
   "objective": "Mengategorikan keterlambatan atau kelebihan realisasi memakai aturan sederhana.",
   "points": [
    "IF memiliki kondisi, hasil bila benar, dan hasil bila salah.",
    "Operator logika mencakup >, >=, <, <=, =, dan <>.",
    "Hasil berupa teks harus diapit tanda kutip.",
    "Status otomatis berubah ketika data sumber berubah.",
    "Kebijakan threshold harus disetujui unit sebelum dipakai pada data nyata."
   ],
   "steps": [
    "Tetapkan aturan simulasi: realisasi melebihi pagu = \"Melebihi\".",
    "Pilih F2 dan masukkan =IF(D2>C2;\"Melebihi\";\"Terkendali\").",
    "Tekan Enter dan amati hasilnya.",
    "Ubah sementara D2 lebih besar dari C2.",
    "Kembalikan data dan pastikan status otomatis berubah."
   ],
   "example": "=IF(D2>C2;\"Melebihi\";\"Terkendali\")",
   "exercise": "Buat label status untuk satu baris dan jelaskan logika keputusan.",
   "check": {
    "question": "Fungsi IF digunakan untuk ...",
    "options": [
     "Memilih keluaran berdasarkan suatu kondisi",
     "Menggabung slide presentasi",
     "Menyimpan file PDF",
     "Menyisipkan gambar"
    ],
    "answer": 0,
    "why": "IF menguji kondisi dan mengembalikan salah satu nilai."
   },
   "roles": [
    "keuangan",
    "perencanaan",
    "kemahasiswaan"
   ],
   "requiresOffice": false,
   "output": "Satu kolom status dengan pengujian kondisi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Mengategorikan keterlambatan atau kelebihan realisasi memakai aturan sederhana. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-02-03",
   "duration": 18
  },
  {
   "title": "Sort, Filter, Table, dan validasi pengisian",
   "objective": "Menyaring daftar permohonan tanpa merusak hubungan antarkolom.",
   "points": [
    "Sort menyusun urutan baris berdasarkan kolom tertentu.",
    "Filter menyembunyikan baris yang tidak memenuhi kriteria.",
    "Excel Table memudahkan pengembangan data.",
    "Data Validation dapat membatasi daftar status atau unit.",
    "Jangan melakukan sort hanya pada satu kolom tabel yang saling berkaitan."
   ],
   "steps": [
    "Pilih seluruh rentang termasuk header pada file latihan.",
    "Gunakan Ctrl+T untuk membuat Table.",
    "Aktifkan filter kategori atau unit.",
    "Buat dropdown Status menggunakan Data Validation (Office desktop).",
    "Periksa total baris setelah filter dibersihkan."
   ],
   "example": "Status yang diizinkan: Diajukan; Diproses; Selesai; Perlu Perbaikan",
   "exercise": "Buat dropdown status untuk register layanan dan filter hanya data \"Diproses\".",
   "check": {
    "question": "Operasi yang menampilkan sebagian baris tanpa menghapus data adalah ...",
    "options": [
     "Filter",
     "Delete Row",
     "Cut",
     "Merge Cells"
    ],
    "answer": 0,
    "why": "Filter membatasi tampilan, sedangkan data sumber tetap ada."
   },
   "roles": [
    "kemahasiswaan",
    "arsip",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Register layanan dengan validasi status dan filter teruji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menyaring daftar permohonan tanpa merusak hubungan antarkolom. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-02-04",
   "duration": 18
  },
  {
   "title": "Grafik perbandingan target dan realisasi",
   "objective": "Menyajikan dua indikator dalam grafik yang terbaca oleh pimpinan unit.",
   "points": [
    "Grafik kolom cocok untuk membandingkan nilai antarkategori.",
    "Judul harus menjelaskan periode dan ukuran.",
    "Sumbu dan satuan mencegah salah baca angka.",
    "Pilih data yang sudah diperiksa, bukan sembarang range.",
    "Jangan memakai efek 3D yang mengaburkan perbedaan."
   ],
   "steps": [
    "Siapkan tabel Unit, Target, Realisasi.",
    "Pilih tiga kolom tersebut termasuk header.",
    "Insert → Clustered Column Chart.",
    "Beri judul \"Target dan Realisasi Layanan\".",
    "Tambahkan label satuan dan tulis satu temuan di bawah grafik."
   ],
   "example": "Unit A: Target 40, Realisasi 36; Unit B: Target 50, Realisasi 45",
   "exercise": "Buat grafik dua unit dan jelaskan unit yang pencapaiannya lebih tinggi.",
   "check": {
    "question": "Grafik yang tepat untuk perbandingan kategori adalah ...",
    "options": [
     "Kolom berkelompok",
     "Diagram hias tanpa data",
     "Tabel isi Word",
     "Transisi Morph"
    ],
    "answer": 0,
    "why": "Kolom berkelompok memudahkan perbandingan nilai dua seri per kategori."
   },
   "roles": [
    "perencanaan",
    "kepegawaian",
    "kemahasiswaan"
   ],
   "requiresOffice": true,
   "output": "Grafik kolom berlabel, dengan satu kalimat insight",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menyajikan dua indikator dalam grafik yang terbaca oleh pimpinan unit. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-02-05",
   "duration": 18
  },
  {
   "title": "COUNTIF dan SUMIF menurut unit atau status",
   "objective": "Menghasilkan rekap kategori pada register administrasi.",
   "points": [
    "COUNTIF menghitung entri yang cocok dengan satu kriteria.",
    "SUMIF menjumlahkan nilai pada baris yang cocok dengan kategori.",
    "Range kriteria dan range nilai harus mempunyai panjang sesuai.",
    "Kriteria teks ditulis dalam kutip atau dirujuk dari sel.",
    "Lakukan pemeriksaan silang dengan data terfilter."
   ],
   "steps": [
    "Identifikasi kategori pada kolom B di kasus Keuangan.",
    "Tulis =COUNTIF(B2:B7;\"Operasional\") pada G5.",
    "Hitung total realisasi kategori yang sama dengan SUMIF.",
    "Bandingkan hasil dengan daftar baris Operasional.",
    "Catat apakah ada spasi atau ejaan kategori yang tidak konsisten."
   ],
   "example": "=COUNTIF(B2:B7;\"Operasional\") · =SUMIF(B2:B7;\"Operasional\";D2:D7)",
   "exercise": "Buat rekap jumlah transaksi dan total realisasi untuk dua kelompok kegiatan.",
   "check": {
    "question": "Rumus untuk menghitung jumlah baris berkategori \"Operasional\" adalah ...",
    "options": [
     "=COUNTIF(B2:B7;\"Operasional\")",
     "=SUM(B2:B7)",
     "=LEFT(B2:B7)",
     "=MAX(B2:B7)"
    ],
    "answer": 0,
    "why": "COUNTIF menghitung jumlah sel sesuai kriteria."
   },
   "roles": [
    "keuangan",
    "perencanaan",
    "aset"
   ],
   "requiresOffice": false,
   "output": "Rekap per kategori dengan hasil terverifikasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Menghasilkan rekap kategori pada register administrasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-03-01",
   "duration": 22
  },
  {
   "title": "SUMIFS dan COUNTIFS untuk dua kriteria",
   "objective": "Melakukan rekap bersyarat berdasarkan unit dan status atau periode.",
   "points": [
    "SUMIFS mendukung lebih dari satu pasangan kriteria.",
    "COUNTIFS menghitung baris yang memenuhi semua kondisi sekaligus.",
    "Rentang harus sejajar dan mempunyai jumlah baris sama.",
    "Kriteria tanggal sebaiknya memakai batas awal dan akhir yang jelas.",
    "Periksa apakah kriteria kosong atau bergeser."
   ],
   "steps": [
    "Buka data transaksi simulasi pada Microsoft Excel.",
    "Siapkan kolom Unit, Status, dan Nilai.",
    "Hitung jumlah layanan Unit A yang berstatus Selesai dengan COUNTIFS.",
    "Hitung nilai transaksi Unit A pada bulan tertentu menggunakan SUMIFS.",
    "Bandingkan satu sampel baris secara manual."
   ],
   "example": "=COUNTIFS(B2:B100;\"Akademik\";E2:E100;\"Selesai\")",
   "exercise": "Buat rekap dua kriteria: jenis layanan dan status, lengkap dengan kontrol total.",
   "check": {
    "question": "Keunggulan COUNTIFS dibanding COUNTIF adalah ...",
    "options": [
     "Dapat memeriksa beberapa kriteria sekaligus",
     "Selalu membuat grafik",
     "Menghapus duplikat secara otomatis",
     "Membuat surat massal"
    ],
    "answer": 0,
    "why": "COUNTIFS menerima lebih dari satu rentang kondisi dan kriteria."
   },
   "roles": [
    "kemahasiswaan",
    "akademik",
    "arsip"
   ],
   "requiresOffice": true,
   "output": "Tabel ringkasan dua kriteria dengan kontrol angka",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Melakukan rekap bersyarat berdasarkan unit dan status atau periode. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-03-02",
   "duration": 22
  },
  {
   "title": "XLOOKUP dan VLOOKUP untuk referensi kode",
   "objective": "Mengisi uraian berdasarkan kode akun, kegiatan, atau layanan fiktif.",
   "points": [
    "Pencarian tepat menghindari salah pengambilan kode.",
    "VLOOKUP mencari pada kolom pertama tabel referensi.",
    "XLOOKUP menentukan kolom pencarian dan hasil secara terpisah.",
    "Jika kode tidak ada, tampilkan status perlu verifikasi.",
    "Jangan memakai lookup sebagai pembenaran mengabaikan pengecekan sumber."
   ],
   "steps": [
    "Siapkan sheet Master_Kode dengan Kode dan Nama Layanan.",
    "Beri kode contoh ADM-01 dan ADM-02.",
    "Gunakan XLOOKUP dengan parameter if_not_found pada kolom uraian.",
    "Bandingkan tiga hasil kode terhadap master.",
    "Tandai kode yang tidak ditemukan untuk perbaikan."
   ],
   "example": "=XLOOKUP(A2;Master_Kode!A:A;Master_Kode!B:B;\"Perlu cek\")",
   "exercise": "Petakan lima kode layanan ke nama layanan dengan pencarian exact match.",
   "check": {
    "question": "Fungsi yang mencari kode dan mengembalikan nilai terkait adalah ...",
    "options": [
     "XLOOKUP",
     "AVERAGE",
     "RIGHT",
     "COUNT"
    ],
    "answer": 0,
    "why": "XLOOKUP memetakan kunci ke kolom hasil yang ditentukan."
   },
   "roles": [
    "akademik",
    "arsip",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Tabel lima kode dengan penandaan kode tidak dikenal",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengisi uraian berdasarkan kode akun, kegiatan, atau layanan fiktif. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-03-03",
   "duration": 22
  },
  {
   "title": "Membersihkan teks, tanggal, dan duplikat data",
   "objective": "Mengoreksi masalah umum sebelum rekap atau penggabungan berkas.",
   "points": [
    "TRIM menghapus spasi berlebih.",
    "LEFT, RIGHT, dan MID mengurai sebagian kode teks.",
    "DATE dan TEXT membantu membentuk/menampilkan tanggal.",
    "Remove Duplicates menghapus baris tertentu dan perlu backup.",
    "Data pribadi dalam daftar layanan harus disamarkan saat latihan."
   ],
   "steps": [
    "Duplikasi file sumber terlebih dahulu.",
    "Buat kolom bantu Kode_Bersih menggunakan TRIM.",
    "Normalisasi ejaan status dan format tanggal.",
    "Identifikasi duplikat berdasarkan kombinasi kode dan tanggal.",
    "Catat baris yang dikoreksi serta alasan perbaikannya."
   ],
   "example": "=TRIM(A2) · =LEFT(B2;3) · =TEXT(C2;\"yyyy-mm-dd\")",
   "exercise": "Bersihkan sepuluh kode layanan contoh yang mengandung spasi dan satu duplikat.",
   "check": {
    "question": "Fungsi untuk menghapus spasi berlebih pada teks adalah ...",
    "options": [
     "TRIM",
     "SUMIFS",
     "MAX",
     "AND"
    ],
    "answer": 0,
    "why": "TRIM menghilangkan spasi berlebih yang sering menyebabkan lookup gagal."
   },
   "roles": [
    "kemahasiswaan",
    "arsip",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Dataset bersih plus log perbaikan data",
   "cautions": [
    "Jangan langsung menghapus duplikat tanpa memeriksa apakah transaksi memang berbeda."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengoreksi masalah umum sebelum rekap atau penggabungan berkas. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-03-04",
   "duration": 22
  },
  {
   "title": "PivotTable untuk rekap multidimensi unit kerja",
   "objective": "Menyusun ringkasan data berdasarkan unit, periode, dan status.",
   "points": [
    "PivotTable mengelompokkan baris sumber ke dalam kategori ringkasan.",
    "Rows dan Columns menentukan dimensi analisis.",
    "Values menentukan agregasi COUNT atau SUM.",
    "Refresh diperlukan bila sumber mengalami perubahan.",
    "Periksa grand total dan nilai kosong sebelum mengirim laporan."
   ],
   "steps": [
    "Jadikan data sumber sebagai Excel Table.",
    "Insert → PivotTable pada sheet baru.",
    "Tempatkan Unit di Rows, Status di Columns, dan Kode di Values sebagai Count.",
    "Aktifkan filter atau slicer bila diperlukan.",
    "Cocokkan Grand Total dengan jumlah record sumber."
   ],
   "example": "Rows: Unit; Columns: Status; Values: Count of Kode",
   "exercise": "Buat PivotTable jumlah layanan per unit dan status pada data fiktif.",
   "check": {
    "question": "Komponen Values pada PivotTable berfungsi untuk ...",
    "options": [
     "Menentukan perhitungan atau agregasi",
     "Memberi nama file",
     "Mengatur margin Word",
     "Membuat animasi objek"
    ],
    "answer": 0,
    "why": "Values memuat ukuran numerik atau hitungan yang diringkas."
   },
   "roles": [
    "kemahasiswaan",
    "perencanaan",
    "kepegawaian"
   ],
   "requiresOffice": true,
   "output": "PivotTable dengan grand total teruji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Menyusun ringkasan data berdasarkan unit, periode, dan status. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-03-05",
   "duration": 22
  },
  {
   "title": "Power Query: konsolidasi berkas per bulan",
   "objective": "Menggabungkan laporan berformat sama ke tabel bersih yang bisa diperbarui.",
   "points": [
    "Power Query memisahkan proses impor, transformasi, dan muat data.",
    "Sumber perlu memiliki skema kolom konsisten.",
    "Kolom tanggal dan nominal harus bertipe benar.",
    "Simpan langkah transformasi untuk audit.",
    "Refresh harus diuji dan dicatat hasilnya."
   ],
   "steps": [
    "Siapkan tiga CSV sintetis Januari–Maret di satu folder.",
    "Data → Get Data → From Folder (Excel desktop yang mendukung).",
    "Gabungkan tabel lalu tetapkan tipe kolom.",
    "Hilangkan baris kosong dan beri kolom Periode.",
    "Load ke Table dan cocokkan jumlah record dengan sumber."
   ],
   "example": "3 CSV × 10 baris → target 30 transaksi tanpa baris ganda",
   "exercise": "Bangun kueri konsolidasi bulanan dan catat total sumber sebelum/sesudah.",
   "check": {
    "question": "Keunggulan Power Query untuk laporan berulang adalah ...",
    "options": [
     "Transformasi dapat diulang melalui refresh",
     "Setiap file harus diketik ulang",
     "Hanya bisa untuk slide",
     "Selalu menghapus file sumber"
    ],
    "answer": 0,
    "why": "Power Query menyimpan rangkaian langkah transformasi yang dapat dijalankan kembali."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": true,
   "output": "Query konsolidasi dan log jumlah baris",
   "cautions": [
    "Periksa perbedaan nama kolom serta format tanggal antarsumber."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menggabungkan laporan berformat sama ke tabel bersih yang bisa diperbarui. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-04-01",
   "duration": 28
  },
  {
   "title": "Rekonsiliasi dan audit rumus keuangan",
   "objective": "Mencari ketidaksesuaian antara buku kas, register, dan ringkasan sebelum pelaporan.",
   "points": [
    "Rekonsiliasi membandingkan dua sumber pada periode dan definisi yang sama.",
    "Selisih nol belum tentu cukup jika transaksi yang berpasangan tidak sesuai.",
    "Gunakan kode unik dan tanggal sebagai pengenal.",
    "IFERROR tidak boleh menutupi error tanpa penjelasan.",
    "Audit trail menyimpan asal, koreksi, dan pemeriksa."
   ],
   "steps": [
    "Pisahkan sheet BukuKas, Register, dan Rekonsiliasi.",
    "Samakan rentang tanggal serta satuan nominal.",
    "Gunakan lookup untuk mencocokkan referensi transaksi.",
    "Hitung selisih per transaksi dan total.",
    "Susun daftar pengecualian untuk diklarifikasi."
   ],
   "example": "Sumber A Rp12.500.000; Sumber B Rp12.300.000; selisih Rp200.000",
   "exercise": "Rancang log rekonsiliasi 10 transaksi fiktif dengan status Cocok/Perlu Periksa.",
   "check": {
    "question": "Jika total dua sumber berbeda, tindakan yang tepat adalah ...",
    "options": [
     "Telusuri transaksi pengecualian dan sumbernya",
     "Ubah salah satu total agar sama",
     "Sembunyikan selisih",
     "Buang baris tanpa dokumentasi"
    ],
    "answer": 0,
    "why": "Selisih harus ditelusuri, bukan disesuaikan tanpa dasar."
   },
   "roles": [
    "keuangan",
    "aset"
   ],
   "requiresOffice": true,
   "output": "Laporan pengecualian dan log koreksi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Mencari ketidaksesuaian antara buku kas, register, dan ringkasan sebelum pelaporan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-04-02",
   "duration": 28
  },
  {
   "title": "Dashboard KPI layanan, slicer, dan pembacaan grafik",
   "objective": "Menampilkan kinerja unit secara ringkas dengan definisi indikator yang jelas.",
   "points": [
    "KPI harus memiliki definisi, satuan, periode, sumber, dan pemilik.",
    "Dashboard ringkas mengutamakan target, realisasi, dan deviasi.",
    "Slicer membantu memfilter PivotTable atau model yang didukung.",
    "Grafik dibuat dari data yang dapat ditelusuri.",
    "Warna indikator tidak boleh menjadi satu-satunya pembeda status."
   ],
   "steps": [
    "Tentukan tiga KPI: layanan selesai, waktu proses, dan tingkat penyelesaian.",
    "Bangun PivotTable dari register sintetis.",
    "Buat kartu metrik dan grafik bulanan.",
    "Tambahkan slicer Unit untuk laporan internal.",
    "Uji bahwa perubahan filter mengubah indikator secara konsisten."
   ],
   "example": "KPI: 80 layanan masuk, 72 selesai, tingkat penyelesaian 90%",
   "exercise": "Bangun dashboard satu halaman untuk kepala bagian dan tulis dua temuan.",
   "check": {
    "question": "Definisi KPI yang dapat diaudit minimal memuat ...",
    "options": [
     "Rumus, satuan, periode, dan sumber data",
     "Hanya warna ikon",
     "Hanya judul menarik",
     "Animasi dan transisi"
    ],
    "answer": 0,
    "why": "KPI perlu didefinisikan agar semua pembaca menafsirkan angka yang sama."
   },
   "roles": [
    "perencanaan",
    "kemahasiswaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Dashboard KPI dan kamus indikator",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menampilkan kinerja unit secara ringkas dengan definisi indikator yang jelas. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-04-03",
   "duration": 28
  },
  {
   "title": "Analisis skenario anggaran dan target",
   "objective": "Membandingkan beberapa asumsi sebelum mengusulkan alokasi kerja.",
   "points": [
    "Skenario dasar memakai asumsi yang terdokumentasi.",
    "What-If Analysis dapat membantu melihat sensitivitas input.",
    "Pisahkan input, perhitungan, dan output.",
    "Jangan mengubah angka realisasi historis untuk memenuhi skenario.",
    "Tuliskan keterbatasan agar hasil tidak diperlakukan sebagai kepastian."
   ],
   "steps": [
    "Buat input pagu, volume layanan, dan biaya rata-rata.",
    "Tetapkan tiga skenario: dasar, efisiensi, peningkatan layanan.",
    "Hitung kebutuhan =volume×biaya.",
    "Bandingkan kebutuhan dan batas pagu.",
    "Tuliskan rekomendasi dengan konsekuensi tiap opsi."
   ],
   "example": "100 layanan × Rp150.000 = Rp15.000.000; naik 10% volume = 110 layanan",
   "exercise": "Buat matriks tiga skenario kebutuhan operasional untuk rapat perencanaan.",
   "check": {
    "question": "Pemisahan penting dalam analisis skenario adalah ...",
    "options": [
     "Asumsi, rumus, dan hasil",
     "Data asli dan judul disatukan dalam satu sel",
     "Semua sel dikunci tanpa penjelasan",
     "Mencampur satuan rupiah dan orang"
    ],
    "answer": 0,
    "why": "Asumsi terdokumentasi memungkinkan evaluasi sensitivitas perubahan."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": true,
   "output": "Matriks tiga skenario beserta asumsi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Membandingkan beberapa asumsi sebelum mengusulkan alokasi kerja. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-04-04",
   "duration": 28
  },
  {
   "title": "Kontrol akses, proteksi, dan data sensitif",
   "objective": "Mendesain workbook administrasi yang meminimalkan risiko kesalahan dan kebocoran.",
   "points": [
    "Gunakan dataset anonim/sintetis untuk latihan.",
    "Sheet Protection membatasi penyuntingan, tetapi bukan solusi keamanan penuh.",
    "Identifikasi kolom input, formula, dan data terbatas.",
    "Berikan akses sesuai kebutuhan kerja dan kebijakan institusi.",
    "Pencadangan dan pemeriksaan versi perlu memiliki penanggung jawab."
   ],
   "steps": [
    "Tandai kolom input yang boleh diubah pengguna.",
    "Warnai sel rumus dan beri keterangan sumber.",
    "Lindungi formula sesuai versi Excel desktop.",
    "Buat versi publik yang tidak memuat nama/NIM/nomor rekening nyata.",
    "Dokumentasikan peran penyiap dan pemeriksa."
   ],
   "example": "Label akses latihan: PUBLIK / INTERNAL / TERBATAS (contoh, bukan klasifikasi resmi)",
   "exercise": "Susun checklist keamanan sebelum workbook rekap diserahkan antarunit.",
   "check": {
    "question": "Perlindungan yang benar bagi data mahasiswa adalah ...",
    "options": [
     "Membatasi akses dan menggunakan data sintetis saat pelatihan",
     "Membagikan seluruh NIM di layar publik",
     "Mengunggah berkas ke layanan tak disetujui",
     "Menghilangkan backup tanpa alasan"
    ],
    "answer": 0,
    "why": "Prinsip minimisasi data dan pengendalian akses mengurangi risiko pemrosesan."
   },
   "roles": [
    "kemahasiswaan",
    "akademik",
    "arsip",
    "kepegawaian"
   ],
   "requiresOffice": true,
   "output": "Checklist kontrol akses dan versi contoh tanpa data pribadi",
   "cautions": [
    "Patuhi kebijakan keamanan kampus dan peraturan terkait; OfficeLab bukan konsultasi kepatuhan."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Mendesain workbook administrasi yang meminimalkan risiko kesalahan dan kebocoran. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-04-05",
   "duration": 28
  },
  {
   "title": "Model anggaran multi-unit dan konsolidasi usulan",
   "objective": "Merancang model perencanaan yang dapat diaudit dari sumber unit hingga ringkasan kampus.",
   "points": [
    "Kamus kode akun harus konsisten antarunit.",
    "Pemisahan input, asumsi, formula, dan output menurunkan risiko.",
    "Gunakan validasi skema dan rekap per unit.",
    "Buat pengujian jumlah total sebelum menggabungkan usulan.",
    "Semua perubahan pagu perlu jejak persetujuan di alur kerja institusi."
   ],
   "steps": [
    "Buat sheet Master_Akun dan daftar unit.",
    "Siapkan format input usulan yang sama untuk tujuh unit.",
    "Validasi kode, unit, jumlah, dan asumsi.",
    "Konsolidasi serta tampilkan total per jenis belanja.",
    "Dokumentasikan sumber dan anomali untuk ditindaklanjuti."
   ],
   "example": "Anggaran Unit A Rp20 juta, B Rp15 juta, C Rp12 juta → total Rp47 juta",
   "exercise": "Siapkan model anggaran sintetis 3 unit dengan cek duplikat dan ringkasan keputusan.",
   "check": {
    "question": "Apa kontrol paling penting pada konsolidasi usulan?",
    "options": [
     "Keseragaman kode dan validasi total sumber",
     "Memilih warna sel yang acak",
     "Menghapus semua formula",
     "Membiarkan periode tidak jelas"
    ],
    "answer": 0,
    "why": "Keseragaman master data dan rekonsiliasi total mendukung akurasi."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": true,
   "output": "Model input-konsolidasi-output dengan catatan audit",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Merancang model perencanaan yang dapat diaudit dari sumber unit hingga ringkasan kampus. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-05-01",
   "duration": 35
  },
  {
   "title": "Kamus data, Power Pivot, dan ukuran DAX",
   "objective": "Memahami model data berelasi serta definisi metrik yang digunakan berulang.",
   "points": [
    "Tabel fakta berisi transaksi/kejadian; dimensi berisi referensi.",
    "Kunci relasi harus unik pada sisi dimensi yang sesuai.",
    "Measure DAX menghitung nilai dalam konteks filter.",
    "Pisahkan metrik sumber dari metrik turunan.",
    "Versi Excel dan lisensi dapat membatasi fitur Power Pivot."
   ],
   "steps": [
    "Definisikan tabel Fakta_Layanan, Dim_Unit, dan Dim_Tanggal.",
    "Buat kunci hubungan UnitID dan DateKey.",
    "Tentukan measure Total_Layanan dan Layanan_Selesai.",
    "Uji filter berdasarkan unit dan bulan.",
    "Simpan kamus data beserta rumus serta satuannya."
   ],
   "example": "Total Layanan := COUNTROWS(Fakta_Layanan)  (contoh DAX)",
   "exercise": "Gambar skema model tiga tabel dan dokumentasikan tiga ukuran kinerja.",
   "check": {
    "question": "Peran tabel dimensi dalam model analitik ialah ...",
    "options": [
     "Memberi atribut untuk pengelompokan dan penyaringan",
     "Mengubah PPT menjadi CSV",
     "Menghapus semua transaksi",
     "Menyimpan seluruh nilai di satu gambar"
    ],
    "answer": 0,
    "why": "Dimensi memberikan konteks untuk fakta yang dihitung."
   },
   "roles": [
    "perencanaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Skema model data dan kamus tiga measure",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Memahami model data berelasi serta definisi metrik yang digunakan berulang. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-05-02",
   "duration": 35
  },
  {
   "title": "Validasi lintas-sistem dan penanganan ketidaksesuaian",
   "objective": "Menguji kecocokan berkas dari unit berbeda sebelum digabung dalam laporan institusi.",
   "points": [
    "Sistem berbeda dapat memakai kode, format, dan periode yang tidak identik.",
    "Mapping kode adalah keputusan bisnis yang harus disahkan pemilik data.",
    "Uji keunikan, kelengkapan, kewajaran, dan keterlacakan.",
    "Laporan pengecualian harus memuat sumber serta prioritas.",
    "Jangan menganggap hasil cocok berarti sumber bebas masalah."
   ],
   "steps": [
    "Buat dua file register sintetis dengan format kolom berbeda.",
    "Susun tabel pemetaan kode antarberkas.",
    "Gabungkan berdasarkan kunci yang telah dibersihkan.",
    "Catat record tanpa pasangan dan nilai yang berbeda.",
    "Buat ringkasan temuan serta usulan perbaikan master data."
   ],
   "example": "UNIT-01 ↔ Bagian Keuangan; status \"Done\" ↔ \"Selesai\" (perlu disahkan)",
   "exercise": "Rancang dan uji 4 aturan validasi untuk konsolidasi 20 record.",
   "check": {
    "question": "Kode berbeda antar-sistem seharusnya ...",
    "options": [
     "Dipetakan dengan aturan terdokumentasi dan divalidasi pemilik data",
     "Dianggap sama tanpa cek",
     "Dihapus dari semua sumber",
     "Disalin sebagai gambar"
    ],
    "answer": 0,
    "why": "Pemetaan terdokumentasi menjaga makna dan jejak audit."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Log aturan validasi, mismatch, dan tindakan koreksi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menguji kecocokan berkas dari unit berbeda sebelum digabung dalam laporan institusi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-05-03",
   "duration": 35
  },
  {
   "title": "Otomatisasi laporan dengan Office Scripts atau VBA",
   "objective": "Menilai kelayakan otomasi tugas rutin dan risiko perubahan otomatis.",
   "points": [
    "Pilih proses yang stabil, berulang, dan terukur sebelum otomatisasi.",
    "VBA digunakan pada lingkungan Office desktop yang mendukung; Office Scripts pada lingkungan yang sesuai.",
    "Kode harus menghindari penghapusan data sumber dan mendokumentasikan hasil.",
    "Uji dengan data contoh serta buat salinan cadangan.",
    "Makro dari sumber tidak tepercaya berisiko keamanan."
   ],
   "steps": [
    "Petakan tugas bulanan: impor, cek, rekap, dan ekspor.",
    "Tentukan input/output serta kondisi gagal.",
    "Buat pseudocode langkah yang akan diotomatisasi.",
    "Uji pada copy dataset dan bandingkan hasil manual.",
    "Susun checklist review dan prosedur rollback."
   ],
   "example": "Pseudocode: IMPORT → VALIDASI → REKAP → CEK TOTAL → EXPORT → LOG",
   "exercise": "Tulis rancangan otomasi rekap bulanan disertai pengujian dan penanganan error.",
   "check": {
    "question": "Mengapa otomasi harus diuji pada salinan data?",
    "options": [
     "Agar kegagalan tidak merusak data operasional",
     "Agar data riil tersebar lebih luas",
     "Karena semua makro selalu aman",
     "Agar audit trail tidak perlu dibuat"
    ],
    "answer": 0,
    "why": "Pengujian pada data simulasi membatasi dampak kesalahan skrip."
   },
   "roles": [
    "keuangan",
    "perencanaan",
    "arsip"
   ],
   "requiresOffice": true,
   "output": "Pseudocode dan rencana uji otomatisasi",
   "cautions": [
    "Jangan mengaktifkan makro dari sumber yang tidak dikenal."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menilai kelayakan otomasi tugas rutin dan risiko perubahan otomatis. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-05-04",
   "duration": 35
  },
  {
   "title": "Proyek akhir Excel: dashboard monitoring administrasi",
   "objective": "Menggabungkan pembersihan data, formula, validasi, dan visualisasi menjadi keluaran siap tinjau.",
   "points": [
    "Proyek akhir harus memakai data sintetis yang konsisten dan terdokumentasi.",
    "Definisi indikator dipisahkan dari dataset sumber.",
    "Rekonsiliasi jumlah harus dapat diulang reviewer.",
    "Dashboard menyajikan temuan dan risiko, bukan sekadar dekorasi.",
    "Catat batas penggunaan dan hasil uji."
   ],
   "steps": [
    "Pilih kasus unit: anggaran, mahasiswa, atau layanan.",
    "Siapkan minimal 30 record sintetis dan kamus kolom.",
    "Buat validasi data dan rekap bersyarat.",
    "Tambahkan dua grafik dan tiga KPI terdefinisi.",
    "Susun catatan audit dan ringkasan rekomendasi satu halaman."
   ],
   "example": "Keluaran: data_sumber.xlsx + dashboard.xlsx + kamus_data + daftar_uji",
   "exercise": "Bangun dashboard monitoring bulanan; tunjukkan kontrol total, 3 KPI, dan 2 grafik.",
   "check": {
    "question": "Proyek Excel dianggap siap tinjau jika ...",
    "options": [
     "Data, rumus, hasil, dan sumber dapat ditelusuri",
     "Hanya tampilannya menarik",
     "Tak ada definisi KPI",
     "Seluruh sel disembunyikan tanpa catatan"
    ],
    "answer": 0,
    "why": "Dashboard yang layak ditinjau dapat direproduksi dari sumbernya."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Paket proyek Excel dengan dashboard dan dokumentasi audit",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menggabungkan pembersihan data, formula, validasi, dan visualisasi menjadi keluaran siap tinjau. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "E-05-05",
   "duration": 35
  }
 ],
 "word": [
  {
   "title": "Mengenal dokumen dinas, Ribbon, dan area kerja",
   "objective": "Menemukan menu dasar Word serta menjelaskan bagian dokumen administrasi.",
   "points": [
    "Dokumen resmi memiliki tujuan, pihak tujuan, tanggal, isi, dan pengesahan sesuai pedoman institusi.",
    "Ribbon dan Quick Access Toolbar menyediakan fungsi pengeditan.",
    "Navigation Pane membantu menemukan heading pada dokumen panjang.",
    "Status Bar menampilkan halaman dan jumlah kata.",
    "Gunakan kop institusi hanya sesuai identitas dan format yang disahkan."
   ],
   "steps": [
    "Buka Word dan buat dokumen baru.",
    "Ketik judul \"Catatan Layanan Administrasi\".",
    "Tuliskan tanggal dan unit pengelola secara fiktif.",
    "Temukan menu Home, Insert, Layout, dan Review.",
    "Simpan sebagai 2026_CatatanLayanan_v01.docx."
   ],
   "example": "Judul: CATATAN LAYANAN ADMINISTRASI · Unit: Bagian Administrasi Kampus",
   "exercise": "Buat satu halaman pengantar layanan dan identifikasi lima bagian dokumen.",
   "check": {
    "question": "Tab Word yang umum digunakan untuk pemeriksaan ejaan adalah ...",
    "options": [
     "Review",
     "Slide Show",
     "Data",
     "Formulas"
    ],
    "answer": 0,
    "why": "Fitur pemeriksaan ada di tab Review/Peninjauan."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Satu halaman catatan layanan berjudul dan bernama berkas",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menemukan menu dasar Word serta menjelaskan bagian dokumen administrasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-01-01",
   "duration": 14
  },
  {
   "title": "Struktur surat: nomor, tanggal, perihal, dan lampiran",
   "objective": "Menyusun elemen surat administratif tanpa menggunakan data atau tanda tangan asli.",
   "points": [
    "Nomor, perihal, tanggal, lampiran, alamat tujuan, dan isi mengikuti tata naskah dinas institusi.",
    "Format latihan hanyalah contoh dan bukan format resmi kampus.",
    "Satu surat idealnya fokus pada satu maksud komunikasi.",
    "Jangan menyalin tanda tangan atau stempel pejabat.",
    "Verifikasi data tujuan sebelum distribusi."
   ],
   "steps": [
    "Buka template latihan Surat Undangan.",
    "Isi nomor surat fiktif dan tanggal contoh.",
    "Tuliskan perihal yang spesifik dan tujuan surat.",
    "Sampaikan agenda, lokasi, dan waktu secara ringkas.",
    "Periksa unsur surat berdasarkan format instansi masing-masing."
   ],
   "example": "Nomor: 001/ADM-LAT/XI/2026 · Perihal: Undangan Koordinasi Pelayanan",
   "exercise": "Lengkapi unsur surat undangan kegiatan pelatihan administrasi dengan identitas sintetis.",
   "check": {
    "question": "Unsur yang umumnya menjelaskan inti surat adalah ...",
    "options": [
     "Perihal",
     "Warna kertas",
     "Wallpaper",
     "Nomor slide"
    ],
    "answer": 0,
    "why": "Perihal merangkum pokok komunikasi."
   },
   "roles": [
    "arsip",
    "kepegawaian",
    "akademik"
   ],
   "requiresOffice": false,
   "output": "Draft surat undangan dengan unsur terisi",
   "cautions": [
    "Nomor surat dan format harus disesuaikan pedoman tata naskah dinas kampus."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menyusun elemen surat administratif tanpa menggunakan data atau tanda tangan asli. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-01-02",
   "duration": 14
  },
  {
   "title": "Paragraf, font, spasi, dan Styles dasar",
   "objective": "Merapikan dokumen yang mudah dibaca dan konsisten tanpa spasi manual berlebihan.",
   "points": [
    "Gunakan Styles untuk judul dan heading agar konsisten.",
    "Atur jarak antarparagraf melalui Paragraph Spacing.",
    "Tab stop dan tabel lebih stabil dibanding rentetan spasi.",
    "Gunakan ukuran teks dan kontras yang cukup.",
    "Hindari memakai Enter berulang untuk pindah halaman."
   ],
   "steps": [
    "Sorot judul dokumen lalu terapkan Heading 1.",
    "Atur body text dengan font yang sesuai pedoman instansi.",
    "Tentukan line spacing dan spacing after secara konsisten.",
    "Aktifkan Show/Hide ¶ untuk memeriksa format tersembunyi.",
    "Simpan dan lihat Print Preview."
   ],
   "example": "Heading 1: NOTULEN RAPAT · Normal: uraian agenda",
   "exercise": "Rapikan satu halaman notulen agar judul, subjudul, dan paragraf terbaca.",
   "check": {
    "question": "Cara paling stabil membuat judul bab konsisten adalah ...",
    "options": [
     "Menggunakan Styles/Heading",
     "Menambah 15 spasi",
     "Mengulang tombol Enter",
     "Memasang gambar judul"
    ],
    "answer": 0,
    "why": "Styles menyatukan format dan mendukung navigasi dokumen."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Dokumen dengan Heading dan spacing konsisten",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Merapikan dokumen yang mudah dibaca dan konsisten tanpa spasi manual berlebihan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-01-03",
   "duration": 14
  },
  {
   "title": "Tabel sederhana untuk daftar hadir dan register",
   "objective": "Menyusun informasi dalam kolom tanpa merusak susunan saat dicetak.",
   "points": [
    "Tabel cocok untuk data dengan kolom berulang.",
    "Header perlu memuat satuan dan deskripsi jelas.",
    "Hindari tabel lebar yang melewati margin.",
    "Ulangi baris header bila tabel melintasi halaman.",
    "Pisahkan data latihan fiktif dari daftar hadir nyata."
   ],
   "steps": [
    "Insert → Table dan buat tabel empat kolom.",
    "Isi No, Unit, Agenda, dan Status Kehadiran.",
    "Tambahkan tiga baris nama unit fiktif.",
    "Atur lebar kolom dan alignment sesuai jenis data.",
    "Periksa Print Preview."
   ],
   "example": "No | Unit | Agenda | Hadir (Ya/Tidak)",
   "exercise": "Buat tabel daftar hadir enam peserta fiktif pada editor Word OfficeLab.",
   "check": {
    "question": "Untuk membuat kolom yang rapi, fitur yang lebih tepat adalah ...",
    "options": [
     "Table",
     "Rentetan spasi",
     "Animasi",
     "PivotTable"
    ],
    "answer": 0,
    "why": "Tabel menjaga sel sejajar secara konsisten."
   },
   "roles": [
    "kepegawaian",
    "kemahasiswaan"
   ],
   "requiresOffice": false,
   "output": "Tabel daftar hadir dengan header dan enam baris",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menyusun informasi dalam kolom tanpa merusak susunan saat dicetak. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-01-04",
   "duration": 14
  },
  {
   "title": "Simpan, ekspor PDF, dan kontrol versi dokumen",
   "objective": "Memelihara versi penyuntingan dan mengirim dokumen dalam format sesuai tujuan.",
   "points": [
    "DOCX cocok untuk naskah yang masih perlu diedit.",
    "PDF berguna untuk distribusi tampilan final tetapi bukan bukti persetujuan otomatis.",
    "Nama file perlu mengandung jenis dokumen, periode, dan versi.",
    "Sebelum berbagi, periksa komentar, metadata, serta lampiran.",
    "Pengelolaan arsip dan retensi mengikuti kebijakan kampus."
   ],
   "steps": [
    "Simpan surat latihan sebagai DOCX versi v01.",
    "Buka File → Save As atau Export sesuai versi Office.",
    "Ekspor salinan PDF untuk pratinjau.",
    "Pastikan halaman, tabel, dan margin tidak berubah.",
    "Simpan hasil pada folder latihan Final."
   ],
   "example": "2026_Arsip_RegisterSurat_v01.docx → 2026_Arsip_RegisterSurat_FINAL.pdf",
   "exercise": "Buat dua versi berkas (editable dan distribusi) dan periksa nama serta isi.",
   "check": {
    "question": "Format yang lazim dipakai untuk naskah Word yang masih diedit adalah ...",
    "options": [
     "DOCX",
     "PPTX",
     "MP3",
     "PNG"
    ],
    "answer": 0,
    "why": "DOCX mempertahankan kemampuan mengedit dokumen Word."
   },
   "roles": [
    "arsip",
    "semua"
   ],
   "requiresOffice": false,
   "output": "DOCX dan salinan PDF yang sudah diperiksa",
   "cautions": [
    "PDF bukan pengganti mekanisme tanda tangan/persetujuan yang diwajibkan."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Memelihara versi penyuntingan dan mengirim dokumen dalam format sesuai tujuan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-01-05",
   "duration": 14
  },
  {
   "title": "Surat undangan rapat dan koordinasi",
   "objective": "Menulis surat undangan yang singkat, lengkap, dan dapat ditindaklanjuti.",
   "points": [
    "Undangan memerlukan tujuan, agenda, waktu, tempat, dan kontak penyelenggara.",
    "Tuliskan tanggal lengkap agar tidak ambigu.",
    "Pilih kata kerja yang jelas dan sopan.",
    "Lakukan pemeriksaan pihak tujuan dan lampiran.",
    "Gunakan format dan jalur paraf resmi jika dokumen dipakai operasional."
   ],
   "steps": [
    "Buka template Undangan Rapat di laboratorium Word.",
    "Isi perihal dan daftar tujuan fiktif.",
    "Tuliskan agenda dan waktu dengan tanggal penuh.",
    "Sisipkan tabel agenda bila diperlukan.",
    "Tinjau semua detail sebelum ekspor."
   ],
   "example": "Agenda: Evaluasi pelayanan administrasi; Waktu: Jumat, 20 November 2026, 09.00 WITA",
   "exercise": "Susun undangan koordinasi antarbagi­an dengan agenda dua butir.",
   "check": {
    "question": "Informasi yang penting ada dalam undangan rapat adalah ...",
    "options": [
     "Tanggal, waktu, agenda, dan tempat",
     "Hanya judul kegiatan",
     "Hanya nama pengirim",
     "Hanya dekorasi"
    ],
    "answer": 0,
    "why": "Rincian logistik dan tujuan membuat undangan dapat ditindaklanjuti."
   },
   "roles": [
    "arsip",
    "kepegawaian",
    "perencanaan"
   ],
   "requiresOffice": false,
   "output": "Surat undangan satu halaman dan agenda",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menulis surat undangan yang singkat, lengkap, dan dapat ditindaklanjuti. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-02-01",
   "duration": 18
  },
  {
   "title": "Surat tugas dan surat keterangan layanan",
   "objective": "Menyusun naskah administratif berdasarkan data layanan yang dapat diverifikasi.",
   "points": [
    "Surat tugas menjelaskan penugasan, lingkup, dan periode.",
    "Surat keterangan menyatakan fakta yang perlu verifikasi sumber.",
    "Identitas dalam latihan harus sintetis.",
    "Hindari klaim yang tidak memiliki dasar pemeriksaan.",
    "Dokumen final memerlukan otorisasi sesuai aturan instansi."
   ],
   "steps": [
    "Pilih format Surat Tugas pada editor.",
    "Isi unit pelaksana dan uraian tugas contoh.",
    "Tambahkan lokasi dan periode penugasan.",
    "Masukkan daftar tugas dengan penomoran.",
    "Tinjau apakah ada ketentuan atau rincian yang tertinggal."
   ],
   "example": "Surat tugas: pendataan inventaris ruang layanan pada 23–25 November 2026",
   "exercise": "Buat surat tugas pendataan aset untuk dua petugas fiktif.",
   "check": {
    "question": "Pernyataan yang tepat tentang surat keterangan adalah ...",
    "options": [
     "Isi harus dapat diverifikasi terhadap sumber",
     "Isi bebas ditebak",
     "Dapat diterbitkan tanpa wewenang",
     "Tidak perlu pemeriksaan identitas"
    ],
    "answer": 0,
    "why": "Surat keterangan memuat pernyataan faktual yang harus terverifikasi."
   },
   "roles": [
    "akademik",
    "kepegawaian",
    "aset"
   ],
   "requiresOffice": false,
   "output": "Draft surat tugas dengan ruang lingkup dan periode",
   "cautions": [
    "Jangan mencantumkan nama pejabat, tanda tangan, atau stempel tanpa izin."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menyusun naskah administratif berdasarkan data layanan yang dapat diverifikasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-02-02",
   "duration": 18
  },
  {
   "title": "Notulen dan berita acara kegiatan",
   "objective": "Mencatat agenda, pembahasan, keputusan, penanggung jawab, dan tindak lanjut.",
   "points": [
    "Notulen membedakan pembahasan dari keputusan.",
    "Tindak lanjut sebaiknya menyebut pemilik tugas dan tenggat.",
    "Berita acara mencatat kejadian atau serah terima secara terstruktur.",
    "Catatan rapat harus netral dan berbasis hasil nyata.",
    "Identifikasi peserta rapat hanya pada dokumen yang diotorisasi."
   ],
   "steps": [
    "Buat judul Notulen Rapat Evaluasi Layanan.",
    "Tambahkan agenda, peserta fiktif, dan waktu.",
    "Tulis dua keputusan dengan action owner.",
    "Sisipkan tabel Tindak Lanjut: aksi, penanggung jawab, tenggat.",
    "Periksa bahwa poin keputusan tidak bertentangan dengan isi."
   ],
   "example": "Keputusan 01: unit layanan memperbarui form registrasi; PIC: Tim Layanan; Tenggat: akhir bulan",
   "exercise": "Susun notulen tiga agenda dan dua tindak lanjut lengkap.",
   "check": {
    "question": "Elemen yang membedakan notulen yang dapat ditindaklanjuti adalah ...",
    "options": [
     "Keputusan, penanggung jawab, dan tenggat",
     "Jenis font semata",
     "Jumlah warna halaman",
     "Hanya daftar hadir"
    ],
    "answer": 0,
    "why": "Penanggung jawab dan tenggat membuat keputusan dapat dipantau."
   },
   "roles": [
    "perencanaan",
    "arsip",
    "semua"
   ],
   "requiresOffice": false,
   "output": "Notulen dua halaman dengan tabel tindak lanjut",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Mencatat agenda, pembahasan, keputusan, penanggung jawab, dan tindak lanjut. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-02-03",
   "duration": 18
  },
  {
   "title": "Tabel inventaris, daftar hadir, dan lampiran",
   "objective": "Menambahkan lampiran tabel yang tidak terpotong dan mudah diverifikasi.",
   "points": [
    "Judul lampiran menjelaskan konteks dan periode.",
    "Kolom kuantitas harus mencantumkan satuan.",
    "Kolom status menggunakan nilai yang seragam.",
    "Periksa pengulangan header saat pergantian halaman.",
    "Tabel yang terlalu besar dapat dibuat sebagai lampiran Excel terpisah."
   ],
   "steps": [
    "Buat dokumen Berita Acara Inventaris.",
    "Sisipkan tabel No, Barang, Jumlah, Satuan, Kondisi.",
    "Isi lima contoh aset sintetis.",
    "Atur heading Lampiran 1 dan nomor halaman.",
    "Tinjau apakah semua kolom muat saat dicetak."
   ],
   "example": "Meja kerja | 3 | unit | Baik; Printer | 1 | unit | Perlu Servis",
   "exercise": "Buat lampiran inventaris lima barang dengan satuan dan kondisi.",
   "check": {
    "question": "Informasi yang perlu ada pada tabel inventaris adalah ...",
    "options": [
     "Nama barang, jumlah, satuan, kondisi",
     "Warna latar saja",
     "Animasi masuk",
     "Nama file tanpa isi"
    ],
    "answer": 0,
    "why": "Kolom yang konsisten mempermudah pemeriksaan inventaris."
   },
   "roles": [
    "aset",
    "arsip"
   ],
   "requiresOffice": false,
   "output": "Lampiran inventaris siap cetak",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menambahkan lampiran tabel yang tidak terpotong dan mudah diverifikasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-02-04",
   "duration": 18
  },
  {
   "title": "Proofreading, bahasa resmi, dan pemeriksaan akhir",
   "objective": "Menghapus kesalahan ketik serta ambiguitas pada surat dan laporan.",
   "points": [
    "Proofreading memeriksa bahasa, angka, tanggal, dan referensi.",
    "Gunakan fitur Spelling & Grammar sebagai bantuan, bukan keputusan final.",
    "Pastikan penulisan nama unit, jabatan, dan singkatan konsisten.",
    "Cari klaim berlebihan atau kalimat yang tidak jelas pelakunya.",
    "Cek apakah semua lampiran benar-benar ada."
   ],
   "steps": [
    "Buka surat latihan yang telah disusun.",
    "Gunakan Ctrl+H untuk menemukan istilah yang tidak konsisten.",
    "Baca ulang seluruh angka dan tanggal.",
    "Periksa format tanggal serta alamat tujuan.",
    "Tandai dokumen selesai diperiksa pada log latihan."
   ],
   "example": "Sebelum: \"data nya akan diperbaiki mungkin\" → Sesudah: \"Tim Layanan memperbaiki data pada 20 November 2026\" bila faktanya terverifikasi",
   "exercise": "Buat daftar minimal lima kesalahan surat contoh dan perbaikannya.",
   "check": {
    "question": "Bukti proofreading yang baik berupa ...",
    "options": [
     "Log kesalahan dan perbaikannya",
     "Mengganti semua kata secara otomatis tanpa cek",
     "Menghapus lampiran tanpa catatan",
     "Mengabaikan angka"
    ],
    "answer": 0,
    "why": "Catatan pemeriksaan memperjelas apa yang telah diperbaiki."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Lembar pemeriksaan akhir yang terisi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menghapus kesalahan ketik serta ambiguitas pada surat dan laporan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-02-05",
   "duration": 18
  },
  {
   "title": "Styles, heading, dan Navigation Pane laporan",
   "objective": "Mengubah laporan panjang menjadi dokumen berstruktur yang mudah dinavigasi.",
   "points": [
    "Heading 1–3 membentuk hierarki judul yang dapat dipakai Navigation Pane.",
    "Jangan membuat heading hanya dengan ukuran font manual.",
    "Style dapat diubah satu kali untuk banyak bagian.",
    "Struktur dokumen perlu konsisten sampai subbab.",
    "Hindari subbab terlalu dalam jika tidak diperlukan."
   ],
   "steps": [
    "Buat laporan dengan bagian Pendahuluan, Kegiatan, Hasil, dan Penutup.",
    "Terapkan Heading 1 untuk bagian utama.",
    "Gunakan Heading 2 untuk rincian di bawahnya.",
    "Aktifkan View → Navigation Pane.",
    "Periksa urutan dan kedalaman heading."
   ],
   "example": "Heading 1: Hasil Kegiatan; Heading 2: Rekap Layanan Mahasiswa",
   "exercise": "Format laporan dua halaman dengan sedikitnya empat heading yang dapat dinavigasi.",
   "check": {
    "question": "Navigation Pane paling efektif bila judul menggunakan ...",
    "options": [
     "Heading Styles",
     "Spasi manual",
     "Text Box acak",
     "Format mata uang"
    ],
    "answer": 0,
    "why": "Navigation Pane mengenali hierarki heading berbasis Styles."
   },
   "roles": [
    "perencanaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Laporan terstruktur dengan navigasi judul",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengubah laporan panjang menjadi dokumen berstruktur yang mudah dinavigasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-03-01",
   "duration": 22
  },
  {
   "title": "Section Break, header/footer, dan halaman berbeda",
   "objective": "Mengatur dokumen campuran potret dan lanskap dengan penomoran terkontrol.",
   "points": [
    "Page Break memulai halaman tanpa mengubah bagian.",
    "Section Break membuat pengaturan halaman berbeda antarbagian.",
    "Header/Footer dapat mengandung nomor halaman dan identitas dokumen.",
    "Link to Previous menentukan keterkaitan antarseksi.",
    "Periksa halaman kosong akibat break berlebihan."
   ],
   "steps": [
    "Buat laporan dua bagian dengan halaman potret.",
    "Sisipkan Section Break sebelum tabel lebar.",
    "Ubah orientasi hanya section tabel menjadi lanskap.",
    "Atur header dan nomor halaman untuk setiap section.",
    "Tinjau total halaman dalam Print Preview."
   ],
   "example": "Bagian I potret; Lampiran tabel lanskap; Bagian III kembali potret",
   "exercise": "Buat laporan tiga section dengan lampiran lanskap satu halaman.",
   "check": {
    "question": "Fitur untuk mengubah orientasi hanya satu bagian ialah ...",
    "options": [
     "Section Break",
     "Ctrl+B",
     "Track Changes",
     "SmartArt"
    ],
    "answer": 0,
    "why": "Section Break memisahkan pengaturan tata letak bagian dokumen."
   },
   "roles": [
    "keuangan",
    "perencanaan",
    "arsip"
   ],
   "requiresOffice": true,
   "output": "Dokumen multi-section dengan nomor halaman konsisten",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengatur dokumen campuran potret dan lanskap dengan penomoran terkontrol. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-03-02",
   "duration": 22
  },
  {
   "title": "Daftar isi otomatis, caption, dan daftar tabel",
   "objective": "Membuat navigasi laporan resmi dari Styles dan caption.",
   "points": [
    "Daftar isi otomatis mengambil heading yang diberi style.",
    "Caption memberi nomor dan label tabel/gambar.",
    "Daftar tabel dapat dibangkitkan dari caption yang benar.",
    "Perbarui field setelah judul atau nomor berubah.",
    "Lakukan pemeriksaan silang nomor tabel serta rujukannya."
   ],
   "steps": [
    "Terapkan Heading 1/2 pada laporan latihan.",
    "References → Table of Contents → Automatic Table.",
    "Beri caption pada dua tabel simulasi.",
    "Sisipkan List of Tables bila diperlukan.",
    "Update Table dan periksa halaman tujuan."
   ],
   "example": "Daftar isi: 1 Pendahuluan; 2 Metode; 3 Hasil; Lampiran",
   "exercise": "Buat daftar isi otomatis dari minimal empat heading dan dua caption tabel.",
   "check": {
    "question": "Jika judul bab berubah, tindakan yang diperlukan adalah ...",
    "options": [
     "Memperbarui daftar isi/field",
     "Mengubah semua heading jadi gambar",
     "Menghapus caption",
     "Membuat slide baru"
    ],
    "answer": 0,
    "why": "Field otomatis perlu di-update agar daftar isi tetap sinkron."
   },
   "roles": [
    "perencanaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Laporan dengan daftar isi dan caption yang diperbarui",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Membuat navigasi laporan resmi dari Styles dan caption. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-03-03",
   "duration": 22
  },
  {
   "title": "Mail Merge surat massal dari Excel",
   "objective": "Menghasilkan surat personalisasi dari daftar penerima yang sah dan terverifikasi.",
   "points": [
    "Mail Merge menghubungkan dokumen utama dengan tabel sumber.",
    "Kolom sumber harus memiliki header yang stabil.",
    "Gunakan daftar fiktif pada praktik agar tidak memproses data pribadi riil.",
    "Pratinjau tiap variasi sebelum membuat surat massal.",
    "Cek format tanggal, alamat, dan kode yang kehilangan nol depan."
   ],
   "steps": [
    "Siapkan data Excel Nama_Penerima dan Unit dengan 5 contoh fiktif.",
    "Buka Mailings → Start Mail Merge → Letters.",
    "Pilih daftar penerima sintetis.",
    "Sisipkan merge fields dalam badan surat.",
    "Preview Results dan tinjau setiap surat sebelum Finish & Merge."
   ],
   "example": "Yth. «Nama_Penerima», Unit «Unit» — Undangan Pelatihan Administrasi",
   "exercise": "Buat lima surat undangan personalisasi dari satu template dan data contoh.",
   "check": {
    "question": "Sumber data yang digunakan Mail Merge adalah ...",
    "options": [
     "Tabel penerima dengan header konsisten",
     "Gambar slide tanpa teks",
     "File audio",
     "Transisi presentasi"
    ],
    "answer": 0,
    "why": "Mail Merge memetakan field dari daftar penerima."
   },
   "roles": [
    "kepegawaian",
    "kemahasiswaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Lima surat hasil merge yang diperiksa",
   "cautions": [
    "Jangan menyebarluaskan daftar penerima atau mencetak data tanpa kewenangan."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Menghasilkan surat personalisasi dari daftar penerima yang sah dan terverifikasi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-03-04",
   "duration": 22
  },
  {
   "title": "Komentar, Track Changes, dan alur reviu",
   "objective": "Mengelola masukan dari beberapa peninjau tanpa kehilangan sejarah revisi.",
   "points": [
    "Track Changes menandai perubahan pada naskah.",
    "Komentar menjelaskan alasan atau pertanyaan review.",
    "Reviu perlu membedakan saran, keputusan, dan persetujuan.",
    "Accept/Reject dilakukan dengan kewenangan yang jelas.",
    "Dokumen final perlu diperiksa dari komentar atau revisi tersisa."
   ],
   "steps": [
    "Aktifkan Review → Track Changes di Word asli.",
    "Perbaiki satu paragraf sehingga tanda revisi muncul.",
    "Tambahkan dua komentar yang dapat ditindaklanjuti.",
    "Simulasikan penerimaan satu perubahan dan penolakan yang lain.",
    "Buat salinan final setelah pemeriksaan menyeluruh."
   ],
   "example": "Komentar: \"Mohon verifikasi dasar angka 72 layanan selesai\"",
   "exercise": "Lakukan putaran reviu surat tugas dengan minimal tiga catatan revisi.",
   "check": {
    "question": "Fungsi Track Changes adalah ...",
    "options": [
     "Menandai riwayat perubahan naskah",
     "Menghitung SUMIFS",
     "Membuat grafik otomatis",
     "Menyisipkan video"
    ],
    "answer": 0,
    "why": "Track Changes membantu meninjau perubahan sebelum disetujui."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Dokumen revisi dan log keputusan editor",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengelola masukan dari beberapa peninjau tanpa kehilangan sejarah revisi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-03-05",
   "duration": 22
  },
  {
   "title": "Template surat terstandar dan Content Controls",
   "objective": "Membuat template yang dapat diisi ulang tanpa mengubah format wajib.",
   "points": [
    "Template memisahkan bagian tetap dari area input.",
    "Content Controls menyediakan kontrol isian tertentu pada Word desktop.",
    "Placeholder harus memiliki petunjuk yang tidak ambigu.",
    "Template perlu menjalani review dan persetujuan format oleh unit berwenang.",
    "Lakukan uji cetak dan uji pemakaian oleh staf berbeda."
   ],
   "steps": [
    "Identifikasi elemen surat yang selalu tetap.",
    "Rancang placeholder tanggal, nomor, perihal, dan penerima.",
    "Gunakan Developer → Controls bila tersedia.",
    "Batasi perubahan bagian tetap sesuai kebutuhan.",
    "Uji template dengan tiga contoh isi berbeda."
   ],
   "example": "Field: [Nomor Surat] [Tanggal] [Unit Tujuan] [Agenda] [Pejabat Berwenang]",
   "exercise": "Rancang template undangan layanan yang dapat digunakan lima kegiatan berbeda.",
   "check": {
    "question": "Keuntungan Content Controls pada template adalah ...",
    "options": [
     "Menyediakan isian terstruktur",
     "Mengganti semua paragraf menjadi gambar",
     "Membuat PivotTable",
     "Menyembunyikan ketentuan lembaga"
    ],
    "answer": 0,
    "why": "Content Controls mengarahkan pengisian pada bagian yang ditentukan."
   },
   "roles": [
    "arsip",
    "kepegawaian",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Template surat teruji dengan panduan isian",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Membuat template yang dapat diisi ulang tanpa mengubah format wajib. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-04-01",
   "duration": 28
  },
  {
   "title": "Formulir layanan mahasiswa dan alur verifikasi",
   "objective": "Merancang formulir administrasi yang ringkas dan meminimalkan pengumpulan data.",
   "points": [
    "Minta hanya informasi yang benar-benar diperlukan untuk layanan.",
    "Buat instruksi dan field wajib yang jelas.",
    "Pisahkan bagian diisi pemohon dengan bagian verifikasi petugas.",
    "Gunakan kode proses, status, dan tanggal yang konsisten.",
    "Formulir tidak otomatis menjadi sah tanpa SOP serta persetujuan institusi."
   ],
   "steps": [
    "Pilih kasus Permohonan Surat Keterangan.",
    "Tentukan informasi minimum yang dibutuhkan.",
    "Susun tabel Formulir dan kotak checklist persyaratan.",
    "Tambahkan ruang nomor agenda serta status verifikasi petugas.",
    "Uji keterbacaan pada cetak hitam putih."
   ],
   "example": "Jenis layanan: Surat Keterangan; Status berkas: Lengkap/Perlu Perbaikan",
   "exercise": "Buat formulir permohonan satu halaman menggunakan identitas contoh.",
   "check": {
    "question": "Prinsip yang tepat saat merancang formulir mahasiswa adalah ...",
    "options": [
     "Mengumpulkan hanya data yang diperlukan",
     "Meminta kata sandi akun mahasiswa",
     "Menyimpan semua identitas tanpa pembatasan",
     "Membuat kolom rahasia wajib untuk latihan"
    ],
    "answer": 0,
    "why": "Minimisasi data mendukung layanan yang proporsional dan aman."
   },
   "roles": [
    "kemahasiswaan",
    "akademik"
   ],
   "requiresOffice": true,
   "output": "Formulir layanan dengan instruksi dan alur verifikasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Merancang formulir administrasi yang ringkas dan meminimalkan pengumpulan data. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-04-02",
   "duration": 28
  },
  {
   "title": "Cross-reference, sitasi internal, dan rujukan tabel",
   "objective": "Menghubungkan bagian dokumen panjang agar nomor dan halaman tidak salah.",
   "points": [
    "Cross-reference merujuk heading, gambar, atau caption secara dinamis.",
    "Caption yang konsisten mendukung daftar gambar/tabel.",
    "Referensi sumber data perlu jelas dan dapat ditelusuri.",
    "Field sebaiknya diperbarui sebelum cetak.",
    "Periksa apakah rujukan tetap benar setelah pengeditan atau pemindahan bagian."
   ],
   "steps": [
    "Sisipkan dua tabel dengan caption pada laporan.",
    "Tambahkan kalimat \"Lihat Tabel 2\" menggunakan Cross-reference.",
    "Pindahkan salah satu tabel ke halaman lain.",
    "Perbarui field Ctrl+A lalu F9 bila didukung.",
    "Pastikan nomor pada naskah dan daftar tabel cocok."
   ],
   "example": "\"Realisasi program disajikan pada Tabel 2\" menggunakan rujukan dinamis",
   "exercise": "Susun laporan dengan dua referensi silang dan pemeriksaan nomor akhir.",
   "check": {
    "question": "Kelebihan Cross-reference dibanding mengetik nomor tabel manual adalah ...",
    "options": [
     "Nomor dapat diperbarui mengikuti struktur dokumen",
     "Semua data menjadi anonim otomatis",
     "Tidak perlu cek dokumen",
     "Tabel berubah menjadi slide"
    ],
    "answer": 0,
    "why": "Cross-reference menggunakan field yang dapat diperbarui."
   },
   "roles": [
    "perencanaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Dokumen dengan dua rujukan silang yang valid",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menghubungkan bagian dokumen panjang agar nomor dan halaman tidak salah. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-04-03",
   "duration": 28
  },
  {
   "title": "Metadata, redaksi, distribusi, dan arsip final",
   "objective": "Meninjau informasi tersembunyi serta versi yang dibagikan kepada pihak berbeda.",
   "points": [
    "Dokumen dapat memiliki komentar, properties, dan informasi tersembunyi.",
    "Document Inspector membantu memeriksa jenis metadata tertentu.",
    "Redaksi data harus benar-benar menghapus informasi pada keluaran final.",
    "Klasifikasi akses dan retensi mengikuti instrumen kearsipan kampus.",
    "Jangan mengandalkan warna kotak hitam semata untuk menyembunyikan isi sensitif."
   ],
   "steps": [
    "Buat salinan dokumen latihan, jangan memodifikasi sumber asli.",
    "Periksa Properties dan Document Inspector pada Word desktop.",
    "Hapus komentar/revisi yang tidak seharusnya dibagikan.",
    "Buat versi distribusi tanpa data contoh yang sensitif.",
    "Catat versi final, tanggal, serta penerima internal."
   ],
   "example": "Sumber_Kepegawaian_v03.docx → Distribusi_Internal_v03.pdf",
   "exercise": "Susun daftar pemeriksaan pradistribusi dokumen layanan.",
   "check": {
    "question": "Langkah aman sebelum mendistribusikan dokumen final adalah ...",
    "options": [
     "Memeriksa komentar, revisi, metadata, dan akses",
     "Memilih font paling dekoratif",
     "Membagikan file sumber ke semua orang",
     "Menyembunyikan teks dengan warna putih saja"
    ],
    "answer": 0,
    "why": "Metadata serta akses perlu ditinjau di samping isi yang terlihat."
   },
   "roles": [
    "arsip",
    "kepegawaian",
    "kemahasiswaan"
   ],
   "requiresOffice": true,
   "output": "Checklist distribusi dan log versi",
   "cautions": [
    "Lakukan redaksi dengan fitur serta prosedur yang benar dan uji hasilnya."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Meninjau informasi tersembunyi serta versi yang dibagikan kepada pihak berbeda. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-04-04",
   "duration": 28
  },
  {
   "title": "Dokumen batch dan standar audit revisi",
   "objective": "Menata produksi dokumen berulang sekaligus menjaga jejak perubahan.",
   "points": [
    "Dokumen batch memerlukan sumber data yang diverifikasi.",
    "Template, versi, dan pengesahan dipisahkan secara jelas.",
    "Log produksi mencatat jenis, nomor, waktu, dan status.",
    "Tiap file harus dapat ditelusuri ke sumbernya.",
    "Tinjauan sampling diperlukan untuk mendeteksi kesalahan yang berulang."
   ],
   "steps": [
    "Siapkan lima data surat simulasi.",
    "Gunakan template dan Mail Merge atau proses pengisian terkontrol.",
    "Beri nama file berdasarkan nomor contoh dan unit.",
    "Pilih dua sampel untuk audit isi serta format.",
    "Buat daftar pengendalian revisi dan status persetujuan."
   ],
   "example": "LOG: NOMOR, TANGGAL, JENIS SURAT, VERSI, PEMERIKSA, STATUS",
   "exercise": "Buat prosedur produksi 10 surat layanan dengan titik kontrol wajib.",
   "check": {
    "question": "Tujuan log produksi surat adalah ...",
    "options": [
     "Melacak sumber, versi, status, dan pemeriksa",
     "Menggantikan tanda tangan pejabat",
     "Menyembunyikan kesalahan",
     "Membuat grafik 3D"
    ],
    "answer": 0,
    "why": "Log produksi mendukung keterlacakan dan pengendalian mutu."
   },
   "roles": [
    "arsip",
    "akademik",
    "kepegawaian"
   ],
   "requiresOffice": true,
   "output": "SOP ringkas produksi batch dan daftar kontrol",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menata produksi dokumen berulang sekaligus menjaga jejak perubahan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-04-05",
   "duration": 28
  },
  {
   "title": "Menulis SOP administrasi berbasis proses",
   "objective": "Merancang SOP unit yang memuat tujuan, ruang lingkup, aktor, tahapan, dan kontrol.",
   "points": [
    "SOP menjelaskan proses yang berlaku berdasarkan kewenangan institusi.",
    "Tiap tahapan harus memiliki pemicu, pelaksana, input, output, dan batas waktu jika sah.",
    "Bagan alur harus konsisten dengan langkah tertulis.",
    "Risiko serta titik kontrol ditandai secara eksplisit.",
    "Draf latihan tidak otomatis berlaku sebagai SOP kampus."
   ],
   "steps": [
    "Pilih proses permohonan surat keterangan.",
    "Gambarkan alur masuk–verifikasi–proses–serah hasil.",
    "Tulis tabel langkah, pelaksana, dokumen, dan kontrol.",
    "Definisikan kondisi permohonan tidak lengkap.",
    "Lakukan simulasi satu kasus lalu revisi draf."
   ],
   "example": "Alur: Terima permohonan → Verifikasi persyaratan → Proses → Periksa → Serahkan",
   "exercise": "Buat draf SOP empat halaman untuk satu layanan administratif.",
   "check": {
    "question": "Komponen yang tepat pada satu langkah SOP adalah ...",
    "options": [
     "Pelaksana, input, tindakan, keluaran, dan kontrol",
     "Hanya nama unit",
     "Warna latar dan dekorasi",
     "Gambar tanpa penjelasan"
    ],
    "answer": 0,
    "why": "SOP yang dapat dipraktikkan menjelaskan siapa melakukan apa dan dengan bukti apa."
   },
   "roles": [
    "akademik",
    "kemahasiswaan",
    "arsip"
   ],
   "requiresOffice": true,
   "output": "Draf SOP dan diagram alur dengan catatan validasi",
   "cautions": [
    "SOP harus ditinjau serta disahkan oleh otoritas institusi sebelum digunakan."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Merancang SOP unit yang memuat tujuan, ruang lingkup, aktor, tahapan, dan kontrol. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-05-01",
   "duration": 35
  },
  {
   "title": "Mengintegrasikan tabel dan grafik Excel ke laporan Word",
   "objective": "Menyisipkan data analitik dengan pilihan tautan atau salinan yang tepat.",
   "points": [
    "Paste Link mengikuti sumber yang tetap dapat diakses dan berisiko putus saat dipindahkan.",
    "Embed menyimpan salinan objek di dalam dokumen, tetapi ukurannya bisa lebih besar.",
    "Bila angka direvisi, sumber perlu diperiksa ulang.",
    "Berikan caption dan referensi untuk tabel/grafik.",
    "Pastikan pembaca memahami periode serta satuan."
   ],
   "steps": [
    "Buat tabel KPI sintetis di Excel.",
    "Copy grafik dan tempel sebagai objek atau gambar sesuai kebutuhan.",
    "Uji apakah sumber tetap dapat diakses pada perangkat lain.",
    "Beri caption dan cantumkan sumber periode.",
    "Simpan dokumen final dan periksa tampilan cetak."
   ],
   "example": "Sumber: Rekap Layanan Unit, Triwulan III 2026 (data simulasi)",
   "exercise": "Buat laporan dua halaman dengan satu grafik dan satu tabel terverifikasi.",
   "check": {
    "question": "Risiko utama objek Word yang ditautkan ke Excel adalah ...",
    "options": [
     "Tautan bisa terputus bila sumber berpindah",
     "Word otomatis menghapus sumber",
     "Grafik selalu menjadi video",
     "Semua angka terverifikasi otomatis"
    ],
    "answer": 0,
    "why": "Objek tertaut membutuhkan file sumber yang dapat diakses."
   },
   "roles": [
    "perencanaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Laporan dengan dua objek data dan catatan sumber",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menyisipkan data analitik dengan pilihan tautan atau salinan yang tepat. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-05-02",
   "duration": 35
  },
  {
   "title": "Aksesibilitas dan pemeriksaan dokumen panjang",
   "objective": "Mengevaluasi dokumen yang siap dibaca secara digital maupun dicetak.",
   "points": [
    "Heading terstruktur memudahkan pembaca dan teknologi asistif.",
    "Alt text penting pada gambar yang memiliki makna.",
    "Tabel memerlukan header serta urutan baca logis.",
    "Accessibility Checker membantu, tetapi tetap perlu review manual.",
    "Periksa konsistensi bahasa, daftar isi, caption, dan nomor halaman."
   ],
   "steps": [
    "Buka laporan minimal enam halaman.",
    "Gunakan Review → Check Accessibility jika tersedia.",
    "Perbaiki heading bertingkat, alt text, dan label tabel.",
    "Periksa daftar isi serta rujukan otomatis.",
    "Lakukan uji cetak grayscale dan baca dari PDF."
   ],
   "example": "Bukti audit: 0 heading loncat, semua gambar bermakna punya deskripsi",
   "exercise": "Buat checklist aksesibilitas sepuluh poin untuk dokumen unit.",
   "check": {
    "question": "Satu praktik yang meningkatkan aksesibilitas adalah ...",
    "options": [
     "Heading semantik dan alt text untuk visual bermakna",
     "Memakai teks 6 pt",
     "Hanya membedakan kategori lewat warna",
     "Mengubah seluruh teks menjadi gambar"
    ],
    "answer": 0,
    "why": "Heading dan alternatif teks membantu navigasi dan pemahaman."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Laporan yang ditinjau beserta hasil Accessibility Checker",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Mengevaluasi dokumen yang siap dibaca secara digital maupun dicetak. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-05-03",
   "duration": 35
  },
  {
   "title": "Tata naskah, klasifikasi, dan retensi arsip institusi",
   "objective": "Mengenali hubungan dokumen Word dengan tata kelola arsip dinamis.",
   "points": [
    "Tata naskah dinas mengatur bentuk komunikasi resmi sesuai institusi.",
    "Klasifikasi arsip memudahkan penataan dan pencarian.",
    "Keamanan dan akses mengikuti klasifikasi yang ditetapkan.",
    "Jadwal retensi arsip (JRA) harus merujuk instrumen resmi yang berlaku.",
    "Jangan membuat angka retensi atau kewenangan pemusnahan dari asumsi pelatihan."
   ],
   "steps": [
    "Pilih tiga jenis dokumen: surat tugas, berita acara, dan laporan.",
    "Tentukan metadata minimum: nomor, tanggal, unit, perihal.",
    "Cari kode klasifikasi dari pedoman unit bila tersedia.",
    "Catat kebutuhan akses dan lokasi penyimpanan sesuai SOP.",
    "Susun register arsip latihan tanpa menetapkan retensi sembarang."
   ],
   "example": "Register contoh: ID, Kode Arsip [sesuai pedoman], Judul, Unit, Tanggal, Akses",
   "exercise": "Buat tabel metadata untuk enam dokumen fiktif dan tiga aturan kontrol.",
   "check": {
    "question": "Jangka retensi arsip resmi seharusnya diambil dari ...",
    "options": [
     "JRA institusi yang berlaku",
     "Perkiraan peserta pelatihan",
     "Jumlah halaman",
     "Warna map penyimpanan"
    ],
    "answer": 0,
    "why": "JRA merupakan acuan retensi yang perlu mengikuti prosedur pengesahan."
   },
   "roles": [
    "arsip"
   ],
   "requiresOffice": true,
   "output": "Register arsip sintetis beserta referensi SOP institusi",
   "cautions": [
    "Contoh bukan pedoman resmi kearsipan atau rekomendasi pemusnahan."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Mengenali hubungan dokumen Word dengan tata kelola arsip dinamis. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-05-04",
   "duration": 35
  },
  {
   "title": "Proyek akhir Word: paket dokumen administrasi",
   "objective": "Menghasilkan berkas terstruktur yang konsisten, mudah diaudit, dan aman dibagikan.",
   "points": [
    "Paket dokumen menyertakan sumber informasi, draf, hasil reviu, dan keluaran final.",
    "Setiap naskah memiliki pembaca dan fungsi yang jelas.",
    "Template dan Styles menjaga keseragaman.",
    "Pengendalian akses dan metadata diterapkan sebelum distribusi.",
    "Penilaian akhir meliputi ketepatan substansi serta teknis format."
   ],
   "steps": [
    "Pilih satu unit serta kasus layanan pada halaman Studi Kasus.",
    "Susun surat, notulen, formulir, dan laporan singkat.",
    "Masukkan tabel atau grafik yang bersumber dari Excel.",
    "Lakukan pemeriksaan revisi, rujukan, serta metadata.",
    "Serahkan paket dengan daftar file dan checklist kualitas."
   ],
   "example": "Paket: Surat.docx; Notulen.docx; Formulir.docx; Laporan.pdf; Log_Revisi.xlsx",
   "exercise": "Buat empat jenis dokumen fiktif terkait satu kegiatan administrasi kampus.",
   "check": {
    "question": "Kriteria paket dokumen final yang baik adalah ...",
    "options": [
     "Konsistensi isi, format, sumber, dan kontrol versi",
     "Banyaknya dekorasi",
     "Tidak ada pemeriksaan",
     "Tanpa metadata maupun penerima"
    ],
    "answer": 0,
    "why": "Dokumen yang siap audit memiliki jejak sumber dan revisi."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Empat dokumen, log reviu, dan checklist distribusi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menghasilkan berkas terstruktur yang konsisten, mudah diaudit, dan aman dibagikan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "W-05-05",
   "duration": 35
  }
 ],
 "powerpoint": [
  {
   "title": "Tujuan paparan dan profil audiens unit kerja",
   "objective": "Menetapkan pesan utama presentasi untuk pimpinan, kolega, atau peserta layanan.",
   "points": [
    "Slide pendukung rapat pimpinan berbeda dari slide panduan petugas layanan.",
    "Tentukan keputusan atau tindakan yang diharapkan setelah paparan.",
    "Gunakan data yang sudah diverifikasi dari unit sumber.",
    "Setiap slide sebaiknya menyampaikan satu pesan utama.",
    "Hindari memuat data pribadi mahasiswa atau pegawai di layar publik."
   ],
   "steps": [
    "Pilih audiens: kepala bagian layanan mahasiswa.",
    "Tentukan tujuan: evaluasi waktu penyelesaian layanan.",
    "Tulis tiga pertanyaan yang perlu dijawab presentasi.",
    "Susun urutan awal, bukti, dan tindak lanjut.",
    "Buat judul deck yang spesifik."
   ],
   "example": "Judul: Evaluasi Ketepatan Layanan Administrasi — Triwulan III 2026",
   "exercise": "Buat outline presentasi tiga slide untuk rapat koordinasi unit.",
   "check": {
    "question": "Presentasi administrasi yang efektif dimulai dengan ...",
    "options": [
     "Tujuan, audiens, dan keputusan yang dituju",
     "Pilihan animasi paling rumit",
     "Menyalin semua tabel mentah",
     "Memilih musik latar"
    ],
    "answer": 0,
    "why": "Tujuan dan audiens menentukan informasi yang perlu ditampilkan."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Outline tiga slide dan profil audiens",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Menetapkan pesan utama presentasi untuk pimpinan, kolega, atau peserta layanan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-01-01",
   "duration": 14
  },
  {
   "title": "Judul slide, hierarki teks, dan ringkasan pesan",
   "objective": "Membedakan judul yang menjelaskan temuan dari judul generik.",
   "points": [
    "Judul temuan mempercepat pembacaan tanpa narasi lisan.",
    "Teks singkat lebih cocok pada slide daripada paragraf laporan.",
    "Gunakan hierarki visual untuk judul, subjudul, dan catatan.",
    "Hindari ukuran huruf kecil dan elemen saling bertumpuk.",
    "Beri konteks waktu dan satuan pada angka."
   ],
   "steps": [
    "Buka simulator PowerPoint dan buat slide baru.",
    "Ganti judul generik menjadi satu temuan terukur.",
    "Tuliskan tiga poin penjelasan maksimum.",
    "Tambahkan satu catatan pembicara terpisah.",
    "Periksa slide pada tampilan presentasi penuh."
   ],
   "example": "Kurang: \"Hasil\" → Lebih baik: \"72 dari 80 layanan diselesaikan pada periode contoh\"",
   "exercise": "Ubah tiga judul generik menjadi judul pesan berbasis data contoh.",
   "check": {
    "question": "Judul slide yang lebih informatif adalah ...",
    "options": [
     "72 dari 80 layanan selesai pada triwulan contoh",
     "Grafik 1",
     "Data",
     "Lain-lain"
    ],
    "answer": 0,
    "why": "Judul insight mengomunikasikan inti temuan dengan konteks."
   },
   "roles": [
    "kemahasiswaan",
    "perencanaan"
   ],
   "requiresOffice": false,
   "output": "Tiga slide dengan judul temuan",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Membedakan judul yang menjelaskan temuan dari judul generik. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-01-02",
   "duration": 14
  },
  {
   "title": "Layout, warna, dan identitas presentasi institusi",
   "objective": "Membuat slide konsisten tanpa meniru atau memalsukan identitas resmi.",
   "points": [
    "Tema mengontrol warna serta jenis huruf secara konsisten.",
    "Ruang kosong membantu fokus pada konten.",
    "Kontras tinggi meningkatkan keterbacaan proyektor.",
    "Gunakan logo resmi hanya dari aset yang disediakan instansi dan dengan izin.",
    "Semua slide perlu margin serta hierarki yang konsisten."
   ],
   "steps": [
    "Pilih salah satu tema simulator.",
    "Atur layout cover dan content secara konsisten.",
    "Gunakan maksimum dua keluarga font.",
    "Pastikan judul tidak terpotong.",
    "Uji pada lebar layar laptop dan layar proyektor."
   ],
   "example": "Cover: Judul, Unit, Periode, Presenter; Konten: Insight, Grafik, Sumber",
   "exercise": "Buat cover dan dua slide isi untuk laporan kinerja unit.",
   "check": {
    "question": "Prinsip desain yang membantu keterbacaan adalah ...",
    "options": [
     "Kontras cukup dan ruang kosong",
     "Menggunakan semua warna sekaligus",
     "Teks 7 pt",
     "Banyak animasi tanpa tujuan"
    ],
    "answer": 0,
    "why": "Kontras dan hierarki visual mendukung fokus pembaca."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Slide cover dan dua layout isi konsisten",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Membuat slide konsisten tanpa meniru atau memalsukan identitas resmi. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-01-03",
   "duration": 14
  },
  {
   "title": "Mengubah tabel menjadi narasi satu slide",
   "objective": "Memilih angka paling penting dari laporan operasional tanpa mengubah makna.",
   "points": [
    "Slide tidak harus memuat seluruh tabel detail.",
    "Pilih dua atau tiga angka inti yang relevan dengan keputusan.",
    "Selalu tampilkan periode, satuan, dan sumber data.",
    "Pisahkan fakta dari interpretasi.",
    "Jangan menampilkan persentase tanpa pembilang dan penyebut bila rawan salah tafsir."
   ],
   "steps": [
    "Gunakan data contoh 80 permohonan, 72 selesai, 8 diproses.",
    "Hitung proporsi selesai =72/80×100%.",
    "Buat satu slide dengan angka kunci 90%.",
    "Tambahkan satu pesan dan satu tindak lanjut.",
    "Periksa kembali bahwa 72+8=80."
   ],
   "example": "90% layanan selesai · 8 permohonan dalam proses (data simulasi)",
   "exercise": "Buat satu slide ringkasan jumlah layanan dan pesan tindak lanjut.",
   "check": {
    "question": "Persentase 72 dari 80 adalah ...",
    "options": [
     "90%",
     "72%",
     "8%",
     "110%"
    ],
    "answer": 0,
    "why": "72 dibagi 80 dikalikan 100% menghasilkan 90%."
   },
   "roles": [
    "akademik",
    "kemahasiswaan"
   ],
   "requiresOffice": false,
   "output": "Slide KPI satu halaman dengan sumber",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Memilih angka paling penting dari laporan operasional tanpa mengubah makna. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-01-04",
   "duration": 14
  },
  {
   "title": "Menyimpan deck, PDF, dan menjalankan slideshow",
   "objective": "Melakukan pemeriksaan presentasi sebelum rapat atau pelatihan.",
   "points": [
    "PPTX adalah berkas PowerPoint yang dapat diedit.",
    "PDF cocok untuk salinan distribusi noninteraktif.",
    "F5 umumnya memulai slideshow dari slide pertama di PowerPoint desktop Windows.",
    "Esc menghentikan tayangan.",
    "Media, font, dan format perlu diuji pada perangkat presentasi."
   ],
   "steps": [
    "Simpan deck sebagai Paparan_Unit_v01.pptx.",
    "Coba Slide Show → From Beginning.",
    "Gunakan tombol panah untuk berpindah slide.",
    "Akhiri slideshow dan periksa urutan slide.",
    "Ekspor PDF dan cek setiap halaman."
   ],
   "example": "File final: 2026_PaparanKinerja_Unit_v02.pptx + salinan PDF",
   "exercise": "Jalankan tayangan tiga slide dalam simulator lalu jelaskan tombol navigasinya.",
   "check": {
    "question": "Shortcut yang umum dipakai memulai tayangan dari awal adalah ...",
    "options": [
     "F5",
     "Ctrl+S",
     "Ctrl+H",
     "Alt+="
    ],
    "answer": 0,
    "why": "F5 memulai slideshow pada PowerPoint desktop Windows."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Deck yang dapat dipresentasikan dan diperiksa",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Basic",
   "min": 14,
   "lead": "Konteks kampus: Melakukan pemeriksaan presentasi sebelum rapat atau pelatihan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-01-05",
   "duration": 14
  },
  {
   "title": "Paparan alur layanan mahasiswa",
   "objective": "Membuat alur layanan yang mudah dipahami pemohon dan petugas.",
   "points": [
    "Gunakan empat tahap maksimum pada satu visual utama.",
    "Bedakan langkah pemohon dari langkah verifikator.",
    "Tampilkan dokumen yang harus dibawa tanpa memuat data pribadi.",
    "Berikan titik bantuan bila permohonan belum lengkap.",
    "Cantumkan bahwa diagram adalah latihan dan perlu disesuaikan SOP."
   ],
   "steps": [
    "Buat slide 1: tujuan layanan.",
    "Slide 2: alur Permohonan→Verifikasi→Proses→Selesai.",
    "Slide 3: dokumen dan kontak unit.",
    "Gunakan ikon sederhana serta teks ringkas.",
    "Latih menjelaskan alur dalam dua menit."
   ],
   "example": "Alur layanan contoh: Isi formulir → Verifikasi petugas → Proses dokumen → Penyerahan",
   "exercise": "Buat tiga slide untuk orientasi petugas layanan mahasiswa.",
   "check": {
    "question": "Diagram alur yang membantu staf layanan harus ...",
    "options": [
     "Menampilkan tahap, aktor, dan hasil",
     "Memuat semua paragraf SOP",
     "Menyembunyikan langkah verifikasi",
     "Menggunakan simbol acak"
    ],
    "answer": 0,
    "why": "Tahap, tanggung jawab, dan hasil mendukung konsistensi pelayanan."
   },
   "roles": [
    "kemahasiswaan",
    "akademik"
   ],
   "requiresOffice": false,
   "output": "Tiga slide alur layanan dan catatan pembicara",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Membuat alur layanan yang mudah dipahami pemohon dan petugas. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-02-01",
   "duration": 18
  },
  {
   "title": "Paparan anggaran: pagu, realisasi, dan deviasi",
   "objective": "Menyajikan perbandingan rupiah secara jelas untuk rapat monitoring.",
   "points": [
    "Pagu adalah batas/anggaran yang digunakan pada kasus simulasi.",
    "Realisasi adalah nilai yang sudah tercatat dalam data sumber.",
    "Selisih = pagu − realisasi.",
    "Tampilkan satuan rupiah dan periode di setiap visual.",
    "Jangan menyimpulkan efektivitas hanya dari persentase serapan."
   ],
   "steps": [
    "Siapkan data Pagu Rp50 juta, Realisasi Rp42 juta.",
    "Hitung selisih Rp8 juta serta serapan 84%.",
    "Buat dua kolom perbandingan.",
    "Tambahkan catatan konteks dan kebutuhan tindak lanjut.",
    "Pastikan seluruh angka sesuai data sumber."
   ],
   "example": "Pagu Rp50.000.000; Realisasi Rp42.000.000; Serapan 84%; Selisih Rp8.000.000",
   "exercise": "Buat satu slide ringkasan anggaran fiktif untuk rapat unit.",
   "check": {
    "question": "Jika pagu 50 juta dan realisasi 42 juta, serapannya ...",
    "options": [
     "84%",
     "8%",
     "92%",
     "119%"
    ],
    "answer": 0,
    "why": "42/50×100%=84%."
   },
   "roles": [
    "keuangan",
    "perencanaan"
   ],
   "requiresOffice": false,
   "output": "Slide perbandingan anggaran dan interpretasi berimbang",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menyajikan perbandingan rupiah secara jelas untuk rapat monitoring. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-02-02",
   "duration": 18
  },
  {
   "title": "Timeline kegiatan dan milestone perencanaan",
   "objective": "Menyusun urutan kegiatan, penanggung jawab, dan tanggal target.",
   "points": [
    "Timeline memvisualisasikan urutan, bukan menggantikan rincian jadwal.",
    "Milestone menandai keluaran terukur.",
    "Tenggat perlu realistis dan sesuai persetujuan pihak terkait.",
    "Dependensi antaraktivitas perlu ditampilkan secara jelas.",
    "Status selesai harus didukung bukti, bukan asumsi."
   ],
   "steps": [
    "Pilih kasus persiapan pelatihan administrasi.",
    "Tentukan tahapan Persiapan, Pelaksanaan, Evaluasi.",
    "Tambahkan tanggal contoh dan PIC per tahap.",
    "Tampilkan tiga milestone pada slide.",
    "Periksa dependensi dan rentang tanggal."
   ],
   "example": "19 Nov: finalisasi materi; 23 Nov: pelatihan; 27 Nov: evaluasi (contoh)",
   "exercise": "Buat slide timeline satu bulan beserta tiga milestone.",
   "check": {
    "question": "Milestone adalah ...",
    "options": [
     "Titik capaian terukur dalam jadwal",
     "Paragraf catatan tanpa waktu",
     "Gambar dekoratif",
     "Rumus rata-rata"
    ],
    "answer": 0,
    "why": "Milestone menandai hasil atau titik penting suatu rencana."
   },
   "roles": [
    "perencanaan",
    "kepegawaian"
   ],
   "requiresOffice": false,
   "output": "Slide timeline beserta PIC dan keluaran",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Menyusun urutan kegiatan, penanggung jawab, dan tanggal target. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-02-03",
   "duration": 18
  },
  {
   "title": "Tabel, gambar, dan sumber visual",
   "objective": "Memakai tabel atau gambar untuk membantu penjelasan tanpa mengaburkan sumber.",
   "points": [
    "Tabel cocok untuk perbandingan terbatas dengan sedikit baris.",
    "Gambar perlu relevan dan berizin.",
    "Diagram sebaiknya memiliki label serta urutan baca.",
    "Cantumkan sumber atau keterangan dataset contoh.",
    "Periksa apakah visual masih terbaca di proyektor."
   ],
   "steps": [
    "Pilih satu tabel dengan tiga baris indikator.",
    "Tulis judul temuan di atas tabel.",
    "Berikan sorotan pada satu angka kunci.",
    "Tambahkan sumber data dan periode.",
    "Buat versi grafik jika tabel terlalu padat."
   ],
   "example": "Indikator: Permohonan masuk 80; selesai 72; tertunda 8",
   "exercise": "Buat dua slide: satu tabel dan satu visualisasi alternatif.",
   "check": {
    "question": "Penggunaan sumber data pada slide membantu ...",
    "options": [
     "Verifikasi dan keterlacakan angka",
     "Menambah animasi",
     "Menghilangkan kebutuhan uji",
     "Mengubah DOCX menjadi XLSX"
    ],
    "answer": 0,
    "why": "Sumber memungkinkan audiens memeriksa dasar klaim."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Dua slide visual data dan catatan sumber",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Memakai tabel atau gambar untuk membantu penjelasan tanpa mengaburkan sumber. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-02-04",
   "duration": 18
  },
  {
   "title": "Speaker Notes dan latihan pemaparan",
   "objective": "Memakai catatan pembicara agar slide ringkas tetapi penjelasan lengkap.",
   "points": [
    "Notes tidak harus tampil pada layar audiens.",
    "Catatan memuat definisi indikator atau transisi narasi.",
    "Presenter View dapat menampilkan notes jika perangkat mendukung.",
    "Durasi latihan membantu menekan paparan melebar.",
    "Siapkan jawaban atas satu pertanyaan kritis per slide."
   ],
   "steps": [
    "Buka salah satu slide di simulator.",
    "Isi notes dengan penjelasan satu paragraf.",
    "Jalankan mode Presentasikan dan gunakan tombol panah.",
    "Latih narasi selama satu menit per slide.",
    "Catat satu titik yang perlu disederhanakan."
   ],
   "example": "Notes: \"Capaian 90% dihitung dari 72 layanan selesai dibagi 80 masuk.\"",
   "exercise": "Buat notes untuk tiga slide dan lakukan simulasi presentasi lima menit.",
   "check": {
    "question": "Manfaat utama Speaker Notes adalah ...",
    "options": [
     "Menyimpan informasi bantu untuk pembicara",
     "Mengubah tabel menjadi rumus",
     "Menghapus presentasi",
     "Membuat daftar isi Word"
    ],
    "answer": 0,
    "why": "Notes mendukung narasi tanpa memenuhi layar dengan teks."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": false,
   "output": "Catatan pembicara dan rencana durasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Beginner",
   "min": 18,
   "lead": "Konteks kampus: Memakai catatan pembicara agar slide ringkas tetapi penjelasan lengkap. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-02-05",
   "duration": 18
  },
  {
   "title": "Storyline masalah–bukti–opsi–rekomendasi",
   "objective": "Mengubah data operasional menjadi presentasi yang mendukung keputusan.",
   "points": [
    "Mulailah dari isu terukur dan mengapa isu itu penting.",
    "Gunakan bukti dari sumber yang dapat diverifikasi.",
    "Pisahkan analisis penyebab dari dugaan yang belum diuji.",
    "Paparkan beberapa opsi dengan konsekuensi yang seimbang.",
    "Akhiri dengan keputusan atau tindak lanjut yang jelas."
   ],
   "steps": [
    "Tulis isu \"8 dari 80 layanan belum selesai\".",
    "Sajikan bukti jumlah per tahap proses.",
    "Tentukan dua penjelasan yang perlu diverifikasi.",
    "Bandingkan dua opsi perbaikan dengan kebutuhan sumber daya.",
    "Buat slide keputusan dan PIC tindak lanjut."
   ],
   "example": "Urutan 5 slide: Isu → Fakta → Analisis → Opsi → Rekomendasi",
   "exercise": "Susun paparan lima slide dari satu kasus keterlambatan layanan.",
   "check": {
    "question": "Perbedaan fakta dan asumsi dalam paparan ialah ...",
    "options": [
     "Fakta didukung data, asumsi perlu diuji",
     "Semua asumsi adalah fakta",
     "Fakta tidak perlu sumber",
     "Asumsi tidak boleh dicatat"
    ],
    "answer": 0,
    "why": "Pemilahan fakta dan hipotesis mencegah kesimpulan prematur."
   },
   "roles": [
    "perencanaan",
    "kemahasiswaan"
   ],
   "requiresOffice": false,
   "output": "Deck lima slide dengan rekomendasi yang bisa ditindaklanjuti",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Mengubah data operasional menjadi presentasi yang mendukung keputusan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-03-01",
   "duration": 22
  },
  {
   "title": "Membuat grafik dari Excel dengan satuan jelas",
   "objective": "Menyisipkan grafik dengan nilai dan sumber yang konsisten terhadap workbook.",
   "points": [
    "Grafik garis cocok untuk tren waktu; kolom untuk perbandingan kategori.",
    "Tulis judul berupa temuan, bukan \"Grafik\".",
    "Pastikan rentang dan angka sesuai workbook sumber.",
    "Tampilkan satuan rupiah/orang dan periode.",
    "Beri catatan bila pembanding mengalami perubahan definisi."
   ],
   "steps": [
    "Buat grafik bulanan dalam Excel sintetis.",
    "Salin ke PowerPoint memakai opsi paste yang sesuai.",
    "Atur judul, sumbu, dan label.",
    "Bandingkan satu titik data dengan Excel.",
    "Sertakan sumber dan tanggal pengambilan data."
   ],
   "example": "Jan 55, Feb 62, Mar 73 layanan selesai (data simulasi)",
   "exercise": "Buat slide tren tiga bulan dengan interpretasi dan verifikasi satu titik.",
   "check": {
    "question": "Untuk menggambarkan perubahan dari Januari ke Maret, grafik yang lazim adalah ...",
    "options": [
     "Grafik garis",
     "Grafik dekoratif tanpa angka",
     "Diagram pohon file",
     "Mail Merge"
    ],
    "answer": 0,
    "why": "Grafik garis lazim untuk memperlihatkan deret waktu."
   },
   "roles": [
    "perencanaan",
    "akademik",
    "kepegawaian"
   ],
   "requiresOffice": true,
   "output": "Slide tren dengan sumber dan angka terverifikasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Menyisipkan grafik dengan nilai dan sumber yang konsisten terhadap workbook. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-03-02",
   "duration": 22
  },
  {
   "title": "Slide Master dan layout presentasi unit",
   "objective": "Membuat template unit yang konsisten sepanjang rapat berkala.",
   "points": [
    "Slide Master mengatur elemen bersama di berbagai layout.",
    "Sediakan cover, pemisah bab, slide data, dan penutup.",
    "Footer dapat memuat unit serta periode, tidak perlu data rahasia.",
    "Gunakan warna identitas yang disahkan institusi.",
    "Periksa apakah layout bekerja untuk judul panjang dan tabel."
   ],
   "steps": [
    "Buka View → Slide Master pada PowerPoint desktop.",
    "Rancang empat layout dengan margin konsisten.",
    "Atur footer Unit, Periode, dan nomor slide.",
    "Uji pemakaian pada lima slide berbeda.",
    "Simpan sebagai template yang dapat diperbarui."
   ],
   "example": "Layout: Cover / KPI / Dua Kolom / Tindak Lanjut",
   "exercise": "Buat master presentasi laporan bulanan dengan empat layout.",
   "check": {
    "question": "Fungsi Slide Master adalah ...",
    "options": [
     "Mengatur gaya konsisten berbagai slide",
     "Mempercepat SUMIF",
     "Menyunting PDF hasil scan",
     "Menggabung baris Excel"
    ],
    "answer": 0,
    "why": "Slide Master menjaga elemen desain bersama di banyak slide."
   },
   "roles": [
    "perencanaan",
    "semua"
   ],
   "requiresOffice": true,
   "output": "Template PowerPoint empat layout teruji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Membuat template unit yang konsisten sepanjang rapat berkala. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-03-03",
   "duration": 22
  },
  {
   "title": "Navigasi hyperlink dan menu paparan",
   "objective": "Membuat navigasi pada deck yang memiliki bagian laporan berbeda.",
   "points": [
    "Hyperlink dapat membuka slide tertentu di deck.",
    "Menu perlu memiliki tombol kembali yang jelas.",
    "Tautan eksternal berisiko tidak tersedia offline.",
    "Latih semua cabang navigasi sebelum acara.",
    "Sediakan urutan linear sebagai cadangan."
   ],
   "steps": [
    "Buat slide menu Unit Keuangan / Layanan / Arsip.",
    "Hubungkan tiap tombol ke slide bagiannya.",
    "Tambahkan tombol Kembali ke Menu.",
    "Jalankan Slide Show dan uji semua tombol.",
    "Catat jika ada tautan rusak atau tidak sesuai."
   ],
   "example": "Menu utama → Subbagian A/B/C → Kembali ke menu utama",
   "exercise": "Buat prototype deck bercabang tiga bagian dan uji semuanya.",
   "check": {
    "question": "Komponen wajib untuk navigasi bercabang adalah ...",
    "options": [
     "Tombol kembali dan pengujian setiap jalur",
     "Tautan acak tanpa label",
     "Sumber yang tidak dapat diakses",
     "Semua slide disembunyikan"
    ],
    "answer": 0,
    "why": "Tombol kembali menjaga jalur presentasi dapat digunakan."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Deck navigasi dengan empat tautan teruji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Membuat navigasi pada deck yang memiliki bagian laporan berbeda. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-03-04",
   "duration": 22
  },
  {
   "title": "Review konten dan latihan menjawab pertanyaan",
   "objective": "Menyiapkan paparan yang dapat dipertanggungjawabkan ketika pimpinan meminta rincian.",
   "points": [
    "Sumber angka harus bisa ditunjukkan saat diminta.",
    "Pisahkan angka yang sudah final dari estimasi.",
    "Speaker Notes memuat definisi dan catatan keterbatasan.",
    "Gunakan appendix untuk rincian yang tidak cocok di slide utama.",
    "Tinjau paparan bersama rekan sebelum digunakan secara operasional."
   ],
   "steps": [
    "Pilih satu deck ringkasan unit.",
    "Buat daftar lima pertanyaan yang mungkin muncul.",
    "Tambahkan tiga lampiran data sumber atau definisi.",
    "Periksa istilah serta angka yang tidak konsisten.",
    "Latih jawaban berbasis data tanpa menebak."
   ],
   "example": "Pertanyaan: Mengapa capaian 90%, apa pembaginya, dan siapa memvalidasi?",
   "exercise": "Buat daftar Q&A untuk paparan kinerja layanan dengan bukti pendukung.",
   "check": {
    "question": "Jika angka diminta tetapi belum terverifikasi, sebaiknya ...",
    "options": [
     "Menyatakan perlu verifikasi sumber dan memberi tindak lanjut",
     "Menciptakan angka perkiraan tanpa label",
     "Menghapus semua data",
     "Mengalihkan jawaban ke dekorasi"
    ],
    "answer": 0,
    "why": "Ketepatan dan keterbatasan data perlu diungkapkan."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Deck dan daftar pertanyaan–jawaban berbukti",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Intermediate",
   "min": 22,
   "lead": "Konteks kampus: Menyiapkan paparan yang dapat dipertanggungjawabkan ketika pimpinan meminta rincian. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-03-05",
   "duration": 22
  },
  {
   "title": "Dashboard satu slide untuk rapat monitoring",
   "objective": "Menyajikan target, realisasi, tren, dan risiko sebagai satu halaman eksekutif.",
   "points": [
    "Pilih maksimal beberapa KPI utama sesuai keputusan rapat.",
    "Tunjukkan angka target, realisasi, dan selisih dengan satuan.",
    "Ringkas tren waktu dan anomali secara visual.",
    "Hindari indikator yang definisinya berubah tanpa catatan.",
    "Berikan ruang khusus tindak lanjut."
   ],
   "steps": [
    "Pilih tiga KPI pada dataset latihan.",
    "Rancang layout 3 kartu dan satu grafik tren.",
    "Cantumkan periode serta keterangan sumber.",
    "Tambahkan satu catatan risiko dan usulan tindak lanjut.",
    "Lakukan review keterbacaan pada tampilan proyektor."
   ],
   "example": "KPI 1: 90% layanan selesai; KPI 2: 8 tertunda; KPI 3: 2 hari median proses (simulasi)",
   "exercise": "Buat dashboard satu slide dan jelaskan dua keputusan yang didukungnya.",
   "check": {
    "question": "Dashboard satu slide sebaiknya berisi ...",
    "options": [
     "Indikator inti, satuan, tren, dan tindak lanjut",
     "Seluruh tabel mentah beratus baris",
     "Animasi sepanjang lima menit",
     "Gambar besar tanpa angka"
    ],
    "answer": 0,
    "why": "Ringkasan eksekutif mengutamakan sinyal yang relevan terhadap keputusan."
   },
   "roles": [
    "perencanaan",
    "kemahasiswaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Satu slide dashboard teruji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menyajikan target, realisasi, tren, dan risiko sebagai satu halaman eksekutif. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-04-01",
   "duration": 28
  },
  {
   "title": "Membaca variansi, tren, dan penyebab yang diuji",
   "objective": "Menyampaikan deviasi kinerja secara jujur tanpa menyimpulkan sebab secara prematur.",
   "points": [
    "Variansi menunjukkan perbedaan dua nilai yang terdefinisi.",
    "Tren membutuhkan data beberapa periode yang sebanding.",
    "Satu anomali tidak membuktikan hubungan sebab-akibat.",
    "Tandai perubahan proses pencatatan agar perbandingan adil.",
    "Pisahkan pernyataan observasi, hipotesis, dan tindakan."
   ],
   "steps": [
    "Buat grafik capaian layanan selama empat bulan.",
    "Tandai bulan yang turun secara signifikan.",
    "Tulis observasi berdasarkan angka.",
    "Cantumkan kemungkinan faktor yang perlu diuji.",
    "Ajukan data tambahan untuk memeriksa hipotesis."
   ],
   "example": "April 82%, Mei 88%, Juni 77%, Juli 90% (contoh)",
   "exercise": "Buat satu slide tren dengan anotasi dan dua pertanyaan investigasi.",
   "check": {
    "question": "Dari satu penurunan kinerja, kesimpulan yang benar adalah ...",
    "options": [
     "Ada perubahan teramati, penyebab perlu diuji",
     "Pasti karena staf tidak disiplin",
     "Data tidak perlu dibandingkan",
     "Semua unit harus dipotong anggarannya"
    ],
    "answer": 0,
    "why": "Deskripsi tren tidak dengan sendirinya membuktikan kausalitas."
   },
   "roles": [
    "keuangan",
    "perencanaan",
    "kepegawaian"
   ],
   "requiresOffice": true,
   "output": "Slide analisis tren dan daftar verifikasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menyampaikan deviasi kinerja secara jujur tanpa menyimpulkan sebab secara prematur. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-04-02",
   "duration": 28
  },
  {
   "title": "Objek linked vs embedded dari Excel",
   "objective": "Memilih cara menyalin grafik atau tabel agar pembaruan dan distribusi terkontrol.",
   "points": [
    "Linked object mengikuti sumber yang tersedia dan relatif memudahkan pembaruan.",
    "Embedded object menyimpan salinan dalam presentasi.",
    "Pemindahan lokasi sumber dapat memutus tautan.",
    "Ekspor PDF tidak mempertahankan pengeditan objek sumber.",
    "Pilih teknik berdasarkan kebutuhan distribusi dan kontrol informasi."
   ],
   "steps": [
    "Siapkan grafik dari Excel latihan.",
    "Tempel salinan linked dan embedded pada dua slide.",
    "Ubah satu angka sumber lalu tinjau efek pada kedua slide.",
    "Simulasikan perpindahan folder sumber.",
    "Buat catatan kapan memilih linked atau embedded."
   ],
   "example": "Linked: dipakai internal saat sumber terjaga; Embedded: salinan untuk distribusi terkontrol",
   "exercise": "Tulis matriks keputusan penggunaan linked/embedded untuk laporan bulanan.",
   "check": {
    "question": "Risiko khas linked chart ialah ...",
    "options": [
     "Tautan putus ketika berkas sumber tidak tersedia",
     "Grafik otomatis berubah menjadi audio",
     "Slide kehilangan judul tanpa sebab",
     "Data selalu tersertifikasi"
    ],
    "answer": 0,
    "why": "Tautan membutuhkan sumber yang masih dapat diakses."
   },
   "roles": [
    "perencanaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Perbandingan dua metode beserta rekomendasi",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Memilih cara menyalin grafik atau tabel agar pembaruan dan distribusi terkontrol. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-04-03",
   "duration": 28
  },
  {
   "title": "Custom Show untuk rapat lintas bagian",
   "objective": "Membuat satu deck dengan jalur khusus pimpinan dan jalur teknis.",
   "points": [
    "Custom Slide Show dapat memilih subset slide untuk audiens tertentu.",
    "Jalur eksekutif memerlukan ringkasan keputusan lebih singkat.",
    "Jalur teknis memerlukan rincian perhitungan serta prosedur.",
    "Menu navigasi dan halaman kembali perlu diuji.",
    "Catatan pembicara dapat berbeda menurut audiens."
   ],
   "steps": [
    "Bangun deck delapan slide tentang kinerja unit.",
    "Tandai empat slide untuk pimpinan dan empat untuk teknis.",
    "Buat dua Custom Shows di PowerPoint desktop.",
    "Uji alur dari awal sampai akhir.",
    "Catat slide bersama dan lampiran yang dipakai."
   ],
   "example": "Jalur pimpinan: 1,2,5,8; jalur teknis: 1,3,4,6,7,8",
   "exercise": "Rancang dua jalur presentasi dari satu data kinerja administrasi.",
   "check": {
    "question": "Kegunaan Custom Show yang tepat adalah ...",
    "options": [
     "Menyajikan subset slide menurut kebutuhan audiens",
     "Mengoreksi jurnal kas otomatis",
     "Menghapus semua data sumber",
     "Membuat surat dinas"
    ],
    "answer": 0,
    "why": "Custom Show memisahkan alur berdasarkan tujuan audiens."
   },
   "roles": [
    "perencanaan",
    "semua"
   ],
   "requiresOffice": true,
   "output": "Deck dengan dua jalur dan catatan uji",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Membuat satu deck dengan jalur khusus pimpinan dan jalur teknis. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-04-04",
   "duration": 28
  },
  {
   "title": "Audit aksesibilitas, notes, dan cadangan offline",
   "objective": "Menyiapkan presentasi operasional yang andal pada berbagai perangkat.",
   "points": [
    "Kontras dan ukuran teks harus terbaca dari belakang ruangan.",
    "Alt text dan reading order membantu aksesibilitas.",
    "Video dan tautan harus diuji tanpa internet bila diperlukan.",
    "Sediakan salinan PDF ketika fitur interaktif tidak tersedia.",
    "Jangan menayangkan daftar personal yang tidak sesuai kewenangan."
   ],
   "steps": [
    "Jalankan Accessibility Checker pada PowerPoint desktop.",
    "Periksa ukuran huruf, alt text, dan sumber gambar.",
    "Tes semua tautan serta media offline.",
    "Simpan PPTX dan PDF ke media cadangan sesuai kebijakan.",
    "Latih pembukaan presentasi pada komputer lain."
   ],
   "example": "Checklist: font, media, tautan, notes, data pribadi, PDF cadangan",
   "exercise": "Buat checklist prapresentasi 10 butir untuk kegiatan kampus.",
   "check": {
    "question": "Cadangan yang tepat untuk presentasi penting adalah ...",
    "options": [
     "Salinan file yang diuji, termasuk PDF noninteraktif",
     "Mengandalkan satu link online",
     "Menyalin data sensitif ke media umum",
     "Melewati uji proyektor"
    ],
    "answer": 0,
    "why": "Pemeriksaan offline mencegah kegagalan karena font, media, dan konektivitas."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Checklist uji dan dua format keluaran",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Advanced",
   "min": 28,
   "lead": "Konteks kampus: Menyiapkan presentasi operasional yang andal pada berbagai perangkat. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-04-05",
   "duration": 28
  },
  {
   "title": "Executive briefing hasil kinerja untuk pimpinan",
   "objective": "Menyusun paparan ringkas dengan angka, risiko, opsi, dan kebutuhan keputusan.",
   "points": [
    "Pimpinan membutuhkan keputusan, bukan daftar semua kegiatan.",
    "Gunakan headline berupa temuan yang didukung data.",
    "Tunjukkan perbandingan terhadap target serta sumbernya.",
    "Tampilkan risiko, opsi, trade-off, dan rekomendasi.",
    "Akhiri dengan tindakan, pemilik tugas, dan waktu yang diusulkan."
   ],
   "steps": [
    "Pilih kasus layanan mahasiswa atau keuangan.",
    "Tulis executive summary satu slide.",
    "Sajikan dua grafik yang telah diverifikasi.",
    "Buat matriks dua opsi tindak lanjut.",
    "Tutup dengan keputusan yang diminta dan PIC."
   ],
   "example": "Keputusan diminta: prioritas perbaikan alur verifikasi berkas pada periode berikut",
   "exercise": "Buat briefing enam slide untuk pimpinan yang memakai data sintetis.",
   "check": {
    "question": "Elemen penutup executive briefing yang kuat adalah ...",
    "options": [
     "Keputusan yang diminta, PIC, dan tindak lanjut",
     "Hanya ucapan terima kasih",
     "Daftar efek animasi",
     "Semua data mentah"
    ],
    "answer": 0,
    "why": "Paparan keputusan harus jelas menyatakan tindak lanjut yang diperlukan."
   },
   "roles": [
    "perencanaan",
    "kemahasiswaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Briefing enam slide dan ringkasan keputusan",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menyusun paparan ringkas dengan angka, risiko, opsi, dan kebutuhan keputusan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-05-01",
   "duration": 35
  },
  {
   "title": "Paparan opsi perencanaan dan anggaran lintas unit",
   "objective": "Membuat analisis skenario yang transparan untuk forum perencanaan.",
   "points": [
    "Bandingkan skenario dengan asumsi dan periode yang sama.",
    "Sajikan biaya, manfaat, risiko, dan keterbatasan tiap opsi.",
    "Jangan menukar nilai asumsi menjadi klaim realisasi.",
    "Grafik perlu baseline dan satuan.",
    "Rekomendasi didukung parameter yang dapat diperiksa."
   ],
   "steps": [
    "Siapkan tiga skenario pagu operasional pada Excel.",
    "Hitung biaya dan sasaran tiap skenario.",
    "Buat grafik perbandingan sederhana.",
    "Tulis rekomendasi dan batas asumsinya.",
    "Siapkan appendix rumus serta sumber angka."
   ],
   "example": "Skenario Dasar 50 juta; Efisiensi 45 juta; Ekspansi 58 juta (simulasi)",
   "exercise": "Susun presentasi 8 slide rapat perencanaan dengan matriks opsi.",
   "check": {
    "question": "Analisis skenario yang valid memerlukan ...",
    "options": [
     "Asumsi dan pembanding yang sepadan",
     "Angka tebakan tanpa label",
     "Hanya satu opsi tanpa alternatif",
     "Menghilangkan risiko"
    ],
    "answer": 0,
    "why": "Kesetaraan asumsi memungkinkan perbandingan yang relevan."
   },
   "roles": [
    "perencanaan",
    "keuangan"
   ],
   "requiresOffice": true,
   "output": "Paparan skenario dan lampiran perhitungan",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Membuat analisis skenario yang transparan untuk forum perencanaan. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-05-02",
   "duration": 35
  },
  {
   "title": "Deck panduan SOP layanan dan kearsipan",
   "objective": "Mentransformasikan SOP terstruktur menjadi media penguatan keterampilan pegawai.",
   "points": [
    "Slide prosedur tidak boleh bertentangan dengan naskah SOP disahkan.",
    "Tampilkan peran, dokumen, tahap, dan pengecualian.",
    "Berikan skenario tidak lengkap untuk latihan pengambilan keputusan.",
    "Pisahkan instruksi kerja dari materi diskusi.",
    "Sertakan versi SOP dan lokasi rujukan resmi."
   ],
   "steps": [
    "Gunakan draf SOP pelayanan administrasi sintetis.",
    "Pecah alur ke empat slide tahap proses.",
    "Buat satu slide kesalahan umum dan kontrol.",
    "Susun dua pertanyaan latihan untuk peserta.",
    "Cocokkan istilah dengan naskah SOP sumber."
   ],
   "example": "Skenario: berkas pemohon tidak lengkap → status Perlu Perbaikan → kembali setelah verifikasi",
   "exercise": "Buat lima slide orientasi SOP yang menyertakan dua skenario kasus.",
   "check": {
    "question": "Yang menjadi rujukan utama prosedur kerja ialah ...",
    "options": [
     "SOP institusi yang telah disahkan",
     "Slide latihan yang belum ditinjau",
     "Template acak di internet",
     "Perkiraan personal"
    ],
    "answer": 0,
    "why": "Slide adalah media pengajaran, bukan pengganti SOP resmi."
   },
   "roles": [
    "arsip",
    "akademik",
    "kemahasiswaan"
   ],
   "requiresOffice": true,
   "output": "Deck sosialisasi SOP dan daftar verifikasi sumber",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Mentransformasikan SOP terstruktur menjadi media penguatan keterampilan pegawai. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-05-03",
   "duration": 35
  },
  {
   "title": "Design system dan template paparan institusi",
   "objective": "Membuat aset visual reusable untuk semua unit tanpa mengorbankan keseragaman.",
   "points": [
    "Design system menetapkan tipografi, palet, grid, ruang, dan komponen.",
    "Template wajib mencakup layout ringkasan, data, alur, dan penutup.",
    "Kontras serta ukuran konten diuji pada perangkat sasaran.",
    "Logo serta identitas visual mengikuti pedoman yang berlaku.",
    "Buat halaman panduan pemakaian dan contoh kesalahan."
   ],
   "steps": [
    "Tetapkan palet tiga warna dari pedoman institusi yang tersedia.",
    "Rancang master lima layout.",
    "Buat contoh KPI, tabel, diagram alur, dan timeline.",
    "Tulis instruksi penggunaan satu halaman.",
    "Uji file pada laptop dan proyektor berbeda."
   ],
   "example": "Layout: Cover, Section, KPI, Process, Decision",
   "exercise": "Buat template tim administrasi yang dapat digunakan untuk berbagai kegiatan.",
   "check": {
    "question": "Keunggulan design system adalah ...",
    "options": [
     "Konsistensi dan efisiensi produksi slide",
     "Memastikan semua data selalu benar",
     "Menggantikan evaluasi konten",
     "Menghapus kebutuhan review"
    ],
    "answer": 0,
    "why": "Design system menstandarkan penyajian, bukan memvalidasi substansi."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Template lima layout dan panduan penggunaan",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Membuat aset visual reusable untuk semua unit tanpa mengorbankan keseragaman. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-05-04",
   "duration": 35
  },
  {
   "title": "Proyek akhir PowerPoint: paparan capaian unit",
   "objective": "Menghasilkan deck yang siap ditinjau dengan narasi, sumber data, dan tindak lanjut.",
   "points": [
    "Paket akhir menggunakan data fiktif terverifikasi dari latihan Excel.",
    "Narasi harus memisahkan fakta dan usulan.",
    "Slide memiliki sumber serta periode yang jelas.",
    "Alur presentasi diuji dari awal sampai akhir.",
    "Ada cadangan PDF, notes, dan log pemeriksaan."
   ],
   "steps": [
    "Pilih satu studi kasus unit kerja dari OfficeLab.",
    "Susun outline 8–10 slide beserta audiens.",
    "Buat grafik dan tabel dari rekap Excel sintetis.",
    "Tambahkan keputusan yang diusulkan dan tahapan implementasi.",
    "Uji tampilan, navigasi, dan backup offline."
   ],
   "example": "Keluaran: PaparanKinerja.pptx, PDF ringkasan, sumber data, checklist uji",
   "exercise": "Buat presentasi kinerja unit 8–10 slide dan sampaikan simulasi lima menit.",
   "check": {
    "question": "Ciri paparan akhir yang siap diuji adalah ...",
    "options": [
     "Angka konsisten, alur jelas, sumber dan tindak lanjut ada",
     "Menggunakan semua efek animasi",
     "Tidak mencantumkan periode",
     "Berisi semua tabel mentah"
    ],
    "answer": 0,
    "why": "Keandalan dan akuntabilitas penting dalam laporan kinerja unit."
   },
   "roles": [
    "semua"
   ],
   "requiresOffice": true,
   "output": "Deck akhir, sumber, dan checklist pengujian",
   "cautions": [
    "Jangan menganggap latihan ini sebagai format atau SOP resmi institusi.",
    "Periksa kecocokan fitur dengan versi Office yang digunakan peserta."
   ],
   "level": "Expert",
   "min": 35,
   "lead": "Konteks kampus: Menghasilkan deck yang siap ditinjau dengan narasi, sumber data, dan tindak lanjut. Contoh dan berkas dalam modul ini bersifat sintetis untuk pelatihan.",
   "tip": "Simpan bukti latihan dengan nama file, versi, dan tanggal; minta pemeriksaan fasilitator sebelum menerapkan hasil ke proses resmi.",
   "rubric": [
    "Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.",
    "Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.",
    "Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.",
    "Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi."
   ],
   "code": "P-05-05",
   "duration": 35
  }
 ]
};
for(const [id,items] of Object.entries(CURRICULUM)){ if(items.length!==25)throw Error('Expected 25 modules: '+id); OFFICE[id].lessons=items; OFFICE[id].quiz=items.map(m=>[m.check.question,m.check.options,m.check.answer,m.check.why]);}
