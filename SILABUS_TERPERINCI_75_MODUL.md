# Silabus Terperinci OfficeLab v2.1 — 75 Modul

**Sasaran:** Tenaga kependidikan/pegawai administrasi perguruan tinggi: Keuangan, Kemahasiswaan, Perencanaan, Persuratan & Arsip, Kepegawaian, Layanan Akademik, Pengadaan & Aset.

**Format tiap modul:** tujuan → konsep → langkah praktik → contoh → tugas → pertanyaan evaluasi → keluaran dan pemeriksaan. Seluruh kasus fiktif. Tanda [Butuh Office desktop] berarti fungsi tidak disimulasikan penuh pada webapp.


# Excel — 25 Modul

## Level Basic — 5 Modul

### E-01-01 · Mengenal workbook, sheet, dan struktur data kampus

**Tujuan:** Mengidentifikasi elemen worksheet dan membangun satu tabel data layanan yang konsisten.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Workbook adalah berkas Excel; worksheet adalah lembar di dalamnya.
- Header menjelaskan makna setiap kolom dan harus unik.
- Satu baris mewakili satu transaksi, permohonan, atau kegiatan.
- Kode layanan disimpan terpisah dari nama, tanggal, dan nilai.
- Hindari menggabung sel di area data yang akan difilter.

**Langkah praktikum:**

1. Buka berkas latihan kosong dan beri nama sheet Register_Layanan.
2. Tulis header No, Kode, Tanggal, Unit, Uraian, Status.
3. Isi tiga baris data contoh dengan jenis data yang konsisten.
4. Pilih sel D3 dan amati Name Box serta Formula Bar.
5. Simpan berkas sebagai Register_Layanan_v01.xlsx.

**Contoh terarah:** A1=Kode; B1=Tanggal; C1=Unit; D1=Uraian; A2=LAY-001; C2=Keuangan

**Tugas praktik peserta:** Buat tabel 6 kolom berisi 5 permohonan fiktif dari dua unit, tanpa baris kosong di tengah.

**Keluaran yang harus dikumpulkan:** Tabel register lima permohonan berformat XLSX

**Soal evaluasi mikro:** Apa struktur yang paling aman untuk rekap layanan?

**Pembahasan jawaban:** Tabel datar dengan header tunggal mendukung sort, filter, rumus, dan audit.

**Catatan kehati-hatian:** Jangan memasukkan nomor identitas mahasiswa asli pada data latihan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-01-02 · Memasukkan kode, tanggal, angka, dan data fiktif

**Tujuan:** Menghindari perubahan format otomatis yang merusak kode atau tanggal.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** keuangan, kemahasiswaan, akademik

**Materi inti:**

- Nomor surat, NIM contoh, dan kode akun dapat diperlakukan sebagai teks.
- Tanggal harus berupa nilai tanggal, bukan hanya tampilan teks.
- Nilai rupiah disimpan sebagai angka tanpa mengetik "Rp" di dalam nilai sel.
- Angka yang tersimpan sebagai teks dapat mengganggu SUM.
- Gunakan contoh sintetis, bukan daftar mahasiswa atau rekening riil.

**Langkah praktikum:**

1. Buat kolom Kode, Tanggal, Nominal, dan Status.
2. Ketik kode 00127 sebagai teks agar nol di depan tetap ada.
3. Masukkan tanggal lalu pilih format tanggal Indonesia.
4. Isi nominal 1500000 sebagai angka, bukan "Rp 1.500.000".
5. Uji SUM pada dua nilai nominal.

**Contoh terarah:** Kode: 00127 (teks) · Tanggal: 18/11/2026 · Nominal: 1500000 (angka)

**Tugas praktik peserta:** Buat tiga baris data pembayaran fiktif; tunjukkan perbedaan angka dan kode teks.

**Keluaran yang harus dikumpulkan:** Dataset sintetis tiga transaksi dengan tipe data benar

**Soal evaluasi mikro:** Mengapa kode 00127 perlu perhatian khusus?

**Pembahasan jawaban:** Kode identifikasi adalah label; format teks membantu menjaga nol di depan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-01-03 · Format Rupiah, persen, tanggal, dan keterbacaan

**Tujuan:** Menerapkan Number Format tanpa mengubah nilai sumber perhitungan.

**Estimasi:** 14 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Format mata uang menambah lambang dan pemisah ribuan secara visual.
- Persentase 0,25 dapat ditampilkan sebagai 25%.
- Tanggal perlu disajikan konsisten di seluruh kolom.
- Gunakan judul dan satuan rupiah atau orang dalam header.
- Atur lebar kolom agar nomor dan nilai tidak terpotong.

**Langkah praktikum:**

1. Pilih kolom nominal pada tabel latihan.
2. Terapkan format Number atau Accounting dengan simbol rupiah.
3. Masukkan 0,85 pada kolom capaian lalu terapkan Percentage.
4. Atur format tanggal secara konsisten.
5. Periksa bahwa Formula Bar tetap memuat nilai numerik.

**Contoh terarah:** 1500000 → Rp1.500.000 (tampilan); 0,85 → 85%

**Tugas praktik peserta:** Format tabel rekap pagu–realisasi dengan kolom Rp dan persentase capaian.

**Keluaran yang harus dikumpulkan:** Tabel dua nilai dengan format rupiah dan persentase

**Soal evaluasi mikro:** Jika 0,85 diformat sebagai persen, hasil tampil adalah ...

**Pembahasan jawaban:** Format Percentage mengalikan tampilan pecahan dengan 100.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-01-04 · Navigasi, seleksi rentang, dan shortcut kerja

**Tujuan:** Memakai shortcut dasar untuk memilih, mencari, dan menyimpan data administratif.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Name Box memudahkan navigasi ke sel yang diketahui.
- Ctrl+F mencari nilai pada worksheet.
- Ctrl+S menyimpan perubahan.
- Ctrl+Z membatalkan tindakan; Ctrl+C dan Ctrl+V menyalin.
- Ctrl+Shift+L dapat mengaktifkan filter pada rentang dengan header.

**Langkah praktikum:**

1. Buka workbook Register_Layanan.
2. Pilih A1:F6 menggunakan Name Box atau seleksi mouse.
3. Tekan Ctrl+F dan cari kode LAY-003.
4. Salin satu sel contoh ke baris lain lalu batalkan dengan Ctrl+Z.
5. Simpan perubahan dengan Ctrl+S.

**Contoh terarah:** Navigasi: A1:F6 → Ctrl+F → LAY-003 → Esc → Ctrl+S

**Tugas praktik peserta:** Temukan dua kode layanan pada file simulasi dan tulis alamat selnya.

**Keluaran yang harus dikumpulkan:** Daftar alamat dua sel hasil pencarian

**Soal evaluasi mikro:** Shortcut untuk menyimpan workbook ialah ...

**Pembahasan jawaban:** Ctrl+S adalah shortcut simpan yang umum digunakan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-01-05 · Penamaan berkas, folder, dan versi kerja

**Tujuan:** Mengelola berkas latihan untuk menghindari salah versi atau kehilangan sumber.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** arsip, semua

**Materi inti:**

- Pisahkan folder Sumber, Proses, dan Final sesuai SOP lokal.
- Nama berkas berisi unit, periode, jenis, dan versi.
- Simpan sumber mentah dalam keadaan tidak berubah.
- CSV cocok untuk transfer tabel sederhana tetapi tidak mempertahankan rumus/format.
- Pemberian akses mengikuti ketentuan unit, bukan sekadar kemampuan teknis.

**Langkah praktikum:**

1. Buat folder Latihan/01_Sumber, 02_Proses, 03_Final.
2. Simpan file latihan sebagai Rekap_Layanan_2026_v01.xlsx.
3. Buat satu salinan versi v02 setelah revisi.
4. Ekspor satu sheet menjadi CSV untuk pertukaran data.
5. Periksa apakah CSV kehilangan formula dan tampilan format.

**Contoh terarah:** 2026_Kemahasiswaan_RegisterBeasiswa_v02.xlsx

**Tugas praktik peserta:** Rancang struktur folder semester berjalan dengan berkas sumber dan final terpisah.

**Keluaran yang harus dikumpulkan:** Struktur folder serta skema penamaan versi kerja

**Soal evaluasi mikro:** Untuk menghindari kehilangan rumus saat berbagi, format yang lebih tepat adalah ...

**Pembahasan jawaban:** XLSX mempertahankan struktur workbook, formula, dan pemformatan.

**Catatan kehati-hatian:** Retensi dan akses arsip harus mengikuti ketentuan institusi.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Beginner — 5 Modul

### E-02-01 · Perhitungan pagu, realisasi, dan selisih

**Tujuan:** Menghitung selisih anggaran menggunakan referensi sel sehingga dapat diperiksa ulang.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Rumus dimulai dengan tanda sama dengan (=).
- Selisih anggaran = pagu dikurangi realisasi.
- Gunakan referensi C2 dan D2, bukan menulis nominal ulang.
- Salin rumus ke baris berikut untuk mengurangi kesalahan ketik.
- Selisih negatif perlu diverifikasi sebagai kemungkinan melebihi pagu, bukan otomatis kesalahan.

**Langkah praktikum:**

1. Buka kasus Keuangan pada simulator Excel.
2. Perhatikan C2=Pagu dan D2=Realisasi.
3. Isi E2 dengan =C2-D2.
4. Salin logika ke E3 sampai E7 bila perlu.
5. Periksa selisih dengan satu perhitungan manual.

**Contoh terarah:** C2=5000000; D2=4200000; E2 =C2-D2 → 800000

**Tugas praktik peserta:** Buat kolom Selisih untuk enam kegiatan, lalu identifikasi nilai negatif.

**Keluaran yang harus dikumpulkan:** Rekap enam kegiatan dengan selisih tepat

**Soal evaluasi mikro:** Rumus selisih pagu minus realisasi adalah ...

**Pembahasan jawaban:** Pengurangan nilai anggaran terhadap realisasi menghasilkan sisa alokasi.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-02-02 · Rekap total, rata-rata, maksimum, dan hitungan baris

**Tujuan:** Merangkum volume layanan dan penggunaan dana tanpa mengetik penjumlahan manual.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- SUM menjumlahkan angka pada rentang.
- AVERAGE menghitung rata-rata angka.
- MAX dan MIN mencari nilai tertinggi/terendah.
- COUNT menghitung sel numerik; COUNTA menghitung sel tak kosong.
- Sel total diberi label jelas dan tidak dicampur ke rentang data berikutnya.

**Langkah praktikum:**

1. Perhatikan enam nilai realisasi D2:D7.
2. Hitung =SUM(D2:D7) pada D8.
3. Hitung =AVERAGE(D2:D7) pada C9.
4. Bandingkan =MAX(D2:D7) dengan angka terbesar.
5. Periksa hasil dengan kalkulator sebagai kontrol.

**Contoh terarah:** =SUM(D2:D7) · =AVERAGE(D2:D7) · =MAX(D2:D7)

**Tugas praktik peserta:** Hitung total, rata-rata, dan realisasi tertinggi dari dataset latihan.

**Keluaran yang harus dikumpulkan:** Tiga indikator agregat dengan nilai sesuai sumber

**Soal evaluasi mikro:** Fungsi untuk mendapatkan nilai rata-rata adalah ...

**Pembahasan jawaban:** AVERAGE menghitung rerata numerik dari rentang.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-02-03 · Status layanan menggunakan IF

**Tujuan:** Mengategorikan keterlambatan atau kelebihan realisasi memakai aturan sederhana.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** keuangan, perencanaan, kemahasiswaan

**Materi inti:**

- IF memiliki kondisi, hasil bila benar, dan hasil bila salah.
- Operator logika mencakup >, >=, <, <=, =, dan <>.
- Hasil berupa teks harus diapit tanda kutip.
- Status otomatis berubah ketika data sumber berubah.
- Kebijakan threshold harus disetujui unit sebelum dipakai pada data nyata.

**Langkah praktikum:**

1. Tetapkan aturan simulasi: realisasi melebihi pagu = "Melebihi".
2. Pilih F2 dan masukkan =IF(D2>C2;"Melebihi";"Terkendali").
3. Tekan Enter dan amati hasilnya.
4. Ubah sementara D2 lebih besar dari C2.
5. Kembalikan data dan pastikan status otomatis berubah.

**Contoh terarah:** =IF(D2>C2;"Melebihi";"Terkendali")

**Tugas praktik peserta:** Buat label status untuk satu baris dan jelaskan logika keputusan.

**Keluaran yang harus dikumpulkan:** Satu kolom status dengan pengujian kondisi

**Soal evaluasi mikro:** Fungsi IF digunakan untuk ...

**Pembahasan jawaban:** IF menguji kondisi dan mengembalikan salah satu nilai.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-02-04 · Sort, Filter, Table, dan validasi pengisian

**Tujuan:** Menyaring daftar permohonan tanpa merusak hubungan antarkolom.

**Estimasi:** 18 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, arsip, akademik

**Materi inti:**

- Sort menyusun urutan baris berdasarkan kolom tertentu.
- Filter menyembunyikan baris yang tidak memenuhi kriteria.
- Excel Table memudahkan pengembangan data.
- Data Validation dapat membatasi daftar status atau unit.
- Jangan melakukan sort hanya pada satu kolom tabel yang saling berkaitan.

**Langkah praktikum:**

1. Pilih seluruh rentang termasuk header pada file latihan.
2. Gunakan Ctrl+T untuk membuat Table.
3. Aktifkan filter kategori atau unit.
4. Buat dropdown Status menggunakan Data Validation (Office desktop).
5. Periksa total baris setelah filter dibersihkan.

**Contoh terarah:** Status yang diizinkan: Diajukan; Diproses; Selesai; Perlu Perbaikan

**Tugas praktik peserta:** Buat dropdown status untuk register layanan dan filter hanya data "Diproses".

**Keluaran yang harus dikumpulkan:** Register layanan dengan validasi status dan filter teruji

**Soal evaluasi mikro:** Operasi yang menampilkan sebagian baris tanpa menghapus data adalah ...

**Pembahasan jawaban:** Filter membatasi tampilan, sedangkan data sumber tetap ada.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-02-05 · Grafik perbandingan target dan realisasi

**Tujuan:** Menyajikan dua indikator dalam grafik yang terbaca oleh pimpinan unit.

**Estimasi:** 18 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, kepegawaian, kemahasiswaan

**Materi inti:**

- Grafik kolom cocok untuk membandingkan nilai antarkategori.
- Judul harus menjelaskan periode dan ukuran.
- Sumbu dan satuan mencegah salah baca angka.
- Pilih data yang sudah diperiksa, bukan sembarang range.
- Jangan memakai efek 3D yang mengaburkan perbedaan.

**Langkah praktikum:**

1. Siapkan tabel Unit, Target, Realisasi.
2. Pilih tiga kolom tersebut termasuk header.
3. Insert → Clustered Column Chart.
4. Beri judul "Target dan Realisasi Layanan".
5. Tambahkan label satuan dan tulis satu temuan di bawah grafik.

**Contoh terarah:** Unit A: Target 40, Realisasi 36; Unit B: Target 50, Realisasi 45

**Tugas praktik peserta:** Buat grafik dua unit dan jelaskan unit yang pencapaiannya lebih tinggi.

**Keluaran yang harus dikumpulkan:** Grafik kolom berlabel, dengan satu kalimat insight

**Soal evaluasi mikro:** Grafik yang tepat untuk perbandingan kategori adalah ...

**Pembahasan jawaban:** Kolom berkelompok memudahkan perbandingan nilai dua seri per kategori.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Intermediate — 5 Modul

### E-03-01 · COUNTIF dan SUMIF menurut unit atau status

**Tujuan:** Menghasilkan rekap kategori pada register administrasi.

**Estimasi:** 22 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** keuangan, perencanaan, aset

**Materi inti:**

- COUNTIF menghitung entri yang cocok dengan satu kriteria.
- SUMIF menjumlahkan nilai pada baris yang cocok dengan kategori.
- Range kriteria dan range nilai harus mempunyai panjang sesuai.
- Kriteria teks ditulis dalam kutip atau dirujuk dari sel.
- Lakukan pemeriksaan silang dengan data terfilter.

**Langkah praktikum:**

1. Identifikasi kategori pada kolom B di kasus Keuangan.
2. Tulis =COUNTIF(B2:B7;"Operasional") pada G5.
3. Hitung total realisasi kategori yang sama dengan SUMIF.
4. Bandingkan hasil dengan daftar baris Operasional.
5. Catat apakah ada spasi atau ejaan kategori yang tidak konsisten.

**Contoh terarah:** =COUNTIF(B2:B7;"Operasional") · =SUMIF(B2:B7;"Operasional";D2:D7)

**Tugas praktik peserta:** Buat rekap jumlah transaksi dan total realisasi untuk dua kelompok kegiatan.

**Keluaran yang harus dikumpulkan:** Rekap per kategori dengan hasil terverifikasi

**Soal evaluasi mikro:** Rumus untuk menghitung jumlah baris berkategori "Operasional" adalah ...

**Pembahasan jawaban:** COUNTIF menghitung jumlah sel sesuai kriteria.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-03-02 · SUMIFS dan COUNTIFS untuk dua kriteria

**Tujuan:** Melakukan rekap bersyarat berdasarkan unit dan status atau periode.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, akademik, arsip

**Materi inti:**

- SUMIFS mendukung lebih dari satu pasangan kriteria.
- COUNTIFS menghitung baris yang memenuhi semua kondisi sekaligus.
- Rentang harus sejajar dan mempunyai jumlah baris sama.
- Kriteria tanggal sebaiknya memakai batas awal dan akhir yang jelas.
- Periksa apakah kriteria kosong atau bergeser.

**Langkah praktikum:**

1. Buka data transaksi simulasi pada Microsoft Excel.
2. Siapkan kolom Unit, Status, dan Nilai.
3. Hitung jumlah layanan Unit A yang berstatus Selesai dengan COUNTIFS.
4. Hitung nilai transaksi Unit A pada bulan tertentu menggunakan SUMIFS.
5. Bandingkan satu sampel baris secara manual.

**Contoh terarah:** =COUNTIFS(B2:B100;"Akademik";E2:E100;"Selesai")

**Tugas praktik peserta:** Buat rekap dua kriteria: jenis layanan dan status, lengkap dengan kontrol total.

**Keluaran yang harus dikumpulkan:** Tabel ringkasan dua kriteria dengan kontrol angka

**Soal evaluasi mikro:** Keunggulan COUNTIFS dibanding COUNTIF adalah ...

**Pembahasan jawaban:** COUNTIFS menerima lebih dari satu rentang kondisi dan kriteria.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-03-03 · XLOOKUP dan VLOOKUP untuk referensi kode

**Tujuan:** Mengisi uraian berdasarkan kode akun, kegiatan, atau layanan fiktif.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** akademik, arsip, keuangan

**Materi inti:**

- Pencarian tepat menghindari salah pengambilan kode.
- VLOOKUP mencari pada kolom pertama tabel referensi.
- XLOOKUP menentukan kolom pencarian dan hasil secara terpisah.
- Jika kode tidak ada, tampilkan status perlu verifikasi.
- Jangan memakai lookup sebagai pembenaran mengabaikan pengecekan sumber.

**Langkah praktikum:**

1. Siapkan sheet Master_Kode dengan Kode dan Nama Layanan.
2. Beri kode contoh ADM-01 dan ADM-02.
3. Gunakan XLOOKUP dengan parameter if_not_found pada kolom uraian.
4. Bandingkan tiga hasil kode terhadap master.
5. Tandai kode yang tidak ditemukan untuk perbaikan.

**Contoh terarah:** =XLOOKUP(A2;Master_Kode!A:A;Master_Kode!B:B;"Perlu cek")

**Tugas praktik peserta:** Petakan lima kode layanan ke nama layanan dengan pencarian exact match.

**Keluaran yang harus dikumpulkan:** Tabel lima kode dengan penandaan kode tidak dikenal

**Soal evaluasi mikro:** Fungsi yang mencari kode dan mengembalikan nilai terkait adalah ...

**Pembahasan jawaban:** XLOOKUP memetakan kunci ke kolom hasil yang ditentukan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-03-04 · Membersihkan teks, tanggal, dan duplikat data

**Tujuan:** Mengoreksi masalah umum sebelum rekap atau penggabungan berkas.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, arsip, akademik

**Materi inti:**

- TRIM menghapus spasi berlebih.
- LEFT, RIGHT, dan MID mengurai sebagian kode teks.
- DATE dan TEXT membantu membentuk/menampilkan tanggal.
- Remove Duplicates menghapus baris tertentu dan perlu backup.
- Data pribadi dalam daftar layanan harus disamarkan saat latihan.

**Langkah praktikum:**

1. Duplikasi file sumber terlebih dahulu.
2. Buat kolom bantu Kode_Bersih menggunakan TRIM.
3. Normalisasi ejaan status dan format tanggal.
4. Identifikasi duplikat berdasarkan kombinasi kode dan tanggal.
5. Catat baris yang dikoreksi serta alasan perbaikannya.

**Contoh terarah:** =TRIM(A2) · =LEFT(B2;3) · =TEXT(C2;"yyyy-mm-dd")

**Tugas praktik peserta:** Bersihkan sepuluh kode layanan contoh yang mengandung spasi dan satu duplikat.

**Keluaran yang harus dikumpulkan:** Dataset bersih plus log perbaikan data

**Soal evaluasi mikro:** Fungsi untuk menghapus spasi berlebih pada teks adalah ...

**Pembahasan jawaban:** TRIM menghilangkan spasi berlebih yang sering menyebabkan lookup gagal.

**Catatan kehati-hatian:** Jangan langsung menghapus duplikat tanpa memeriksa apakah transaksi memang berbeda.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-03-05 · PivotTable untuk rekap multidimensi unit kerja

**Tujuan:** Menyusun ringkasan data berdasarkan unit, periode, dan status.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, perencanaan, kepegawaian

**Materi inti:**

- PivotTable mengelompokkan baris sumber ke dalam kategori ringkasan.
- Rows dan Columns menentukan dimensi analisis.
- Values menentukan agregasi COUNT atau SUM.
- Refresh diperlukan bila sumber mengalami perubahan.
- Periksa grand total dan nilai kosong sebelum mengirim laporan.

**Langkah praktikum:**

1. Jadikan data sumber sebagai Excel Table.
2. Insert → PivotTable pada sheet baru.
3. Tempatkan Unit di Rows, Status di Columns, dan Kode di Values sebagai Count.
4. Aktifkan filter atau slicer bila diperlukan.
5. Cocokkan Grand Total dengan jumlah record sumber.

**Contoh terarah:** Rows: Unit; Columns: Status; Values: Count of Kode

**Tugas praktik peserta:** Buat PivotTable jumlah layanan per unit dan status pada data fiktif.

**Keluaran yang harus dikumpulkan:** PivotTable dengan grand total teruji

**Soal evaluasi mikro:** Komponen Values pada PivotTable berfungsi untuk ...

**Pembahasan jawaban:** Values memuat ukuran numerik atau hitungan yang diringkas.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Advanced — 5 Modul

### E-04-01 · Power Query: konsolidasi berkas per bulan

**Tujuan:** Menggabungkan laporan berformat sama ke tabel bersih yang bisa diperbarui.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Power Query memisahkan proses impor, transformasi, dan muat data.
- Sumber perlu memiliki skema kolom konsisten.
- Kolom tanggal dan nominal harus bertipe benar.
- Simpan langkah transformasi untuk audit.
- Refresh harus diuji dan dicatat hasilnya.

**Langkah praktikum:**

1. Siapkan tiga CSV sintetis Januari–Maret di satu folder.
2. Data → Get Data → From Folder (Excel desktop yang mendukung).
3. Gabungkan tabel lalu tetapkan tipe kolom.
4. Hilangkan baris kosong dan beri kolom Periode.
5. Load ke Table dan cocokkan jumlah record dengan sumber.

**Contoh terarah:** 3 CSV × 10 baris → target 30 transaksi tanpa baris ganda

**Tugas praktik peserta:** Bangun kueri konsolidasi bulanan dan catat total sumber sebelum/sesudah.

**Keluaran yang harus dikumpulkan:** Query konsolidasi dan log jumlah baris

**Soal evaluasi mikro:** Keunggulan Power Query untuk laporan berulang adalah ...

**Pembahasan jawaban:** Power Query menyimpan rangkaian langkah transformasi yang dapat dijalankan kembali.

**Catatan kehati-hatian:** Periksa perbedaan nama kolom serta format tanggal antarsumber.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-04-02 · Rekonsiliasi dan audit rumus keuangan

**Tujuan:** Mencari ketidaksesuaian antara buku kas, register, dan ringkasan sebelum pelaporan.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, aset

**Materi inti:**

- Rekonsiliasi membandingkan dua sumber pada periode dan definisi yang sama.
- Selisih nol belum tentu cukup jika transaksi yang berpasangan tidak sesuai.
- Gunakan kode unik dan tanggal sebagai pengenal.
- IFERROR tidak boleh menutupi error tanpa penjelasan.
- Audit trail menyimpan asal, koreksi, dan pemeriksa.

**Langkah praktikum:**

1. Pisahkan sheet BukuKas, Register, dan Rekonsiliasi.
2. Samakan rentang tanggal serta satuan nominal.
3. Gunakan lookup untuk mencocokkan referensi transaksi.
4. Hitung selisih per transaksi dan total.
5. Susun daftar pengecualian untuk diklarifikasi.

**Contoh terarah:** Sumber A Rp12.500.000; Sumber B Rp12.300.000; selisih Rp200.000

**Tugas praktik peserta:** Rancang log rekonsiliasi 10 transaksi fiktif dengan status Cocok/Perlu Periksa.

**Keluaran yang harus dikumpulkan:** Laporan pengecualian dan log koreksi

**Soal evaluasi mikro:** Jika total dua sumber berbeda, tindakan yang tepat adalah ...

**Pembahasan jawaban:** Selisih harus ditelusuri, bukan disesuaikan tanpa dasar.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-04-03 · Dashboard KPI layanan, slicer, dan pembacaan grafik

**Tujuan:** Menampilkan kinerja unit secara ringkas dengan definisi indikator yang jelas.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, kemahasiswaan, akademik

**Materi inti:**

- KPI harus memiliki definisi, satuan, periode, sumber, dan pemilik.
- Dashboard ringkas mengutamakan target, realisasi, dan deviasi.
- Slicer membantu memfilter PivotTable atau model yang didukung.
- Grafik dibuat dari data yang dapat ditelusuri.
- Warna indikator tidak boleh menjadi satu-satunya pembeda status.

**Langkah praktikum:**

1. Tentukan tiga KPI: layanan selesai, waktu proses, dan tingkat penyelesaian.
2. Bangun PivotTable dari register sintetis.
3. Buat kartu metrik dan grafik bulanan.
4. Tambahkan slicer Unit untuk laporan internal.
5. Uji bahwa perubahan filter mengubah indikator secara konsisten.

**Contoh terarah:** KPI: 80 layanan masuk, 72 selesai, tingkat penyelesaian 90%

**Tugas praktik peserta:** Bangun dashboard satu halaman untuk kepala bagian dan tulis dua temuan.

**Keluaran yang harus dikumpulkan:** Dashboard KPI dan kamus indikator

**Soal evaluasi mikro:** Definisi KPI yang dapat diaudit minimal memuat ...

**Pembahasan jawaban:** KPI perlu didefinisikan agar semua pembaca menafsirkan angka yang sama.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-04-04 · Analisis skenario anggaran dan target

**Tujuan:** Membandingkan beberapa asumsi sebelum mengusulkan alokasi kerja.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Skenario dasar memakai asumsi yang terdokumentasi.
- What-If Analysis dapat membantu melihat sensitivitas input.
- Pisahkan input, perhitungan, dan output.
- Jangan mengubah angka realisasi historis untuk memenuhi skenario.
- Tuliskan keterbatasan agar hasil tidak diperlakukan sebagai kepastian.

**Langkah praktikum:**

1. Buat input pagu, volume layanan, dan biaya rata-rata.
2. Tetapkan tiga skenario: dasar, efisiensi, peningkatan layanan.
3. Hitung kebutuhan =volume×biaya.
4. Bandingkan kebutuhan dan batas pagu.
5. Tuliskan rekomendasi dengan konsekuensi tiap opsi.

**Contoh terarah:** 100 layanan × Rp150.000 = Rp15.000.000; naik 10% volume = 110 layanan

**Tugas praktik peserta:** Buat matriks tiga skenario kebutuhan operasional untuk rapat perencanaan.

**Keluaran yang harus dikumpulkan:** Matriks tiga skenario beserta asumsi

**Soal evaluasi mikro:** Pemisahan penting dalam analisis skenario adalah ...

**Pembahasan jawaban:** Asumsi terdokumentasi memungkinkan evaluasi sensitivitas perubahan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-04-05 · Kontrol akses, proteksi, dan data sensitif

**Tujuan:** Mendesain workbook administrasi yang meminimalkan risiko kesalahan dan kebocoran.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, akademik, arsip, kepegawaian

**Materi inti:**

- Gunakan dataset anonim/sintetis untuk latihan.
- Sheet Protection membatasi penyuntingan, tetapi bukan solusi keamanan penuh.
- Identifikasi kolom input, formula, dan data terbatas.
- Berikan akses sesuai kebutuhan kerja dan kebijakan institusi.
- Pencadangan dan pemeriksaan versi perlu memiliki penanggung jawab.

**Langkah praktikum:**

1. Tandai kolom input yang boleh diubah pengguna.
2. Warnai sel rumus dan beri keterangan sumber.
3. Lindungi formula sesuai versi Excel desktop.
4. Buat versi publik yang tidak memuat nama/NIM/nomor rekening nyata.
5. Dokumentasikan peran penyiap dan pemeriksa.

**Contoh terarah:** Label akses latihan: PUBLIK / INTERNAL / TERBATAS (contoh, bukan klasifikasi resmi)

**Tugas praktik peserta:** Susun checklist keamanan sebelum workbook rekap diserahkan antarunit.

**Keluaran yang harus dikumpulkan:** Checklist kontrol akses dan versi contoh tanpa data pribadi

**Soal evaluasi mikro:** Perlindungan yang benar bagi data mahasiswa adalah ...

**Pembahasan jawaban:** Prinsip minimisasi data dan pengendalian akses mengurangi risiko pemrosesan.

**Catatan kehati-hatian:** Patuhi kebijakan keamanan kampus dan peraturan terkait; OfficeLab bukan konsultasi kepatuhan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Expert — 5 Modul

### E-05-01 · Model anggaran multi-unit dan konsolidasi usulan

**Tujuan:** Merancang model perencanaan yang dapat diaudit dari sumber unit hingga ringkasan kampus.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Kamus kode akun harus konsisten antarunit.
- Pemisahan input, asumsi, formula, dan output menurunkan risiko.
- Gunakan validasi skema dan rekap per unit.
- Buat pengujian jumlah total sebelum menggabungkan usulan.
- Semua perubahan pagu perlu jejak persetujuan di alur kerja institusi.

**Langkah praktikum:**

1. Buat sheet Master_Akun dan daftar unit.
2. Siapkan format input usulan yang sama untuk tujuh unit.
3. Validasi kode, unit, jumlah, dan asumsi.
4. Konsolidasi serta tampilkan total per jenis belanja.
5. Dokumentasikan sumber dan anomali untuk ditindaklanjuti.

**Contoh terarah:** Anggaran Unit A Rp20 juta, B Rp15 juta, C Rp12 juta → total Rp47 juta

**Tugas praktik peserta:** Siapkan model anggaran sintetis 3 unit dengan cek duplikat dan ringkasan keputusan.

**Keluaran yang harus dikumpulkan:** Model input-konsolidasi-output dengan catatan audit

**Soal evaluasi mikro:** Apa kontrol paling penting pada konsolidasi usulan?

**Pembahasan jawaban:** Keseragaman master data dan rekonsiliasi total mendukung akurasi.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-05-02 · Kamus data, Power Pivot, dan ukuran DAX

**Tujuan:** Memahami model data berelasi serta definisi metrik yang digunakan berulang.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, akademik

**Materi inti:**

- Tabel fakta berisi transaksi/kejadian; dimensi berisi referensi.
- Kunci relasi harus unik pada sisi dimensi yang sesuai.
- Measure DAX menghitung nilai dalam konteks filter.
- Pisahkan metrik sumber dari metrik turunan.
- Versi Excel dan lisensi dapat membatasi fitur Power Pivot.

**Langkah praktikum:**

1. Definisikan tabel Fakta_Layanan, Dim_Unit, dan Dim_Tanggal.
2. Buat kunci hubungan UnitID dan DateKey.
3. Tentukan measure Total_Layanan dan Layanan_Selesai.
4. Uji filter berdasarkan unit dan bulan.
5. Simpan kamus data beserta rumus serta satuannya.

**Contoh terarah:** Total Layanan := COUNTROWS(Fakta_Layanan)  (contoh DAX)

**Tugas praktik peserta:** Gambar skema model tiga tabel dan dokumentasikan tiga ukuran kinerja.

**Keluaran yang harus dikumpulkan:** Skema model data dan kamus tiga measure

**Soal evaluasi mikro:** Peran tabel dimensi dalam model analitik ialah ...

**Pembahasan jawaban:** Dimensi memberikan konteks untuk fakta yang dihitung.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-05-03 · Validasi lintas-sistem dan penanganan ketidaksesuaian

**Tujuan:** Menguji kecocokan berkas dari unit berbeda sebelum digabung dalam laporan institusi.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Sistem berbeda dapat memakai kode, format, dan periode yang tidak identik.
- Mapping kode adalah keputusan bisnis yang harus disahkan pemilik data.
- Uji keunikan, kelengkapan, kewajaran, dan keterlacakan.
- Laporan pengecualian harus memuat sumber serta prioritas.
- Jangan menganggap hasil cocok berarti sumber bebas masalah.

**Langkah praktikum:**

1. Buat dua file register sintetis dengan format kolom berbeda.
2. Susun tabel pemetaan kode antarberkas.
3. Gabungkan berdasarkan kunci yang telah dibersihkan.
4. Catat record tanpa pasangan dan nilai yang berbeda.
5. Buat ringkasan temuan serta usulan perbaikan master data.

**Contoh terarah:** UNIT-01 ↔ Bagian Keuangan; status "Done" ↔ "Selesai" (perlu disahkan)

**Tugas praktik peserta:** Rancang dan uji 4 aturan validasi untuk konsolidasi 20 record.

**Keluaran yang harus dikumpulkan:** Log aturan validasi, mismatch, dan tindakan koreksi

**Soal evaluasi mikro:** Kode berbeda antar-sistem seharusnya ...

**Pembahasan jawaban:** Pemetaan terdokumentasi menjaga makna dan jejak audit.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-05-04 · Otomatisasi laporan dengan Office Scripts atau VBA

**Tujuan:** Menilai kelayakan otomasi tugas rutin dan risiko perubahan otomatis.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan, arsip

**Materi inti:**

- Pilih proses yang stabil, berulang, dan terukur sebelum otomatisasi.
- VBA digunakan pada lingkungan Office desktop yang mendukung; Office Scripts pada lingkungan yang sesuai.
- Kode harus menghindari penghapusan data sumber dan mendokumentasikan hasil.
- Uji dengan data contoh serta buat salinan cadangan.
- Makro dari sumber tidak tepercaya berisiko keamanan.

**Langkah praktikum:**

1. Petakan tugas bulanan: impor, cek, rekap, dan ekspor.
2. Tentukan input/output serta kondisi gagal.
3. Buat pseudocode langkah yang akan diotomatisasi.
4. Uji pada copy dataset dan bandingkan hasil manual.
5. Susun checklist review dan prosedur rollback.

**Contoh terarah:** Pseudocode: IMPORT → VALIDASI → REKAP → CEK TOTAL → EXPORT → LOG

**Tugas praktik peserta:** Tulis rancangan otomasi rekap bulanan disertai pengujian dan penanganan error.

**Keluaran yang harus dikumpulkan:** Pseudocode dan rencana uji otomatisasi

**Soal evaluasi mikro:** Mengapa otomasi harus diuji pada salinan data?

**Pembahasan jawaban:** Pengujian pada data simulasi membatasi dampak kesalahan skrip.

**Catatan kehati-hatian:** Jangan mengaktifkan makro dari sumber yang tidak dikenal.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### E-05-05 · Proyek akhir Excel: dashboard monitoring administrasi

**Tujuan:** Menggabungkan pembersihan data, formula, validasi, dan visualisasi menjadi keluaran siap tinjau.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Proyek akhir harus memakai data sintetis yang konsisten dan terdokumentasi.
- Definisi indikator dipisahkan dari dataset sumber.
- Rekonsiliasi jumlah harus dapat diulang reviewer.
- Dashboard menyajikan temuan dan risiko, bukan sekadar dekorasi.
- Catat batas penggunaan dan hasil uji.

**Langkah praktikum:**

1. Pilih kasus unit: anggaran, mahasiswa, atau layanan.
2. Siapkan minimal 30 record sintetis dan kamus kolom.
3. Buat validasi data dan rekap bersyarat.
4. Tambahkan dua grafik dan tiga KPI terdefinisi.
5. Susun catatan audit dan ringkasan rekomendasi satu halaman.

**Contoh terarah:** Keluaran: data_sumber.xlsx + dashboard.xlsx + kamus_data + daftar_uji

**Tugas praktik peserta:** Bangun dashboard monitoring bulanan; tunjukkan kontrol total, 3 KPI, dan 2 grafik.

**Keluaran yang harus dikumpulkan:** Paket proyek Excel dengan dashboard dan dokumentasi audit

**Soal evaluasi mikro:** Proyek Excel dianggap siap tinjau jika ...

**Pembahasan jawaban:** Dashboard yang layak ditinjau dapat direproduksi dari sumbernya.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


# Word — 25 Modul

## Level Basic — 5 Modul

### W-01-01 · Mengenal dokumen dinas, Ribbon, dan area kerja

**Tujuan:** Menemukan menu dasar Word serta menjelaskan bagian dokumen administrasi.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Dokumen resmi memiliki tujuan, pihak tujuan, tanggal, isi, dan pengesahan sesuai pedoman institusi.
- Ribbon dan Quick Access Toolbar menyediakan fungsi pengeditan.
- Navigation Pane membantu menemukan heading pada dokumen panjang.
- Status Bar menampilkan halaman dan jumlah kata.
- Gunakan kop institusi hanya sesuai identitas dan format yang disahkan.

**Langkah praktikum:**

1. Buka Word dan buat dokumen baru.
2. Ketik judul "Catatan Layanan Administrasi".
3. Tuliskan tanggal dan unit pengelola secara fiktif.
4. Temukan menu Home, Insert, Layout, dan Review.
5. Simpan sebagai 2026_CatatanLayanan_v01.docx.

**Contoh terarah:** Judul: CATATAN LAYANAN ADMINISTRASI · Unit: Bagian Administrasi Kampus

**Tugas praktik peserta:** Buat satu halaman pengantar layanan dan identifikasi lima bagian dokumen.

**Keluaran yang harus dikumpulkan:** Satu halaman catatan layanan berjudul dan bernama berkas

**Soal evaluasi mikro:** Tab Word yang umum digunakan untuk pemeriksaan ejaan adalah ...

**Pembahasan jawaban:** Fitur pemeriksaan ada di tab Review/Peninjauan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-01-02 · Struktur surat: nomor, tanggal, perihal, dan lampiran

**Tujuan:** Menyusun elemen surat administratif tanpa menggunakan data atau tanda tangan asli.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** arsip, kepegawaian, akademik

**Materi inti:**

- Nomor, perihal, tanggal, lampiran, alamat tujuan, dan isi mengikuti tata naskah dinas institusi.
- Format latihan hanyalah contoh dan bukan format resmi kampus.
- Satu surat idealnya fokus pada satu maksud komunikasi.
- Jangan menyalin tanda tangan atau stempel pejabat.
- Verifikasi data tujuan sebelum distribusi.

**Langkah praktikum:**

1. Buka template latihan Surat Undangan.
2. Isi nomor surat fiktif dan tanggal contoh.
3. Tuliskan perihal yang spesifik dan tujuan surat.
4. Sampaikan agenda, lokasi, dan waktu secara ringkas.
5. Periksa unsur surat berdasarkan format instansi masing-masing.

**Contoh terarah:** Nomor: 001/ADM-LAT/XI/2026 · Perihal: Undangan Koordinasi Pelayanan

**Tugas praktik peserta:** Lengkapi unsur surat undangan kegiatan pelatihan administrasi dengan identitas sintetis.

**Keluaran yang harus dikumpulkan:** Draft surat undangan dengan unsur terisi

**Soal evaluasi mikro:** Unsur yang umumnya menjelaskan inti surat adalah ...

**Pembahasan jawaban:** Perihal merangkum pokok komunikasi.

**Catatan kehati-hatian:** Nomor surat dan format harus disesuaikan pedoman tata naskah dinas kampus.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-01-03 · Paragraf, font, spasi, dan Styles dasar

**Tujuan:** Merapikan dokumen yang mudah dibaca dan konsisten tanpa spasi manual berlebihan.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Gunakan Styles untuk judul dan heading agar konsisten.
- Atur jarak antarparagraf melalui Paragraph Spacing.
- Tab stop dan tabel lebih stabil dibanding rentetan spasi.
- Gunakan ukuran teks dan kontras yang cukup.
- Hindari memakai Enter berulang untuk pindah halaman.

**Langkah praktikum:**

1. Sorot judul dokumen lalu terapkan Heading 1.
2. Atur body text dengan font yang sesuai pedoman instansi.
3. Tentukan line spacing dan spacing after secara konsisten.
4. Aktifkan Show/Hide ¶ untuk memeriksa format tersembunyi.
5. Simpan dan lihat Print Preview.

**Contoh terarah:** Heading 1: NOTULEN RAPAT · Normal: uraian agenda

**Tugas praktik peserta:** Rapikan satu halaman notulen agar judul, subjudul, dan paragraf terbaca.

**Keluaran yang harus dikumpulkan:** Dokumen dengan Heading dan spacing konsisten

**Soal evaluasi mikro:** Cara paling stabil membuat judul bab konsisten adalah ...

**Pembahasan jawaban:** Styles menyatukan format dan mendukung navigasi dokumen.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-01-04 · Tabel sederhana untuk daftar hadir dan register

**Tujuan:** Menyusun informasi dalam kolom tanpa merusak susunan saat dicetak.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** kepegawaian, kemahasiswaan

**Materi inti:**

- Tabel cocok untuk data dengan kolom berulang.
- Header perlu memuat satuan dan deskripsi jelas.
- Hindari tabel lebar yang melewati margin.
- Ulangi baris header bila tabel melintasi halaman.
- Pisahkan data latihan fiktif dari daftar hadir nyata.

**Langkah praktikum:**

1. Insert → Table dan buat tabel empat kolom.
2. Isi No, Unit, Agenda, dan Status Kehadiran.
3. Tambahkan tiga baris nama unit fiktif.
4. Atur lebar kolom dan alignment sesuai jenis data.
5. Periksa Print Preview.

**Contoh terarah:** No | Unit | Agenda | Hadir (Ya/Tidak)

**Tugas praktik peserta:** Buat tabel daftar hadir enam peserta fiktif pada editor Word OfficeLab.

**Keluaran yang harus dikumpulkan:** Tabel daftar hadir dengan header dan enam baris

**Soal evaluasi mikro:** Untuk membuat kolom yang rapi, fitur yang lebih tepat adalah ...

**Pembahasan jawaban:** Tabel menjaga sel sejajar secara konsisten.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-01-05 · Simpan, ekspor PDF, dan kontrol versi dokumen

**Tujuan:** Memelihara versi penyuntingan dan mengirim dokumen dalam format sesuai tujuan.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** arsip, semua

**Materi inti:**

- DOCX cocok untuk naskah yang masih perlu diedit.
- PDF berguna untuk distribusi tampilan final tetapi bukan bukti persetujuan otomatis.
- Nama file perlu mengandung jenis dokumen, periode, dan versi.
- Sebelum berbagi, periksa komentar, metadata, serta lampiran.
- Pengelolaan arsip dan retensi mengikuti kebijakan kampus.

**Langkah praktikum:**

1. Simpan surat latihan sebagai DOCX versi v01.
2. Buka File → Save As atau Export sesuai versi Office.
3. Ekspor salinan PDF untuk pratinjau.
4. Pastikan halaman, tabel, dan margin tidak berubah.
5. Simpan hasil pada folder latihan Final.

**Contoh terarah:** 2026_Arsip_RegisterSurat_v01.docx → 2026_Arsip_RegisterSurat_FINAL.pdf

**Tugas praktik peserta:** Buat dua versi berkas (editable dan distribusi) dan periksa nama serta isi.

**Keluaran yang harus dikumpulkan:** DOCX dan salinan PDF yang sudah diperiksa

**Soal evaluasi mikro:** Format yang lazim dipakai untuk naskah Word yang masih diedit adalah ...

**Pembahasan jawaban:** DOCX mempertahankan kemampuan mengedit dokumen Word.

**Catatan kehati-hatian:** PDF bukan pengganti mekanisme tanda tangan/persetujuan yang diwajibkan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Beginner — 5 Modul

### W-02-01 · Surat undangan rapat dan koordinasi

**Tujuan:** Menulis surat undangan yang singkat, lengkap, dan dapat ditindaklanjuti.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** arsip, kepegawaian, perencanaan

**Materi inti:**

- Undangan memerlukan tujuan, agenda, waktu, tempat, dan kontak penyelenggara.
- Tuliskan tanggal lengkap agar tidak ambigu.
- Pilih kata kerja yang jelas dan sopan.
- Lakukan pemeriksaan pihak tujuan dan lampiran.
- Gunakan format dan jalur paraf resmi jika dokumen dipakai operasional.

**Langkah praktikum:**

1. Buka template Undangan Rapat di laboratorium Word.
2. Isi perihal dan daftar tujuan fiktif.
3. Tuliskan agenda dan waktu dengan tanggal penuh.
4. Sisipkan tabel agenda bila diperlukan.
5. Tinjau semua detail sebelum ekspor.

**Contoh terarah:** Agenda: Evaluasi pelayanan administrasi; Waktu: Jumat, 20 November 2026, 09.00 WITA

**Tugas praktik peserta:** Susun undangan koordinasi antarbagi­an dengan agenda dua butir.

**Keluaran yang harus dikumpulkan:** Surat undangan satu halaman dan agenda

**Soal evaluasi mikro:** Informasi yang penting ada dalam undangan rapat adalah ...

**Pembahasan jawaban:** Rincian logistik dan tujuan membuat undangan dapat ditindaklanjuti.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-02-02 · Surat tugas dan surat keterangan layanan

**Tujuan:** Menyusun naskah administratif berdasarkan data layanan yang dapat diverifikasi.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** akademik, kepegawaian, aset

**Materi inti:**

- Surat tugas menjelaskan penugasan, lingkup, dan periode.
- Surat keterangan menyatakan fakta yang perlu verifikasi sumber.
- Identitas dalam latihan harus sintetis.
- Hindari klaim yang tidak memiliki dasar pemeriksaan.
- Dokumen final memerlukan otorisasi sesuai aturan instansi.

**Langkah praktikum:**

1. Pilih format Surat Tugas pada editor.
2. Isi unit pelaksana dan uraian tugas contoh.
3. Tambahkan lokasi dan periode penugasan.
4. Masukkan daftar tugas dengan penomoran.
5. Tinjau apakah ada ketentuan atau rincian yang tertinggal.

**Contoh terarah:** Surat tugas: pendataan inventaris ruang layanan pada 23–25 November 2026

**Tugas praktik peserta:** Buat surat tugas pendataan aset untuk dua petugas fiktif.

**Keluaran yang harus dikumpulkan:** Draft surat tugas dengan ruang lingkup dan periode

**Soal evaluasi mikro:** Pernyataan yang tepat tentang surat keterangan adalah ...

**Pembahasan jawaban:** Surat keterangan memuat pernyataan faktual yang harus terverifikasi.

**Catatan kehati-hatian:** Jangan mencantumkan nama pejabat, tanda tangan, atau stempel tanpa izin.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-02-03 · Notulen dan berita acara kegiatan

**Tujuan:** Mencatat agenda, pembahasan, keputusan, penanggung jawab, dan tindak lanjut.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** perencanaan, arsip, semua

**Materi inti:**

- Notulen membedakan pembahasan dari keputusan.
- Tindak lanjut sebaiknya menyebut pemilik tugas dan tenggat.
- Berita acara mencatat kejadian atau serah terima secara terstruktur.
- Catatan rapat harus netral dan berbasis hasil nyata.
- Identifikasi peserta rapat hanya pada dokumen yang diotorisasi.

**Langkah praktikum:**

1. Buat judul Notulen Rapat Evaluasi Layanan.
2. Tambahkan agenda, peserta fiktif, dan waktu.
3. Tulis dua keputusan dengan action owner.
4. Sisipkan tabel Tindak Lanjut: aksi, penanggung jawab, tenggat.
5. Periksa bahwa poin keputusan tidak bertentangan dengan isi.

**Contoh terarah:** Keputusan 01: unit layanan memperbarui form registrasi; PIC: Tim Layanan; Tenggat: akhir bulan

**Tugas praktik peserta:** Susun notulen tiga agenda dan dua tindak lanjut lengkap.

**Keluaran yang harus dikumpulkan:** Notulen dua halaman dengan tabel tindak lanjut

**Soal evaluasi mikro:** Elemen yang membedakan notulen yang dapat ditindaklanjuti adalah ...

**Pembahasan jawaban:** Penanggung jawab dan tenggat membuat keputusan dapat dipantau.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-02-04 · Tabel inventaris, daftar hadir, dan lampiran

**Tujuan:** Menambahkan lampiran tabel yang tidak terpotong dan mudah diverifikasi.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** aset, arsip

**Materi inti:**

- Judul lampiran menjelaskan konteks dan periode.
- Kolom kuantitas harus mencantumkan satuan.
- Kolom status menggunakan nilai yang seragam.
- Periksa pengulangan header saat pergantian halaman.
- Tabel yang terlalu besar dapat dibuat sebagai lampiran Excel terpisah.

**Langkah praktikum:**

1. Buat dokumen Berita Acara Inventaris.
2. Sisipkan tabel No, Barang, Jumlah, Satuan, Kondisi.
3. Isi lima contoh aset sintetis.
4. Atur heading Lampiran 1 dan nomor halaman.
5. Tinjau apakah semua kolom muat saat dicetak.

**Contoh terarah:** Meja kerja | 3 | unit | Baik; Printer | 1 | unit | Perlu Servis

**Tugas praktik peserta:** Buat lampiran inventaris lima barang dengan satuan dan kondisi.

**Keluaran yang harus dikumpulkan:** Lampiran inventaris siap cetak

**Soal evaluasi mikro:** Informasi yang perlu ada pada tabel inventaris adalah ...

**Pembahasan jawaban:** Kolom yang konsisten mempermudah pemeriksaan inventaris.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-02-05 · Proofreading, bahasa resmi, dan pemeriksaan akhir

**Tujuan:** Menghapus kesalahan ketik serta ambiguitas pada surat dan laporan.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Proofreading memeriksa bahasa, angka, tanggal, dan referensi.
- Gunakan fitur Spelling & Grammar sebagai bantuan, bukan keputusan final.
- Pastikan penulisan nama unit, jabatan, dan singkatan konsisten.
- Cari klaim berlebihan atau kalimat yang tidak jelas pelakunya.
- Cek apakah semua lampiran benar-benar ada.

**Langkah praktikum:**

1. Buka surat latihan yang telah disusun.
2. Gunakan Ctrl+H untuk menemukan istilah yang tidak konsisten.
3. Baca ulang seluruh angka dan tanggal.
4. Periksa format tanggal serta alamat tujuan.
5. Tandai dokumen selesai diperiksa pada log latihan.

**Contoh terarah:** Sebelum: "data nya akan diperbaiki mungkin" → Sesudah: "Tim Layanan memperbaiki data pada 20 November 2026" bila faktanya terverifikasi

**Tugas praktik peserta:** Buat daftar minimal lima kesalahan surat contoh dan perbaikannya.

**Keluaran yang harus dikumpulkan:** Lembar pemeriksaan akhir yang terisi

**Soal evaluasi mikro:** Bukti proofreading yang baik berupa ...

**Pembahasan jawaban:** Catatan pemeriksaan memperjelas apa yang telah diperbaiki.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Intermediate — 5 Modul

### W-03-01 · Styles, heading, dan Navigation Pane laporan

**Tujuan:** Mengubah laporan panjang menjadi dokumen berstruktur yang mudah dinavigasi.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, akademik

**Materi inti:**

- Heading 1–3 membentuk hierarki judul yang dapat dipakai Navigation Pane.
- Jangan membuat heading hanya dengan ukuran font manual.
- Style dapat diubah satu kali untuk banyak bagian.
- Struktur dokumen perlu konsisten sampai subbab.
- Hindari subbab terlalu dalam jika tidak diperlukan.

**Langkah praktikum:**

1. Buat laporan dengan bagian Pendahuluan, Kegiatan, Hasil, dan Penutup.
2. Terapkan Heading 1 untuk bagian utama.
3. Gunakan Heading 2 untuk rincian di bawahnya.
4. Aktifkan View → Navigation Pane.
5. Periksa urutan dan kedalaman heading.

**Contoh terarah:** Heading 1: Hasil Kegiatan; Heading 2: Rekap Layanan Mahasiswa

**Tugas praktik peserta:** Format laporan dua halaman dengan sedikitnya empat heading yang dapat dinavigasi.

**Keluaran yang harus dikumpulkan:** Laporan terstruktur dengan navigasi judul

**Soal evaluasi mikro:** Navigation Pane paling efektif bila judul menggunakan ...

**Pembahasan jawaban:** Navigation Pane mengenali hierarki heading berbasis Styles.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-03-02 · Section Break, header/footer, dan halaman berbeda

**Tujuan:** Mengatur dokumen campuran potret dan lanskap dengan penomoran terkontrol.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan, arsip

**Materi inti:**

- Page Break memulai halaman tanpa mengubah bagian.
- Section Break membuat pengaturan halaman berbeda antarbagian.
- Header/Footer dapat mengandung nomor halaman dan identitas dokumen.
- Link to Previous menentukan keterkaitan antarseksi.
- Periksa halaman kosong akibat break berlebihan.

**Langkah praktikum:**

1. Buat laporan dua bagian dengan halaman potret.
2. Sisipkan Section Break sebelum tabel lebar.
3. Ubah orientasi hanya section tabel menjadi lanskap.
4. Atur header dan nomor halaman untuk setiap section.
5. Tinjau total halaman dalam Print Preview.

**Contoh terarah:** Bagian I potret; Lampiran tabel lanskap; Bagian III kembali potret

**Tugas praktik peserta:** Buat laporan tiga section dengan lampiran lanskap satu halaman.

**Keluaran yang harus dikumpulkan:** Dokumen multi-section dengan nomor halaman konsisten

**Soal evaluasi mikro:** Fitur untuk mengubah orientasi hanya satu bagian ialah ...

**Pembahasan jawaban:** Section Break memisahkan pengaturan tata letak bagian dokumen.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-03-03 · Daftar isi otomatis, caption, dan daftar tabel

**Tujuan:** Membuat navigasi laporan resmi dari Styles dan caption.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, akademik

**Materi inti:**

- Daftar isi otomatis mengambil heading yang diberi style.
- Caption memberi nomor dan label tabel/gambar.
- Daftar tabel dapat dibangkitkan dari caption yang benar.
- Perbarui field setelah judul atau nomor berubah.
- Lakukan pemeriksaan silang nomor tabel serta rujukannya.

**Langkah praktikum:**

1. Terapkan Heading 1/2 pada laporan latihan.
2. References → Table of Contents → Automatic Table.
3. Beri caption pada dua tabel simulasi.
4. Sisipkan List of Tables bila diperlukan.
5. Update Table dan periksa halaman tujuan.

**Contoh terarah:** Daftar isi: 1 Pendahuluan; 2 Metode; 3 Hasil; Lampiran

**Tugas praktik peserta:** Buat daftar isi otomatis dari minimal empat heading dan dua caption tabel.

**Keluaran yang harus dikumpulkan:** Laporan dengan daftar isi dan caption yang diperbarui

**Soal evaluasi mikro:** Jika judul bab berubah, tindakan yang diperlukan adalah ...

**Pembahasan jawaban:** Field otomatis perlu di-update agar daftar isi tetap sinkron.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-03-04 · Mail Merge surat massal dari Excel

**Tujuan:** Menghasilkan surat personalisasi dari daftar penerima yang sah dan terverifikasi.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kepegawaian, kemahasiswaan, akademik

**Materi inti:**

- Mail Merge menghubungkan dokumen utama dengan tabel sumber.
- Kolom sumber harus memiliki header yang stabil.
- Gunakan daftar fiktif pada praktik agar tidak memproses data pribadi riil.
- Pratinjau tiap variasi sebelum membuat surat massal.
- Cek format tanggal, alamat, dan kode yang kehilangan nol depan.

**Langkah praktikum:**

1. Siapkan data Excel Nama_Penerima dan Unit dengan 5 contoh fiktif.
2. Buka Mailings → Start Mail Merge → Letters.
3. Pilih daftar penerima sintetis.
4. Sisipkan merge fields dalam badan surat.
5. Preview Results dan tinjau setiap surat sebelum Finish & Merge.

**Contoh terarah:** Yth. «Nama_Penerima», Unit «Unit» — Undangan Pelatihan Administrasi

**Tugas praktik peserta:** Buat lima surat undangan personalisasi dari satu template dan data contoh.

**Keluaran yang harus dikumpulkan:** Lima surat hasil merge yang diperiksa

**Soal evaluasi mikro:** Sumber data yang digunakan Mail Merge adalah ...

**Pembahasan jawaban:** Mail Merge memetakan field dari daftar penerima.

**Catatan kehati-hatian:** Jangan menyebarluaskan daftar penerima atau mencetak data tanpa kewenangan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-03-05 · Komentar, Track Changes, dan alur reviu

**Tujuan:** Mengelola masukan dari beberapa peninjau tanpa kehilangan sejarah revisi.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Track Changes menandai perubahan pada naskah.
- Komentar menjelaskan alasan atau pertanyaan review.
- Reviu perlu membedakan saran, keputusan, dan persetujuan.
- Accept/Reject dilakukan dengan kewenangan yang jelas.
- Dokumen final perlu diperiksa dari komentar atau revisi tersisa.

**Langkah praktikum:**

1. Aktifkan Review → Track Changes di Word asli.
2. Perbaiki satu paragraf sehingga tanda revisi muncul.
3. Tambahkan dua komentar yang dapat ditindaklanjuti.
4. Simulasikan penerimaan satu perubahan dan penolakan yang lain.
5. Buat salinan final setelah pemeriksaan menyeluruh.

**Contoh terarah:** Komentar: "Mohon verifikasi dasar angka 72 layanan selesai"

**Tugas praktik peserta:** Lakukan putaran reviu surat tugas dengan minimal tiga catatan revisi.

**Keluaran yang harus dikumpulkan:** Dokumen revisi dan log keputusan editor

**Soal evaluasi mikro:** Fungsi Track Changes adalah ...

**Pembahasan jawaban:** Track Changes membantu meninjau perubahan sebelum disetujui.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Advanced — 5 Modul

### W-04-01 · Template surat terstandar dan Content Controls

**Tujuan:** Membuat template yang dapat diisi ulang tanpa mengubah format wajib.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** arsip, kepegawaian, akademik

**Materi inti:**

- Template memisahkan bagian tetap dari area input.
- Content Controls menyediakan kontrol isian tertentu pada Word desktop.
- Placeholder harus memiliki petunjuk yang tidak ambigu.
- Template perlu menjalani review dan persetujuan format oleh unit berwenang.
- Lakukan uji cetak dan uji pemakaian oleh staf berbeda.

**Langkah praktikum:**

1. Identifikasi elemen surat yang selalu tetap.
2. Rancang placeholder tanggal, nomor, perihal, dan penerima.
3. Gunakan Developer → Controls bila tersedia.
4. Batasi perubahan bagian tetap sesuai kebutuhan.
5. Uji template dengan tiga contoh isi berbeda.

**Contoh terarah:** Field: [Nomor Surat] [Tanggal] [Unit Tujuan] [Agenda] [Pejabat Berwenang]

**Tugas praktik peserta:** Rancang template undangan layanan yang dapat digunakan lima kegiatan berbeda.

**Keluaran yang harus dikumpulkan:** Template surat teruji dengan panduan isian

**Soal evaluasi mikro:** Keuntungan Content Controls pada template adalah ...

**Pembahasan jawaban:** Content Controls mengarahkan pengisian pada bagian yang ditentukan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-04-02 · Formulir layanan mahasiswa dan alur verifikasi

**Tujuan:** Merancang formulir administrasi yang ringkas dan meminimalkan pengumpulan data.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** kemahasiswaan, akademik

**Materi inti:**

- Minta hanya informasi yang benar-benar diperlukan untuk layanan.
- Buat instruksi dan field wajib yang jelas.
- Pisahkan bagian diisi pemohon dengan bagian verifikasi petugas.
- Gunakan kode proses, status, dan tanggal yang konsisten.
- Formulir tidak otomatis menjadi sah tanpa SOP serta persetujuan institusi.

**Langkah praktikum:**

1. Pilih kasus Permohonan Surat Keterangan.
2. Tentukan informasi minimum yang dibutuhkan.
3. Susun tabel Formulir dan kotak checklist persyaratan.
4. Tambahkan ruang nomor agenda serta status verifikasi petugas.
5. Uji keterbacaan pada cetak hitam putih.

**Contoh terarah:** Jenis layanan: Surat Keterangan; Status berkas: Lengkap/Perlu Perbaikan

**Tugas praktik peserta:** Buat formulir permohonan satu halaman menggunakan identitas contoh.

**Keluaran yang harus dikumpulkan:** Formulir layanan dengan instruksi dan alur verifikasi

**Soal evaluasi mikro:** Prinsip yang tepat saat merancang formulir mahasiswa adalah ...

**Pembahasan jawaban:** Minimisasi data mendukung layanan yang proporsional dan aman.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-04-03 · Cross-reference, sitasi internal, dan rujukan tabel

**Tujuan:** Menghubungkan bagian dokumen panjang agar nomor dan halaman tidak salah.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, keuangan

**Materi inti:**

- Cross-reference merujuk heading, gambar, atau caption secara dinamis.
- Caption yang konsisten mendukung daftar gambar/tabel.
- Referensi sumber data perlu jelas dan dapat ditelusuri.
- Field sebaiknya diperbarui sebelum cetak.
- Periksa apakah rujukan tetap benar setelah pengeditan atau pemindahan bagian.

**Langkah praktikum:**

1. Sisipkan dua tabel dengan caption pada laporan.
2. Tambahkan kalimat "Lihat Tabel 2" menggunakan Cross-reference.
3. Pindahkan salah satu tabel ke halaman lain.
4. Perbarui field Ctrl+A lalu F9 bila didukung.
5. Pastikan nomor pada naskah dan daftar tabel cocok.

**Contoh terarah:** "Realisasi program disajikan pada Tabel 2" menggunakan rujukan dinamis

**Tugas praktik peserta:** Susun laporan dengan dua referensi silang dan pemeriksaan nomor akhir.

**Keluaran yang harus dikumpulkan:** Dokumen dengan dua rujukan silang yang valid

**Soal evaluasi mikro:** Kelebihan Cross-reference dibanding mengetik nomor tabel manual adalah ...

**Pembahasan jawaban:** Cross-reference menggunakan field yang dapat diperbarui.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-04-04 · Metadata, redaksi, distribusi, dan arsip final

**Tujuan:** Meninjau informasi tersembunyi serta versi yang dibagikan kepada pihak berbeda.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** arsip, kepegawaian, kemahasiswaan

**Materi inti:**

- Dokumen dapat memiliki komentar, properties, dan informasi tersembunyi.
- Document Inspector membantu memeriksa jenis metadata tertentu.
- Redaksi data harus benar-benar menghapus informasi pada keluaran final.
- Klasifikasi akses dan retensi mengikuti instrumen kearsipan kampus.
- Jangan mengandalkan warna kotak hitam semata untuk menyembunyikan isi sensitif.

**Langkah praktikum:**

1. Buat salinan dokumen latihan, jangan memodifikasi sumber asli.
2. Periksa Properties dan Document Inspector pada Word desktop.
3. Hapus komentar/revisi yang tidak seharusnya dibagikan.
4. Buat versi distribusi tanpa data contoh yang sensitif.
5. Catat versi final, tanggal, serta penerima internal.

**Contoh terarah:** Sumber_Kepegawaian_v03.docx → Distribusi_Internal_v03.pdf

**Tugas praktik peserta:** Susun daftar pemeriksaan pradistribusi dokumen layanan.

**Keluaran yang harus dikumpulkan:** Checklist distribusi dan log versi

**Soal evaluasi mikro:** Langkah aman sebelum mendistribusikan dokumen final adalah ...

**Pembahasan jawaban:** Metadata serta akses perlu ditinjau di samping isi yang terlihat.

**Catatan kehati-hatian:** Lakukan redaksi dengan fitur serta prosedur yang benar dan uji hasilnya.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-04-05 · Dokumen batch dan standar audit revisi

**Tujuan:** Menata produksi dokumen berulang sekaligus menjaga jejak perubahan.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** arsip, akademik, kepegawaian

**Materi inti:**

- Dokumen batch memerlukan sumber data yang diverifikasi.
- Template, versi, dan pengesahan dipisahkan secara jelas.
- Log produksi mencatat jenis, nomor, waktu, dan status.
- Tiap file harus dapat ditelusuri ke sumbernya.
- Tinjauan sampling diperlukan untuk mendeteksi kesalahan yang berulang.

**Langkah praktikum:**

1. Siapkan lima data surat simulasi.
2. Gunakan template dan Mail Merge atau proses pengisian terkontrol.
3. Beri nama file berdasarkan nomor contoh dan unit.
4. Pilih dua sampel untuk audit isi serta format.
5. Buat daftar pengendalian revisi dan status persetujuan.

**Contoh terarah:** LOG: NOMOR, TANGGAL, JENIS SURAT, VERSI, PEMERIKSA, STATUS

**Tugas praktik peserta:** Buat prosedur produksi 10 surat layanan dengan titik kontrol wajib.

**Keluaran yang harus dikumpulkan:** SOP ringkas produksi batch dan daftar kontrol

**Soal evaluasi mikro:** Tujuan log produksi surat adalah ...

**Pembahasan jawaban:** Log produksi mendukung keterlacakan dan pengendalian mutu.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Expert — 5 Modul

### W-05-01 · Menulis SOP administrasi berbasis proses

**Tujuan:** Merancang SOP unit yang memuat tujuan, ruang lingkup, aktor, tahapan, dan kontrol.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** akademik, kemahasiswaan, arsip

**Materi inti:**

- SOP menjelaskan proses yang berlaku berdasarkan kewenangan institusi.
- Tiap tahapan harus memiliki pemicu, pelaksana, input, output, dan batas waktu jika sah.
- Bagan alur harus konsisten dengan langkah tertulis.
- Risiko serta titik kontrol ditandai secara eksplisit.
- Draf latihan tidak otomatis berlaku sebagai SOP kampus.

**Langkah praktikum:**

1. Pilih proses permohonan surat keterangan.
2. Gambarkan alur masuk–verifikasi–proses–serah hasil.
3. Tulis tabel langkah, pelaksana, dokumen, dan kontrol.
4. Definisikan kondisi permohonan tidak lengkap.
5. Lakukan simulasi satu kasus lalu revisi draf.

**Contoh terarah:** Alur: Terima permohonan → Verifikasi persyaratan → Proses → Periksa → Serahkan

**Tugas praktik peserta:** Buat draf SOP empat halaman untuk satu layanan administratif.

**Keluaran yang harus dikumpulkan:** Draf SOP dan diagram alur dengan catatan validasi

**Soal evaluasi mikro:** Komponen yang tepat pada satu langkah SOP adalah ...

**Pembahasan jawaban:** SOP yang dapat dipraktikkan menjelaskan siapa melakukan apa dan dengan bukti apa.

**Catatan kehati-hatian:** SOP harus ditinjau serta disahkan oleh otoritas institusi sebelum digunakan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-05-02 · Mengintegrasikan tabel dan grafik Excel ke laporan Word

**Tujuan:** Menyisipkan data analitik dengan pilihan tautan atau salinan yang tepat.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, keuangan

**Materi inti:**

- Paste Link mengikuti sumber yang tetap dapat diakses dan berisiko putus saat dipindahkan.
- Embed menyimpan salinan objek di dalam dokumen, tetapi ukurannya bisa lebih besar.
- Bila angka direvisi, sumber perlu diperiksa ulang.
- Berikan caption dan referensi untuk tabel/grafik.
- Pastikan pembaca memahami periode serta satuan.

**Langkah praktikum:**

1. Buat tabel KPI sintetis di Excel.
2. Copy grafik dan tempel sebagai objek atau gambar sesuai kebutuhan.
3. Uji apakah sumber tetap dapat diakses pada perangkat lain.
4. Beri caption dan cantumkan sumber periode.
5. Simpan dokumen final dan periksa tampilan cetak.

**Contoh terarah:** Sumber: Rekap Layanan Unit, Triwulan III 2026 (data simulasi)

**Tugas praktik peserta:** Buat laporan dua halaman dengan satu grafik dan satu tabel terverifikasi.

**Keluaran yang harus dikumpulkan:** Laporan dengan dua objek data dan catatan sumber

**Soal evaluasi mikro:** Risiko utama objek Word yang ditautkan ke Excel adalah ...

**Pembahasan jawaban:** Objek tertaut membutuhkan file sumber yang dapat diakses.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-05-03 · Aksesibilitas dan pemeriksaan dokumen panjang

**Tujuan:** Mengevaluasi dokumen yang siap dibaca secara digital maupun dicetak.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Heading terstruktur memudahkan pembaca dan teknologi asistif.
- Alt text penting pada gambar yang memiliki makna.
- Tabel memerlukan header serta urutan baca logis.
- Accessibility Checker membantu, tetapi tetap perlu review manual.
- Periksa konsistensi bahasa, daftar isi, caption, dan nomor halaman.

**Langkah praktikum:**

1. Buka laporan minimal enam halaman.
2. Gunakan Review → Check Accessibility jika tersedia.
3. Perbaiki heading bertingkat, alt text, dan label tabel.
4. Periksa daftar isi serta rujukan otomatis.
5. Lakukan uji cetak grayscale dan baca dari PDF.

**Contoh terarah:** Bukti audit: 0 heading loncat, semua gambar bermakna punya deskripsi

**Tugas praktik peserta:** Buat checklist aksesibilitas sepuluh poin untuk dokumen unit.

**Keluaran yang harus dikumpulkan:** Laporan yang ditinjau beserta hasil Accessibility Checker

**Soal evaluasi mikro:** Satu praktik yang meningkatkan aksesibilitas adalah ...

**Pembahasan jawaban:** Heading dan alternatif teks membantu navigasi dan pemahaman.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-05-04 · Tata naskah, klasifikasi, dan retensi arsip institusi

**Tujuan:** Mengenali hubungan dokumen Word dengan tata kelola arsip dinamis.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** arsip

**Materi inti:**

- Tata naskah dinas mengatur bentuk komunikasi resmi sesuai institusi.
- Klasifikasi arsip memudahkan penataan dan pencarian.
- Keamanan dan akses mengikuti klasifikasi yang ditetapkan.
- Jadwal retensi arsip (JRA) harus merujuk instrumen resmi yang berlaku.
- Jangan membuat angka retensi atau kewenangan pemusnahan dari asumsi pelatihan.

**Langkah praktikum:**

1. Pilih tiga jenis dokumen: surat tugas, berita acara, dan laporan.
2. Tentukan metadata minimum: nomor, tanggal, unit, perihal.
3. Cari kode klasifikasi dari pedoman unit bila tersedia.
4. Catat kebutuhan akses dan lokasi penyimpanan sesuai SOP.
5. Susun register arsip latihan tanpa menetapkan retensi sembarang.

**Contoh terarah:** Register contoh: ID, Kode Arsip [sesuai pedoman], Judul, Unit, Tanggal, Akses

**Tugas praktik peserta:** Buat tabel metadata untuk enam dokumen fiktif dan tiga aturan kontrol.

**Keluaran yang harus dikumpulkan:** Register arsip sintetis beserta referensi SOP institusi

**Soal evaluasi mikro:** Jangka retensi arsip resmi seharusnya diambil dari ...

**Pembahasan jawaban:** JRA merupakan acuan retensi yang perlu mengikuti prosedur pengesahan.

**Catatan kehati-hatian:** Contoh bukan pedoman resmi kearsipan atau rekomendasi pemusnahan.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### W-05-05 · Proyek akhir Word: paket dokumen administrasi

**Tujuan:** Menghasilkan berkas terstruktur yang konsisten, mudah diaudit, dan aman dibagikan.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Paket dokumen menyertakan sumber informasi, draf, hasil reviu, dan keluaran final.
- Setiap naskah memiliki pembaca dan fungsi yang jelas.
- Template dan Styles menjaga keseragaman.
- Pengendalian akses dan metadata diterapkan sebelum distribusi.
- Penilaian akhir meliputi ketepatan substansi serta teknis format.

**Langkah praktikum:**

1. Pilih satu unit serta kasus layanan pada halaman Studi Kasus.
2. Susun surat, notulen, formulir, dan laporan singkat.
3. Masukkan tabel atau grafik yang bersumber dari Excel.
4. Lakukan pemeriksaan revisi, rujukan, serta metadata.
5. Serahkan paket dengan daftar file dan checklist kualitas.

**Contoh terarah:** Paket: Surat.docx; Notulen.docx; Formulir.docx; Laporan.pdf; Log_Revisi.xlsx

**Tugas praktik peserta:** Buat empat jenis dokumen fiktif terkait satu kegiatan administrasi kampus.

**Keluaran yang harus dikumpulkan:** Empat dokumen, log reviu, dan checklist distribusi

**Soal evaluasi mikro:** Kriteria paket dokumen final yang baik adalah ...

**Pembahasan jawaban:** Dokumen yang siap audit memiliki jejak sumber dan revisi.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


# PowerPoint — 25 Modul

## Level Basic — 5 Modul

### P-01-01 · Tujuan paparan dan profil audiens unit kerja

**Tujuan:** Menetapkan pesan utama presentasi untuk pimpinan, kolega, atau peserta layanan.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Slide pendukung rapat pimpinan berbeda dari slide panduan petugas layanan.
- Tentukan keputusan atau tindakan yang diharapkan setelah paparan.
- Gunakan data yang sudah diverifikasi dari unit sumber.
- Setiap slide sebaiknya menyampaikan satu pesan utama.
- Hindari memuat data pribadi mahasiswa atau pegawai di layar publik.

**Langkah praktikum:**

1. Pilih audiens: kepala bagian layanan mahasiswa.
2. Tentukan tujuan: evaluasi waktu penyelesaian layanan.
3. Tulis tiga pertanyaan yang perlu dijawab presentasi.
4. Susun urutan awal, bukti, dan tindak lanjut.
5. Buat judul deck yang spesifik.

**Contoh terarah:** Judul: Evaluasi Ketepatan Layanan Administrasi — Triwulan III 2026

**Tugas praktik peserta:** Buat outline presentasi tiga slide untuk rapat koordinasi unit.

**Keluaran yang harus dikumpulkan:** Outline tiga slide dan profil audiens

**Soal evaluasi mikro:** Presentasi administrasi yang efektif dimulai dengan ...

**Pembahasan jawaban:** Tujuan dan audiens menentukan informasi yang perlu ditampilkan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-01-02 · Judul slide, hierarki teks, dan ringkasan pesan

**Tujuan:** Membedakan judul yang menjelaskan temuan dari judul generik.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** kemahasiswaan, perencanaan

**Materi inti:**

- Judul temuan mempercepat pembacaan tanpa narasi lisan.
- Teks singkat lebih cocok pada slide daripada paragraf laporan.
- Gunakan hierarki visual untuk judul, subjudul, dan catatan.
- Hindari ukuran huruf kecil dan elemen saling bertumpuk.
- Beri konteks waktu dan satuan pada angka.

**Langkah praktikum:**

1. Buka simulator PowerPoint dan buat slide baru.
2. Ganti judul generik menjadi satu temuan terukur.
3. Tuliskan tiga poin penjelasan maksimum.
4. Tambahkan satu catatan pembicara terpisah.
5. Periksa slide pada tampilan presentasi penuh.

**Contoh terarah:** Kurang: "Hasil" → Lebih baik: "72 dari 80 layanan diselesaikan pada periode contoh"

**Tugas praktik peserta:** Ubah tiga judul generik menjadi judul pesan berbasis data contoh.

**Keluaran yang harus dikumpulkan:** Tiga slide dengan judul temuan

**Soal evaluasi mikro:** Judul slide yang lebih informatif adalah ...

**Pembahasan jawaban:** Judul insight mengomunikasikan inti temuan dengan konteks.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-01-03 · Layout, warna, dan identitas presentasi institusi

**Tujuan:** Membuat slide konsisten tanpa meniru atau memalsukan identitas resmi.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Tema mengontrol warna serta jenis huruf secara konsisten.
- Ruang kosong membantu fokus pada konten.
- Kontras tinggi meningkatkan keterbacaan proyektor.
- Gunakan logo resmi hanya dari aset yang disediakan instansi dan dengan izin.
- Semua slide perlu margin serta hierarki yang konsisten.

**Langkah praktikum:**

1. Pilih salah satu tema simulator.
2. Atur layout cover dan content secara konsisten.
3. Gunakan maksimum dua keluarga font.
4. Pastikan judul tidak terpotong.
5. Uji pada lebar layar laptop dan layar proyektor.

**Contoh terarah:** Cover: Judul, Unit, Periode, Presenter; Konten: Insight, Grafik, Sumber

**Tugas praktik peserta:** Buat cover dan dua slide isi untuk laporan kinerja unit.

**Keluaran yang harus dikumpulkan:** Slide cover dan dua layout isi konsisten

**Soal evaluasi mikro:** Prinsip desain yang membantu keterbacaan adalah ...

**Pembahasan jawaban:** Kontras dan hierarki visual mendukung fokus pembaca.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-01-04 · Mengubah tabel menjadi narasi satu slide

**Tujuan:** Memilih angka paling penting dari laporan operasional tanpa mengubah makna.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** akademik, kemahasiswaan

**Materi inti:**

- Slide tidak harus memuat seluruh tabel detail.
- Pilih dua atau tiga angka inti yang relevan dengan keputusan.
- Selalu tampilkan periode, satuan, dan sumber data.
- Pisahkan fakta dari interpretasi.
- Jangan menampilkan persentase tanpa pembilang dan penyebut bila rawan salah tafsir.

**Langkah praktikum:**

1. Gunakan data contoh 80 permohonan, 72 selesai, 8 diproses.
2. Hitung proporsi selesai =72/80×100%.
3. Buat satu slide dengan angka kunci 90%.
4. Tambahkan satu pesan dan satu tindak lanjut.
5. Periksa kembali bahwa 72+8=80.

**Contoh terarah:** 90% layanan selesai · 8 permohonan dalam proses (data simulasi)

**Tugas praktik peserta:** Buat satu slide ringkasan jumlah layanan dan pesan tindak lanjut.

**Keluaran yang harus dikumpulkan:** Slide KPI satu halaman dengan sumber

**Soal evaluasi mikro:** Persentase 72 dari 80 adalah ...

**Pembahasan jawaban:** 72 dibagi 80 dikalikan 100% menghasilkan 90%.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-01-05 · Menyimpan deck, PDF, dan menjalankan slideshow

**Tujuan:** Melakukan pemeriksaan presentasi sebelum rapat atau pelatihan.

**Estimasi:** 14 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- PPTX adalah berkas PowerPoint yang dapat diedit.
- PDF cocok untuk salinan distribusi noninteraktif.
- F5 umumnya memulai slideshow dari slide pertama di PowerPoint desktop Windows.
- Esc menghentikan tayangan.
- Media, font, dan format perlu diuji pada perangkat presentasi.

**Langkah praktikum:**

1. Simpan deck sebagai Paparan_Unit_v01.pptx.
2. Coba Slide Show → From Beginning.
3. Gunakan tombol panah untuk berpindah slide.
4. Akhiri slideshow dan periksa urutan slide.
5. Ekspor PDF dan cek setiap halaman.

**Contoh terarah:** File final: 2026_PaparanKinerja_Unit_v02.pptx + salinan PDF

**Tugas praktik peserta:** Jalankan tayangan tiga slide dalam simulator lalu jelaskan tombol navigasinya.

**Keluaran yang harus dikumpulkan:** Deck yang dapat dipresentasikan dan diperiksa

**Soal evaluasi mikro:** Shortcut yang umum dipakai memulai tayangan dari awal adalah ...

**Pembahasan jawaban:** F5 memulai slideshow pada PowerPoint desktop Windows.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Beginner — 5 Modul

### P-02-01 · Paparan alur layanan mahasiswa

**Tujuan:** Membuat alur layanan yang mudah dipahami pemohon dan petugas.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** kemahasiswaan, akademik

**Materi inti:**

- Gunakan empat tahap maksimum pada satu visual utama.
- Bedakan langkah pemohon dari langkah verifikator.
- Tampilkan dokumen yang harus dibawa tanpa memuat data pribadi.
- Berikan titik bantuan bila permohonan belum lengkap.
- Cantumkan bahwa diagram adalah latihan dan perlu disesuaikan SOP.

**Langkah praktikum:**

1. Buat slide 1: tujuan layanan.
2. Slide 2: alur Permohonan→Verifikasi→Proses→Selesai.
3. Slide 3: dokumen dan kontak unit.
4. Gunakan ikon sederhana serta teks ringkas.
5. Latih menjelaskan alur dalam dua menit.

**Contoh terarah:** Alur layanan contoh: Isi formulir → Verifikasi petugas → Proses dokumen → Penyerahan

**Tugas praktik peserta:** Buat tiga slide untuk orientasi petugas layanan mahasiswa.

**Keluaran yang harus dikumpulkan:** Tiga slide alur layanan dan catatan pembicara

**Soal evaluasi mikro:** Diagram alur yang membantu staf layanan harus ...

**Pembahasan jawaban:** Tahap, tanggung jawab, dan hasil mendukung konsistensi pelayanan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-02-02 · Paparan anggaran: pagu, realisasi, dan deviasi

**Tujuan:** Menyajikan perbandingan rupiah secara jelas untuk rapat monitoring.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** keuangan, perencanaan

**Materi inti:**

- Pagu adalah batas/anggaran yang digunakan pada kasus simulasi.
- Realisasi adalah nilai yang sudah tercatat dalam data sumber.
- Selisih = pagu − realisasi.
- Tampilkan satuan rupiah dan periode di setiap visual.
- Jangan menyimpulkan efektivitas hanya dari persentase serapan.

**Langkah praktikum:**

1. Siapkan data Pagu Rp50 juta, Realisasi Rp42 juta.
2. Hitung selisih Rp8 juta serta serapan 84%.
3. Buat dua kolom perbandingan.
4. Tambahkan catatan konteks dan kebutuhan tindak lanjut.
5. Pastikan seluruh angka sesuai data sumber.

**Contoh terarah:** Pagu Rp50.000.000; Realisasi Rp42.000.000; Serapan 84%; Selisih Rp8.000.000

**Tugas praktik peserta:** Buat satu slide ringkasan anggaran fiktif untuk rapat unit.

**Keluaran yang harus dikumpulkan:** Slide perbandingan anggaran dan interpretasi berimbang

**Soal evaluasi mikro:** Jika pagu 50 juta dan realisasi 42 juta, serapannya ...

**Pembahasan jawaban:** 42/50×100%=84%.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-02-03 · Timeline kegiatan dan milestone perencanaan

**Tujuan:** Menyusun urutan kegiatan, penanggung jawab, dan tanggal target.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** perencanaan, kepegawaian

**Materi inti:**

- Timeline memvisualisasikan urutan, bukan menggantikan rincian jadwal.
- Milestone menandai keluaran terukur.
- Tenggat perlu realistis dan sesuai persetujuan pihak terkait.
- Dependensi antaraktivitas perlu ditampilkan secara jelas.
- Status selesai harus didukung bukti, bukan asumsi.

**Langkah praktikum:**

1. Pilih kasus persiapan pelatihan administrasi.
2. Tentukan tahapan Persiapan, Pelaksanaan, Evaluasi.
3. Tambahkan tanggal contoh dan PIC per tahap.
4. Tampilkan tiga milestone pada slide.
5. Periksa dependensi dan rentang tanggal.

**Contoh terarah:** 19 Nov: finalisasi materi; 23 Nov: pelatihan; 27 Nov: evaluasi (contoh)

**Tugas praktik peserta:** Buat slide timeline satu bulan beserta tiga milestone.

**Keluaran yang harus dikumpulkan:** Slide timeline beserta PIC dan keluaran

**Soal evaluasi mikro:** Milestone adalah ...

**Pembahasan jawaban:** Milestone menandai hasil atau titik penting suatu rencana.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-02-04 · Tabel, gambar, dan sumber visual

**Tujuan:** Memakai tabel atau gambar untuk membantu penjelasan tanpa mengaburkan sumber.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Tabel cocok untuk perbandingan terbatas dengan sedikit baris.
- Gambar perlu relevan dan berizin.
- Diagram sebaiknya memiliki label serta urutan baca.
- Cantumkan sumber atau keterangan dataset contoh.
- Periksa apakah visual masih terbaca di proyektor.

**Langkah praktikum:**

1. Pilih satu tabel dengan tiga baris indikator.
2. Tulis judul temuan di atas tabel.
3. Berikan sorotan pada satu angka kunci.
4. Tambahkan sumber data dan periode.
5. Buat versi grafik jika tabel terlalu padat.

**Contoh terarah:** Indikator: Permohonan masuk 80; selesai 72; tertunda 8

**Tugas praktik peserta:** Buat dua slide: satu tabel dan satu visualisasi alternatif.

**Keluaran yang harus dikumpulkan:** Dua slide visual data dan catatan sumber

**Soal evaluasi mikro:** Penggunaan sumber data pada slide membantu ...

**Pembahasan jawaban:** Sumber memungkinkan audiens memeriksa dasar klaim.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-02-05 · Speaker Notes dan latihan pemaparan

**Tujuan:** Memakai catatan pembicara agar slide ringkas tetapi penjelasan lengkap.

**Estimasi:** 18 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** semua

**Materi inti:**

- Notes tidak harus tampil pada layar audiens.
- Catatan memuat definisi indikator atau transisi narasi.
- Presenter View dapat menampilkan notes jika perangkat mendukung.
- Durasi latihan membantu menekan paparan melebar.
- Siapkan jawaban atas satu pertanyaan kritis per slide.

**Langkah praktikum:**

1. Buka salah satu slide di simulator.
2. Isi notes dengan penjelasan satu paragraf.
3. Jalankan mode Presentasikan dan gunakan tombol panah.
4. Latih narasi selama satu menit per slide.
5. Catat satu titik yang perlu disederhanakan.

**Contoh terarah:** Notes: "Capaian 90% dihitung dari 72 layanan selesai dibagi 80 masuk."

**Tugas praktik peserta:** Buat notes untuk tiga slide dan lakukan simulasi presentasi lima menit.

**Keluaran yang harus dikumpulkan:** Catatan pembicara dan rencana durasi

**Soal evaluasi mikro:** Manfaat utama Speaker Notes adalah ...

**Pembahasan jawaban:** Notes mendukung narasi tanpa memenuhi layar dengan teks.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Intermediate — 5 Modul

### P-03-01 · Storyline masalah–bukti–opsi–rekomendasi

**Tujuan:** Mengubah data operasional menjadi presentasi yang mendukung keputusan.

**Estimasi:** 22 menit · **Kebutuhan:** dapat dimulai pada webapp · **Unit yang terkait:** perencanaan, kemahasiswaan

**Materi inti:**

- Mulailah dari isu terukur dan mengapa isu itu penting.
- Gunakan bukti dari sumber yang dapat diverifikasi.
- Pisahkan analisis penyebab dari dugaan yang belum diuji.
- Paparkan beberapa opsi dengan konsekuensi yang seimbang.
- Akhiri dengan keputusan atau tindak lanjut yang jelas.

**Langkah praktikum:**

1. Tulis isu "8 dari 80 layanan belum selesai".
2. Sajikan bukti jumlah per tahap proses.
3. Tentukan dua penjelasan yang perlu diverifikasi.
4. Bandingkan dua opsi perbaikan dengan kebutuhan sumber daya.
5. Buat slide keputusan dan PIC tindak lanjut.

**Contoh terarah:** Urutan 5 slide: Isu → Fakta → Analisis → Opsi → Rekomendasi

**Tugas praktik peserta:** Susun paparan lima slide dari satu kasus keterlambatan layanan.

**Keluaran yang harus dikumpulkan:** Deck lima slide dengan rekomendasi yang bisa ditindaklanjuti

**Soal evaluasi mikro:** Perbedaan fakta dan asumsi dalam paparan ialah ...

**Pembahasan jawaban:** Pemilahan fakta dan hipotesis mencegah kesimpulan prematur.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-03-02 · Membuat grafik dari Excel dengan satuan jelas

**Tujuan:** Menyisipkan grafik dengan nilai dan sumber yang konsisten terhadap workbook.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, akademik, kepegawaian

**Materi inti:**

- Grafik garis cocok untuk tren waktu; kolom untuk perbandingan kategori.
- Tulis judul berupa temuan, bukan "Grafik".
- Pastikan rentang dan angka sesuai workbook sumber.
- Tampilkan satuan rupiah/orang dan periode.
- Beri catatan bila pembanding mengalami perubahan definisi.

**Langkah praktikum:**

1. Buat grafik bulanan dalam Excel sintetis.
2. Salin ke PowerPoint memakai opsi paste yang sesuai.
3. Atur judul, sumbu, dan label.
4. Bandingkan satu titik data dengan Excel.
5. Sertakan sumber dan tanggal pengambilan data.

**Contoh terarah:** Jan 55, Feb 62, Mar 73 layanan selesai (data simulasi)

**Tugas praktik peserta:** Buat slide tren tiga bulan dengan interpretasi dan verifikasi satu titik.

**Keluaran yang harus dikumpulkan:** Slide tren dengan sumber dan angka terverifikasi

**Soal evaluasi mikro:** Untuk menggambarkan perubahan dari Januari ke Maret, grafik yang lazim adalah ...

**Pembahasan jawaban:** Grafik garis lazim untuk memperlihatkan deret waktu.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-03-03 · Slide Master dan layout presentasi unit

**Tujuan:** Membuat template unit yang konsisten sepanjang rapat berkala.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, semua

**Materi inti:**

- Slide Master mengatur elemen bersama di berbagai layout.
- Sediakan cover, pemisah bab, slide data, dan penutup.
- Footer dapat memuat unit serta periode, tidak perlu data rahasia.
- Gunakan warna identitas yang disahkan institusi.
- Periksa apakah layout bekerja untuk judul panjang dan tabel.

**Langkah praktikum:**

1. Buka View → Slide Master pada PowerPoint desktop.
2. Rancang empat layout dengan margin konsisten.
3. Atur footer Unit, Periode, dan nomor slide.
4. Uji pemakaian pada lima slide berbeda.
5. Simpan sebagai template yang dapat diperbarui.

**Contoh terarah:** Layout: Cover / KPI / Dua Kolom / Tindak Lanjut

**Tugas praktik peserta:** Buat master presentasi laporan bulanan dengan empat layout.

**Keluaran yang harus dikumpulkan:** Template PowerPoint empat layout teruji

**Soal evaluasi mikro:** Fungsi Slide Master adalah ...

**Pembahasan jawaban:** Slide Master menjaga elemen desain bersama di banyak slide.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-03-04 · Navigasi hyperlink dan menu paparan

**Tujuan:** Membuat navigasi pada deck yang memiliki bagian laporan berbeda.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Hyperlink dapat membuka slide tertentu di deck.
- Menu perlu memiliki tombol kembali yang jelas.
- Tautan eksternal berisiko tidak tersedia offline.
- Latih semua cabang navigasi sebelum acara.
- Sediakan urutan linear sebagai cadangan.

**Langkah praktikum:**

1. Buat slide menu Unit Keuangan / Layanan / Arsip.
2. Hubungkan tiap tombol ke slide bagiannya.
3. Tambahkan tombol Kembali ke Menu.
4. Jalankan Slide Show dan uji semua tombol.
5. Catat jika ada tautan rusak atau tidak sesuai.

**Contoh terarah:** Menu utama → Subbagian A/B/C → Kembali ke menu utama

**Tugas praktik peserta:** Buat prototype deck bercabang tiga bagian dan uji semuanya.

**Keluaran yang harus dikumpulkan:** Deck navigasi dengan empat tautan teruji

**Soal evaluasi mikro:** Komponen wajib untuk navigasi bercabang adalah ...

**Pembahasan jawaban:** Tombol kembali menjaga jalur presentasi dapat digunakan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-03-05 · Review konten dan latihan menjawab pertanyaan

**Tujuan:** Menyiapkan paparan yang dapat dipertanggungjawabkan ketika pimpinan meminta rincian.

**Estimasi:** 22 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Sumber angka harus bisa ditunjukkan saat diminta.
- Pisahkan angka yang sudah final dari estimasi.
- Speaker Notes memuat definisi dan catatan keterbatasan.
- Gunakan appendix untuk rincian yang tidak cocok di slide utama.
- Tinjau paparan bersama rekan sebelum digunakan secara operasional.

**Langkah praktikum:**

1. Pilih satu deck ringkasan unit.
2. Buat daftar lima pertanyaan yang mungkin muncul.
3. Tambahkan tiga lampiran data sumber atau definisi.
4. Periksa istilah serta angka yang tidak konsisten.
5. Latih jawaban berbasis data tanpa menebak.

**Contoh terarah:** Pertanyaan: Mengapa capaian 90%, apa pembaginya, dan siapa memvalidasi?

**Tugas praktik peserta:** Buat daftar Q&A untuk paparan kinerja layanan dengan bukti pendukung.

**Keluaran yang harus dikumpulkan:** Deck dan daftar pertanyaan–jawaban berbukti

**Soal evaluasi mikro:** Jika angka diminta tetapi belum terverifikasi, sebaiknya ...

**Pembahasan jawaban:** Ketepatan dan keterbatasan data perlu diungkapkan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Advanced — 5 Modul

### P-04-01 · Dashboard satu slide untuk rapat monitoring

**Tujuan:** Menyajikan target, realisasi, tren, dan risiko sebagai satu halaman eksekutif.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, kemahasiswaan, keuangan

**Materi inti:**

- Pilih maksimal beberapa KPI utama sesuai keputusan rapat.
- Tunjukkan angka target, realisasi, dan selisih dengan satuan.
- Ringkas tren waktu dan anomali secara visual.
- Hindari indikator yang definisinya berubah tanpa catatan.
- Berikan ruang khusus tindak lanjut.

**Langkah praktikum:**

1. Pilih tiga KPI pada dataset latihan.
2. Rancang layout 3 kartu dan satu grafik tren.
3. Cantumkan periode serta keterangan sumber.
4. Tambahkan satu catatan risiko dan usulan tindak lanjut.
5. Lakukan review keterbacaan pada tampilan proyektor.

**Contoh terarah:** KPI 1: 90% layanan selesai; KPI 2: 8 tertunda; KPI 3: 2 hari median proses (simulasi)

**Tugas praktik peserta:** Buat dashboard satu slide dan jelaskan dua keputusan yang didukungnya.

**Keluaran yang harus dikumpulkan:** Satu slide dashboard teruji

**Soal evaluasi mikro:** Dashboard satu slide sebaiknya berisi ...

**Pembahasan jawaban:** Ringkasan eksekutif mengutamakan sinyal yang relevan terhadap keputusan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-04-02 · Membaca variansi, tren, dan penyebab yang diuji

**Tujuan:** Menyampaikan deviasi kinerja secara jujur tanpa menyimpulkan sebab secara prematur.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** keuangan, perencanaan, kepegawaian

**Materi inti:**

- Variansi menunjukkan perbedaan dua nilai yang terdefinisi.
- Tren membutuhkan data beberapa periode yang sebanding.
- Satu anomali tidak membuktikan hubungan sebab-akibat.
- Tandai perubahan proses pencatatan agar perbandingan adil.
- Pisahkan pernyataan observasi, hipotesis, dan tindakan.

**Langkah praktikum:**

1. Buat grafik capaian layanan selama empat bulan.
2. Tandai bulan yang turun secara signifikan.
3. Tulis observasi berdasarkan angka.
4. Cantumkan kemungkinan faktor yang perlu diuji.
5. Ajukan data tambahan untuk memeriksa hipotesis.

**Contoh terarah:** April 82%, Mei 88%, Juni 77%, Juli 90% (contoh)

**Tugas praktik peserta:** Buat satu slide tren dengan anotasi dan dua pertanyaan investigasi.

**Keluaran yang harus dikumpulkan:** Slide analisis tren dan daftar verifikasi

**Soal evaluasi mikro:** Dari satu penurunan kinerja, kesimpulan yang benar adalah ...

**Pembahasan jawaban:** Deskripsi tren tidak dengan sendirinya membuktikan kausalitas.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-04-03 · Objek linked vs embedded dari Excel

**Tujuan:** Memilih cara menyalin grafik atau tabel agar pembaruan dan distribusi terkontrol.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, keuangan

**Materi inti:**

- Linked object mengikuti sumber yang tersedia dan relatif memudahkan pembaruan.
- Embedded object menyimpan salinan dalam presentasi.
- Pemindahan lokasi sumber dapat memutus tautan.
- Ekspor PDF tidak mempertahankan pengeditan objek sumber.
- Pilih teknik berdasarkan kebutuhan distribusi dan kontrol informasi.

**Langkah praktikum:**

1. Siapkan grafik dari Excel latihan.
2. Tempel salinan linked dan embedded pada dua slide.
3. Ubah satu angka sumber lalu tinjau efek pada kedua slide.
4. Simulasikan perpindahan folder sumber.
5. Buat catatan kapan memilih linked atau embedded.

**Contoh terarah:** Linked: dipakai internal saat sumber terjaga; Embedded: salinan untuk distribusi terkontrol

**Tugas praktik peserta:** Tulis matriks keputusan penggunaan linked/embedded untuk laporan bulanan.

**Keluaran yang harus dikumpulkan:** Perbandingan dua metode beserta rekomendasi

**Soal evaluasi mikro:** Risiko khas linked chart ialah ...

**Pembahasan jawaban:** Tautan membutuhkan sumber yang masih dapat diakses.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-04-04 · Custom Show untuk rapat lintas bagian

**Tujuan:** Membuat satu deck dengan jalur khusus pimpinan dan jalur teknis.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, semua

**Materi inti:**

- Custom Slide Show dapat memilih subset slide untuk audiens tertentu.
- Jalur eksekutif memerlukan ringkasan keputusan lebih singkat.
- Jalur teknis memerlukan rincian perhitungan serta prosedur.
- Menu navigasi dan halaman kembali perlu diuji.
- Catatan pembicara dapat berbeda menurut audiens.

**Langkah praktikum:**

1. Bangun deck delapan slide tentang kinerja unit.
2. Tandai empat slide untuk pimpinan dan empat untuk teknis.
3. Buat dua Custom Shows di PowerPoint desktop.
4. Uji alur dari awal sampai akhir.
5. Catat slide bersama dan lampiran yang dipakai.

**Contoh terarah:** Jalur pimpinan: 1,2,5,8; jalur teknis: 1,3,4,6,7,8

**Tugas praktik peserta:** Rancang dua jalur presentasi dari satu data kinerja administrasi.

**Keluaran yang harus dikumpulkan:** Deck dengan dua jalur dan catatan uji

**Soal evaluasi mikro:** Kegunaan Custom Show yang tepat adalah ...

**Pembahasan jawaban:** Custom Show memisahkan alur berdasarkan tujuan audiens.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-04-05 · Audit aksesibilitas, notes, dan cadangan offline

**Tujuan:** Menyiapkan presentasi operasional yang andal pada berbagai perangkat.

**Estimasi:** 28 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Kontras dan ukuran teks harus terbaca dari belakang ruangan.
- Alt text dan reading order membantu aksesibilitas.
- Video dan tautan harus diuji tanpa internet bila diperlukan.
- Sediakan salinan PDF ketika fitur interaktif tidak tersedia.
- Jangan menayangkan daftar personal yang tidak sesuai kewenangan.

**Langkah praktikum:**

1. Jalankan Accessibility Checker pada PowerPoint desktop.
2. Periksa ukuran huruf, alt text, dan sumber gambar.
3. Tes semua tautan serta media offline.
4. Simpan PPTX dan PDF ke media cadangan sesuai kebijakan.
5. Latih pembukaan presentasi pada komputer lain.

**Contoh terarah:** Checklist: font, media, tautan, notes, data pribadi, PDF cadangan

**Tugas praktik peserta:** Buat checklist prapresentasi 10 butir untuk kegiatan kampus.

**Keluaran yang harus dikumpulkan:** Checklist uji dan dua format keluaran

**Soal evaluasi mikro:** Cadangan yang tepat untuk presentasi penting adalah ...

**Pembahasan jawaban:** Pemeriksaan offline mencegah kegagalan karena font, media, dan konektivitas.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


## Level Expert — 5 Modul

### P-05-01 · Executive briefing hasil kinerja untuk pimpinan

**Tujuan:** Menyusun paparan ringkas dengan angka, risiko, opsi, dan kebutuhan keputusan.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, kemahasiswaan, keuangan

**Materi inti:**

- Pimpinan membutuhkan keputusan, bukan daftar semua kegiatan.
- Gunakan headline berupa temuan yang didukung data.
- Tunjukkan perbandingan terhadap target serta sumbernya.
- Tampilkan risiko, opsi, trade-off, dan rekomendasi.
- Akhiri dengan tindakan, pemilik tugas, dan waktu yang diusulkan.

**Langkah praktikum:**

1. Pilih kasus layanan mahasiswa atau keuangan.
2. Tulis executive summary satu slide.
3. Sajikan dua grafik yang telah diverifikasi.
4. Buat matriks dua opsi tindak lanjut.
5. Tutup dengan keputusan yang diminta dan PIC.

**Contoh terarah:** Keputusan diminta: prioritas perbaikan alur verifikasi berkas pada periode berikut

**Tugas praktik peserta:** Buat briefing enam slide untuk pimpinan yang memakai data sintetis.

**Keluaran yang harus dikumpulkan:** Briefing enam slide dan ringkasan keputusan

**Soal evaluasi mikro:** Elemen penutup executive briefing yang kuat adalah ...

**Pembahasan jawaban:** Paparan keputusan harus jelas menyatakan tindak lanjut yang diperlukan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-05-02 · Paparan opsi perencanaan dan anggaran lintas unit

**Tujuan:** Membuat analisis skenario yang transparan untuk forum perencanaan.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** perencanaan, keuangan

**Materi inti:**

- Bandingkan skenario dengan asumsi dan periode yang sama.
- Sajikan biaya, manfaat, risiko, dan keterbatasan tiap opsi.
- Jangan menukar nilai asumsi menjadi klaim realisasi.
- Grafik perlu baseline dan satuan.
- Rekomendasi didukung parameter yang dapat diperiksa.

**Langkah praktikum:**

1. Siapkan tiga skenario pagu operasional pada Excel.
2. Hitung biaya dan sasaran tiap skenario.
3. Buat grafik perbandingan sederhana.
4. Tulis rekomendasi dan batas asumsinya.
5. Siapkan appendix rumus serta sumber angka.

**Contoh terarah:** Skenario Dasar 50 juta; Efisiensi 45 juta; Ekspansi 58 juta (simulasi)

**Tugas praktik peserta:** Susun presentasi 8 slide rapat perencanaan dengan matriks opsi.

**Keluaran yang harus dikumpulkan:** Paparan skenario dan lampiran perhitungan

**Soal evaluasi mikro:** Analisis skenario yang valid memerlukan ...

**Pembahasan jawaban:** Kesetaraan asumsi memungkinkan perbandingan yang relevan.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-05-03 · Deck panduan SOP layanan dan kearsipan

**Tujuan:** Mentransformasikan SOP terstruktur menjadi media penguatan keterampilan pegawai.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** arsip, akademik, kemahasiswaan

**Materi inti:**

- Slide prosedur tidak boleh bertentangan dengan naskah SOP disahkan.
- Tampilkan peran, dokumen, tahap, dan pengecualian.
- Berikan skenario tidak lengkap untuk latihan pengambilan keputusan.
- Pisahkan instruksi kerja dari materi diskusi.
- Sertakan versi SOP dan lokasi rujukan resmi.

**Langkah praktikum:**

1. Gunakan draf SOP pelayanan administrasi sintetis.
2. Pecah alur ke empat slide tahap proses.
3. Buat satu slide kesalahan umum dan kontrol.
4. Susun dua pertanyaan latihan untuk peserta.
5. Cocokkan istilah dengan naskah SOP sumber.

**Contoh terarah:** Skenario: berkas pemohon tidak lengkap → status Perlu Perbaikan → kembali setelah verifikasi

**Tugas praktik peserta:** Buat lima slide orientasi SOP yang menyertakan dua skenario kasus.

**Keluaran yang harus dikumpulkan:** Deck sosialisasi SOP dan daftar verifikasi sumber

**Soal evaluasi mikro:** Yang menjadi rujukan utama prosedur kerja ialah ...

**Pembahasan jawaban:** Slide adalah media pengajaran, bukan pengganti SOP resmi.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-05-04 · Design system dan template paparan institusi

**Tujuan:** Membuat aset visual reusable untuk semua unit tanpa mengorbankan keseragaman.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Design system menetapkan tipografi, palet, grid, ruang, dan komponen.
- Template wajib mencakup layout ringkasan, data, alur, dan penutup.
- Kontras serta ukuran konten diuji pada perangkat sasaran.
- Logo serta identitas visual mengikuti pedoman yang berlaku.
- Buat halaman panduan pemakaian dan contoh kesalahan.

**Langkah praktikum:**

1. Tetapkan palet tiga warna dari pedoman institusi yang tersedia.
2. Rancang master lima layout.
3. Buat contoh KPI, tabel, diagram alur, dan timeline.
4. Tulis instruksi penggunaan satu halaman.
5. Uji file pada laptop dan proyektor berbeda.

**Contoh terarah:** Layout: Cover, Section, KPI, Process, Decision

**Tugas praktik peserta:** Buat template tim administrasi yang dapat digunakan untuk berbagai kegiatan.

**Keluaran yang harus dikumpulkan:** Template lima layout dan panduan penggunaan

**Soal evaluasi mikro:** Keunggulan design system adalah ...

**Pembahasan jawaban:** Design system menstandarkan penyajian, bukan memvalidasi substansi.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.

### P-05-05 · Proyek akhir PowerPoint: paparan capaian unit

**Tujuan:** Menghasilkan deck yang siap ditinjau dengan narasi, sumber data, dan tindak lanjut.

**Estimasi:** 35 menit · **Kebutuhan:** Microsoft Office desktop untuk praktik fitur penuh · **Unit yang terkait:** semua

**Materi inti:**

- Paket akhir menggunakan data fiktif terverifikasi dari latihan Excel.
- Narasi harus memisahkan fakta dan usulan.
- Slide memiliki sumber serta periode yang jelas.
- Alur presentasi diuji dari awal sampai akhir.
- Ada cadangan PDF, notes, dan log pemeriksaan.

**Langkah praktikum:**

1. Pilih satu studi kasus unit kerja dari OfficeLab.
2. Susun outline 8–10 slide beserta audiens.
3. Buat grafik dan tabel dari rekap Excel sintetis.
4. Tambahkan keputusan yang diusulkan dan tahapan implementasi.
5. Uji tampilan, navigasi, dan backup offline.

**Contoh terarah:** Keluaran: PaparanKinerja.pptx, PDF ringkasan, sumber data, checklist uji

**Tugas praktik peserta:** Buat presentasi kinerja unit 8–10 slide dan sampaikan simulasi lima menit.

**Keluaran yang harus dikumpulkan:** Deck akhir, sumber, dan checklist pengujian

**Soal evaluasi mikro:** Ciri paparan akhir yang siap diuji adalah ...

**Pembahasan jawaban:** Keandalan dan akuntabilitas penting dalam laporan kinerja unit.

**Catatan kehati-hatian:** Jangan menganggap latihan ini sebagai format atau SOP resmi institusi. Periksa kecocokan fitur dengan versi Office yang digunakan peserta.

**Rubrik pemeriksaan modul:**

- Ketepatan: isi, angka, tanggal, dan istilah sesuai dengan data contoh yang diberikan.
- Kelengkapan: seluruh langkah, bukti kerja, dan keluaran latihan tersedia.
- Keterlacakan: sumber, satuan, periode, nama file, dan revisi bisa diperiksa kembali.
- Keamanan: tidak menggunakan data pribadi/rahasia asli tanpa otorisasi.


---
Setiap tugas dan pertanyaan berada pula di dalam aplikasi offline agar dapat diikuti secara interaktif.
