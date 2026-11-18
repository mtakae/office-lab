# OfficeLab v2.1.2 — Halaman About dan profil pengembang

**Yang baru:** menu **Tentang OfficeLab** pada sidebar dan tombol ⓘ di header menampilkan informasi aplikasi, sasaran pelatihan, manfaat, langkah penggunaan, tujuh jalur unit kerja, lima level, privasi, serta identitas pengembang dengan foto yang disertakan. Pembaruan ini mempertahankan 75 modul, seluruh simulasi, navigasi sidebar yang bisa di-scroll, dan penyimpanan progres versi 2.1.1.

**Pengembang:** Mahardika I. Takaendengan, SST., M.Si · Program Studi Sistem Informasi - FMIPA UNSRAT · mahardika@unsrat.ac.id.

**Menggunakan HTML offline:** jalankan `OfficeLab_v2_1_2_Offline.html` langsung di Chrome, Edge, atau Firefox. Foto pengembang sudah tertanam di dalam HTML, sehingga tidak perlu internet maupun file gambar terpisah.

**Menggunakan versi modular:** buka `OfficeLab_v2_1_2/index.html`; pertahankan folder `assets/` bersama berkas JS/CSS. Foto ada di `assets/developer-profile.png`.

**Catatan:** konten adalah media pendidikan mandiri dan tidak mewakili pengesahan atau kebijakan resmi Microsoft maupun universitas. Data praktikum bersifat sintetis. Hindari pengisian informasi rahasia. Simpan cadangan progres dari menu Pengaturan sebelum pembersihan penyimpanan browser.

---

# OfficeLab v2.1.1 — Perbaikan navigasi sidebar

**Pembaruan dari v2.1:** Sidebar pada desktop kini dapat digulir secara vertikal secara mandiri menggunakan mouse wheel, trackpad, scroll bar, atau keyboard ketika sidebar difokuskan. Seluruh menu dan indikator progres dapat dijangkau pada layar yang pendek. Tidak ada perubahan pada kurikulum, dataset, kuis, atau format progres yang tersimpan di browser.

**Cara pakai:** Buka `index.html` (modular) atau file `OfficeLab_v2_1_1_Offline.html` (satu file). Jika halaman v2.1 lama masih terbuka, tutup tab lama dan buka file ini. Di desktop, arahkan kursor ke sidebar lalu gulir ke bawah. Pada layar kecil, menu bagian tetap dapat digulir mendatar di baris navigasi atas.

---

# OfficeLab v2.1 — Pelatihan Administrasi Kampus

OfficeLab adalah aplikasi pembelajaran lokal **Microsoft Excel, Word dan PowerPoint** untuk **tenaga kependidikan/pegawai administratif perguruan tinggi**, bukan sistem pengelolaan data kampus produksi. Aplikasi menggunakan **data fiktif/sintetis** dan dapat dijalankan tanpa jaringan, login, dan server.

## Mulai cepat

1. Untuk penggunaan biasa, buka `OfficeLab_v2_1_Offline.html` dengan Chrome, Edge, atau Firefox.
2. Pilih aplikasi (Excel / Word / PowerPoint) serta level (Basic / Beginner / Intermediate / Advanced / Expert).
3. Pilih **Materi** untuk mempelajari modul; setiap modul berisi tujuan, konsep, langkah praktikum, contoh, latihan, pertanyaan mikro, keluaran, dan rubrik.
4. Pilih **Jalur Unit Kerja** untuk studi kasus Keuangan, Kemahasiswaan, Perencanaan, Persuratan & Arsip, Kepegawaian, Layanan Akademik, atau Pengadaan & Aset.
5. Lanjutkan ke **Praktik** dan kerjakan simulator. Untuk latihan yang memerlukan fitur Power Query, Mail Merge, PivotTable, Slide Master, VBA, dan sejenisnya gunakan Microsoft Office desktop yang sesuai.
6. Ikuti **Evaluasi per level** (5 pertanyaan per level; ambang nilai 80%). Selesaikan kelima modul dan latihan mikro untuk menandai satu level tuntas.
7. Gunakan **Pengaturan → Ekspor Progres** untuk menyimpan cadangan JSON, atau impor cadangan jika berpindah perangkat/browser.

## Keterangan paket

- `OfficeLab_v2_1_Offline.html`: keseluruhan webapp dalam satu berkas, tanpa CDN, server atau dependensi eksternal.
- `index.html`, `style.css`, `cases.js`, `data.js`, `curriculum.js`, `app.js`: sumber kode webapp modular.
- `datasets/`: 7 berkas CSV sintetis untuk dibuka di Microsoft Excel.
- `SILABUS_TERPERINCI_75_MODUL.md`: uraian semua 75 modul, termasuk langkah praktik dan tugas.
- `PANDUAN_FASILITATOR.md`: panduan desain pelatihan, alur pembelajaran, rubrik, dan pengamanan data.
- `PANDUAN_PELATIHAN_CETAK.html`: lembar ringkasan kurikulum dan kegiatan untuk dicetak.
- `Template_Rekap_Penilaian.csv`: header administrasi penilaian peserta untuk fasilitator.

## Struktur kurikulum

| Level | Excel | Word | PowerPoint | Perkiraan waktu baca + praktik terarah* |
|---|---:|---:|---:|---:|
| Basic | 5 | 5 | 5 | 210 menit |
| Beginner | 5 | 5 | 5 | 270 menit |
| Intermediate | 5 | 5 | 5 | 330 menit |
| Advanced | 5 | 5 | 5 | 420 menit |
| Expert | 5 | 5 | 5 | 525 menit |
| **Total** | **25** | **25** | **25** | **1755 menit (29 jam 15 menit)** |

*Estimasi modul bersifat indikatif; diskusi, evaluasi, pengerjaan proyek, dan koreksi fasilitator memerlukan waktu tambahan. Durasi nyata tergantung kemampuan dan perangkat peserta.

## Jalur studi kasus sintetis

| Unit kerja | Tema kasus | Keluaran lintas aplikasi |
|---|---|---|
| Keuangan | Pengendalian pagu dan realisasi | Rekap Excel + dokumen Word + paparan PowerPoint |
| Kemahasiswaan | Pemantauan permohonan dan beasiswa | Rekap Excel + dokumen Word + paparan PowerPoint |
| Perencanaan | Monitoring target dan capaian program | Rekap Excel + dokumen Word + paparan PowerPoint |
| Persuratan & Arsip | Register surat dan metadata arsip | Rekap Excel + dokumen Word + paparan PowerPoint |
| Kepegawaian | Rekap pelatihan dan layanan pegawai | Rekap Excel + dokumen Word + paparan PowerPoint |
| Layanan Akademik | Administrasi surat dan registrasi akademik | Rekap Excel + dokumen Word + paparan PowerPoint |
| Pengadaan & Aset | Inventaris dan penggunaan sarana | Rekap Excel + dokumen Word + paparan PowerPoint |

## Batasan dan keamanan

- **Semua nama, angka, unit, dan berkas contoh adalah fiktif**, bukan informasi institusi dan tidak digunakan untuk menyatakan kondisi nyata.
- Aplikasi **bukan** Microsoft Excel/Word/PowerPoint asli. Simulator Excel berisi subset fungsi; hasil fitur Office lanjutan harus diverifikasi di Microsoft Office desktop.
- Ekspor dari praktik Excel berupa **CSV**, dari praktik Word dan PowerPoint berupa **HTML**. Webapp tidak menghasilkan berkas `.xlsx`, `.docx`, atau `.pptx` native. File CSV tidak menyimpan rumus/format workbook Excel secara utuh.
- Progress dan draf latihan disimpan pada `localStorage` browser. Bersihkan data pada komputer umum, ekspor progres jika perlu; jangan memasukkan data rahasia/asli.
- Bukti Belajar Mandiri **bukan sertifikat resmi** atau pengakuan kompetensi terverifikasi. Fasilitator harus memeriksa langsung dokumen praktik dan memutuskan penilaian akhir.
- Pedoman persuratan, tata naskah dinas, klasifikasi keamanan dokumen, JRA/retensi arsip, dan perlindungan data harus mengacu **ketentuan yang berlaku di masing-masing institusi**. Contoh OfficeLab tidak menggantikan SOP tersebut.

## Rujukan umum

- Microsoft Excel help & learning: https://support.microsoft.com/excel
- Microsoft Word help & learning: https://support.microsoft.com/word
- Microsoft PowerPoint help & learning: https://support.microsoft.com/powerpoint
- Portal ANRI: https://anri.go.id/

---
OfficeLab v2.1 · Dokumentasi pendamping untuk pelatihan administratif perguruan tinggi.
