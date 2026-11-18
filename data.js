'use strict';
// Semua materi disimpan langsung di perangkat agar aplikasi dapat dibuka tanpa koneksi.
const OFFICE = {
  excel: {
    name:'Excel', short:'X', tagline:'Kuasai register layanan, rekap anggaran, dan analisis data kampus.', description:'Dari penataan dataset unit kerja hingga validasi, rekap pagu–realisasi, rumus, dashboard, dan audit angka.',
    lessons:[
      {title:'Mengenal workbook & worksheet',level:'Dasar',min:6,lead:'Excel bekerja dengan buku kerja (workbook) yang berisi satu atau lebih lembar kerja (worksheet).',objective:'Membedakan workbook, worksheet, sel, baris, kolom, dan rentang.',points:['Workbook merupakan file Excel; worksheet merupakan lembar di dalam file tersebut.','Kolom diberi huruf (A, B, C), sedangkan baris diberi nomor (1, 2, 3).','Sel adalah pertemuan kolom dan baris, contohnya B4. Rentang A2:C7 mencakup beberapa sel sekaligus.','Name Box menunjukkan alamat sel dan Formula Bar menampilkan isi sel yang dipilih.'],example:'A1 = "Produk"\nB1 = "Harga"\nA2 = "Buku"\nB2 = 12000',exercise:'Buka tab Praktik Excel. Pilih sel B2 dan perhatikan bahwa alamat sel terlihat di sebelah bilah rumus.',tip:'Biasakan menempatkan nama kolom pada satu baris header dan satu record pada setiap baris.'},
      {title:'Memasukkan data & referensi sel',level:'Dasar',min:7,lead:'Data teks, angka, dan tanggal memiliki cara penggunaan yang berbeda dalam perhitungan.',objective:'Menulis alamat sel serta mengenali referensi relatif dan absolut.',points:['Referensi relatif A2 berubah ketika rumus disalin ke lokasi lain.','Referensi absolut $A$2 tetap menunjuk sel yang sama meskipun rumus disalin.','Referensi campuran $A2 mengunci kolom, sedangkan A$2 mengunci baris.','Gunakan format data konsisten; angka yang tersimpan sebagai teks dapat mengganggu rumus.'],example:'=C2*D2\n=C2*$J$1     // contoh: diskon tetap di J1',exercise:'Di tab Praktik, ubah D2 dari 5 menjadi 6. Perhatikan nilai E2 yang otomatis berubah.',tip:'Tekan F4 ketika mengedit referensi di Excel desktop untuk berpindah di antara bentuk relatif, absolut, dan campuran.'},
      {title:'Operasi hitung dan urutan perhitungan',level:'Dasar',min:7,lead:'Rumus dimulai dengan tanda sama dengan (=) dan bisa menggabungkan sel serta operator matematika.',objective:'Menggunakan +, -, *, /, ^ dan tanda kurung dengan benar.',points:['Operator + untuk penjumlahan, - untuk pengurangan, * untuk perkalian, / untuk pembagian, dan ^ untuk pangkat.','Excel mengutamakan tanda kurung, pangkat, perkalian/pembagian, kemudian penjumlahan/pengurangan.','Referensikan sel agar hasil otomatis berubah saat angka sumber diperbarui.','Gunakan rumus yang dapat diperiksa, bukan mengetik nilai hasil perhitungan secara manual.'],example:'=C2*D2\n=(C2*D2)*90%\n=(10+2)/3',exercise:'Ketik =C2*D2 di E2 pada lembar latihan, lalu tekan Enter.',tip:'Tanda titik atau koma desimal dan pemisah argumen dapat berbeda menurut pengaturan regional Excel.'},
      {title:'Rumus statistik: SUM sampai COUNT',level:'Dasar',min:8,lead:'Fungsi agregasi merangkum banyak baris menjadi informasi yang bermakna.',objective:'Membedakan SUM, AVERAGE, MIN, MAX, COUNT, dan COUNTA.',points:['SUM menambahkan angka dalam sebuah rentang.','AVERAGE menghitung rata-rata nilai numerik.','MIN dan MAX mencari angka terkecil dan terbesar.','COUNT menghitung sel berisi angka, sedangkan COUNTA menghitung sel tidak kosong.'],example:'=SUM(E2:E7)\n=AVERAGE(C2:C7)\n=MAX(C2:C7)\n=COUNT(D2:D7)',exercise:'Kerjakan tantangan SUM pada sel E8 dan AVERAGE pada sel C9.',tip:'Rentang biasanya lebih mudah dikelola dibanding menambahkan satu per satu sel, misalnya =E2+E3+E4.'},
      {title:'Logika dengan IF, AND, dan OR',level:'Menengah',min:9,lead:'Fungsi logika membantu membuat kategori atau keputusan berdasarkan suatu kondisi.',objective:'Menyusun pengujian kondisi menggunakan IF dan memahami AND/OR.',points:['IF memiliki tiga argumen: kondisi, nilai jika benar, nilai jika salah.','Operator perbandingan meliputi >, <, >=, <=, =, dan <>.','AND bernilai TRUE jika semua kondisi benar; OR bernilai TRUE jika salah satu kondisi benar.','Gunakan tanda kutip untuk hasil teks, misalnya "Lulus".'],example:'=IF(D2>=5;"Banyak";"Sedikit")\n=IF(AND(C2>=10000;D2>=5);"Prioritas";"Reguler")',exercise:'Di sel F2 buat IF: bila D2 >= 5 tampilkan "Banyak", selain itu "Sedikit".',tip:'Contoh menggunakan titik koma (;). Jika Excel Anda memakai koma (,), sesuaikan pemisahnya.'},
      {title:'Perhitungan bersyarat: COUNTIF & SUMIF',level:'Menengah',min:9,lead:'Saat data berisi beberapa kategori, perhitungan dapat dibatasi berdasarkan sebuah kriteria.',objective:'Menghitung jumlah entri dan total angka dengan satu syarat.',points:['COUNTIF(rentang;kriteria) menghitung sel yang memenuhi kriteria.','SUMIF(rentang_kriteria;kriteria;rentang_jumlah) menjumlahkan nilai yang terkait dengan kriteria.','Kriteria teks ditulis di dalam tanda kutip, contohnya "ATK".','Untuk beberapa kondisi, gunakan COUNTIFS atau SUMIFS.'],example:'=COUNTIF(B2:B7;"ATK")\n=SUMIF(B2:B7;"Elektronik";E2:E7)',exercise:'Kerjakan tantangan COUNTIF pada D9 dan SUMIF pada E9.',tip:'Pastikan rentang kondisi dan rentang jumlah memiliki panjang yang sama.'},
      {title:'Fungsi teks dan pembulatan',level:'Menengah',min:7,lead:'Data nyata sering perlu dibersihkan, digabung, atau dibulatkan sebelum dipakai.',objective:'Menggunakan ROUND, CONCAT, LEFT, RIGHT, LEN, dan TRIM.',points:['ROUND(angka;digit) membulatkan ke jumlah digit tertentu.','CONCAT menggabungkan beberapa potongan teks.','LEFT dan RIGHT mengambil karakter di awal atau akhir teks.','TRIM menghapus spasi berlebih, sementara LEN menghitung banyak karakter.'],example:'=ROUND(E2*0.9;0)\n=CONCAT(A3;" - ";B3)\n=LEFT(A2;2)',exercise:'Kerjakan tantangan ROUND pada G2 dan CONCAT pada G3.',tip:'Gunakan fungsi TRIM sebelum membandingkan data teks yang berasal dari sumber berbeda.'},
      {title:'Mengambil data dengan VLOOKUP & XLOOKUP',level:'Menengah',min:10,lead:'Fungsi pencarian mengambil nilai dari tabel tanpa memasukkan data secara manual berulang kali.',objective:'Memahami pencarian tepat (exact match) menggunakan VLOOKUP dan XLOOKUP.',points:['VLOOKUP mencari nilai pada kolom pertama dari rentang tabel dan mengembalikan nilai dari kolom bernomor tertentu.','Gunakan FALSE sebagai argumen terakhir VLOOKUP untuk pencarian tepat.','XLOOKUP memisahkan rentang pencarian dan rentang hasil sehingga lebih fleksibel.','XLOOKUP tersedia di versi Excel yang mendukungnya, termasuk Microsoft 365 dan versi modern yang sesuai.'],example:'=VLOOKUP("Mouse";A2:C7;3;FALSE)\n=XLOOKUP("Flashdisk";A2:A7;C2:C7)',exercise:'Di G4 cari harga Mouse memakai VLOOKUP dan di G5 cari harga Flashdisk memakai XLOOKUP.',tip:'Untuk dataset yang berkembang, Excel Table membantu menjaga rentang tetap mudah dipelihara.'},
      {title:'Sort, Filter, dan Excel Table',level:'Menengah',min:8,lead:'Mengurutkan dan memfilter mempermudah peninjauan baris yang banyak tanpa mengubah makna datanya.',objective:'Memilih teknik sort, filter, dan Excel Table.',points:['Sort mengatur urutan baris menurut nilai atau teks; pastikan seluruh tabel ikut dipilih.','Filter menampilkan subset data sesuai kriteria tanpa menghapus data lainnya.','Excel Table (Ctrl+T) memudahkan header, filter, dan referensi terstruktur.','Periksa baris kosong, header ganda, dan tipe data campuran sebelum analisis.'],example:'Ctrl+T          // Buat Excel Table\nCtrl+Shift+L    // Aktif/nonaktif filter',exercise:'Pada aplikasi Excel desktop, buat sebuah tabel dan filter Kategori = ATK. Di OfficeLab, amati data kategori pada worksheet.',tip:'Jangan mengurutkan satu kolom sendirian jika kolom tersebut merupakan bagian dari satu tabel data.'},
      {title:'Grafik dan PivotTable: pengantar',level:'Lanjutan',min:9,lead:'Grafik mengomunikasikan pola, sedangkan PivotTable meringkas data menurut dimensi kategori.',objective:'Memilih visualisasi yang sesuai dan memahami peran PivotTable.',points:['Grafik kolom cocok untuk perbandingan kategori; grafik garis cocok untuk tren waktu.','Gunakan judul dan label sumbu yang informatif serta skala yang tidak menyesatkan.','PivotTable mengelompokkan dan merangkum data, misalnya jumlah penjualan berdasarkan kategori.','Selalu periksa jumlah dan sumber data sebelum mengambil kesimpulan.'],example:'Kategori       Total\nATK            154000\nElektronik     555000',exercise:'Berdasarkan data latihan, identifikasi kategori dengan total nilai penjualan terbesar.',tip:'Jangan memakai grafik 3D jika mengganggu pembacaan nilai.'},
      {title:'Menangani error & mengaudit rumus',level:'Lanjutan',min:8,lead:'Kesalahan umum biasanya muncul karena data, referensi, atau penulisan argumen tidak sesuai.',objective:'Mengenali penyebab #DIV/0!, #VALUE!, #REF!, dan #N/A.',points:['#DIV/0! muncul ketika membagi angka dengan nol.','#VALUE! sering muncul saat data atau argumen tidak sesuai tipe yang diperlukan.','#REF! mengindikasikan referensi yang tidak valid.','#N/A umum muncul ketika nilai yang dicari tidak ditemukan.','IFERROR dapat menyediakan nilai alternatif ketika sebuah rumus menghasilkan error.'],example:'=IFERROR(A2/B2;0)\n=IFERROR(XLOOKUP("X";A2:A7;C2:C7);"Tidak ada")',exercise:'Buat percobaan =10/0 di sel G7 pada simulator. Perhatikan respons error, lalu perbaiki rumus.',tip:'Pahami penyebabnya terlebih dahulu. Menyembunyikan semua error tanpa pemeriksaan dapat menutupi masalah kualitas data.'}
    ],
    shortcuts:[
      ['Ctrl + S','Simpan workbook','Umum'],['Ctrl + N','Buat workbook baru','Umum'],['Ctrl + O','Buka workbook','Umum'],['Ctrl + P','Cetak / pratinjau','Umum'],['Ctrl + Z','Urungkan tindakan','Umum'],['Ctrl + Y','Ulangi tindakan','Umum'],['Ctrl + C','Salin pilihan','Umum'],['Ctrl + V','Tempel pilihan','Umum'],['Ctrl + X','Potong pilihan','Umum'],['Ctrl + F','Cari isi lembar','Umum'],['Ctrl + H','Cari dan ganti','Umum'],
      ['F2','Edit sel aktif','Sel'],['Alt + =','Sisipkan AutoSum','Rumus'],['F4','Ganti jenis referensi saat mengedit rumus','Rumus'],['Ctrl + `','Tampilkan/sembunyikan rumus','Rumus'],['Ctrl + 1','Buka Format Cells','Format'],['Ctrl + B','Tebalkan teks dalam sel','Format'],['Ctrl + Shift + L','Aktifkan/nonaktifkan filter','Data'],['Ctrl + T','Buat Excel Table','Data'],['Ctrl + E','Flash Fill','Data'],['Ctrl + D','Isi ke bawah (Fill Down)','Data'],['Ctrl + R','Isi ke kanan (Fill Right)','Data'],['Ctrl + ;','Masukkan tanggal saat ini','Data'],['Alt + Enter','Baris baru dalam satu sel','Sel'],['Ctrl + Space','Pilih seluruh kolom','Navigasi'],['Shift + Space','Pilih seluruh baris','Navigasi'],['Ctrl + Home','Ke awal worksheet','Navigasi'],['Ctrl + End','Ke sel terakhir area digunakan','Navigasi'],['Ctrl + Page Down','Ke worksheet berikutnya','Navigasi'],['Ctrl + Page Up','Ke worksheet sebelumnya','Navigasi'],['Shift + F11','Sisipkan worksheet','Data'],['Ctrl + tombol panah','Lompat ke tepi area data','Navigasi']
    ],
    formulas:[
      {name:'SUM',category:'Matematika',syntax:'=SUM(rentang)',desc:'Menjumlahkan semua angka pada sebuah rentang.',example:'=SUM(E2:E7)',result:'709.000',target:'E8'},
      {name:'AVERAGE',category:'Statistik',syntax:'=AVERAGE(rentang)',desc:'Menghitung nilai rata-rata dari data numerik.',example:'=AVERAGE(C2:C7)',result:'56.000',target:'C9'},
      {name:'MIN',category:'Statistik',syntax:'=MIN(rentang)',desc:'Mengambil angka paling kecil.',example:'=MIN(C2:C7)',result:'5.000',target:'G6'},
      {name:'MAX',category:'Statistik',syntax:'=MAX(rentang)',desc:'Mengambil angka paling besar.',example:'=MAX(C2:C7)',result:'150.000',target:'F9'},
      {name:'COUNT',category:'Statistik',syntax:'=COUNT(rentang)',desc:'Menghitung banyak sel berisi angka.',example:'=COUNT(D2:D7)',result:'6',target:'G6'},
      {name:'COUNTA',category:'Statistik',syntax:'=COUNTA(rentang)',desc:'Menghitung banyak sel tidak kosong.',example:'=COUNTA(A2:A7)',result:'6',target:'G6'},
      {name:'IF',category:'Logika',syntax:'=IF(kondisi;benar;salah)',desc:'Menentukan hasil berdasarkan kondisi logis.',example:'=IF(D2>=5;"Banyak";"Sedikit")',result:'Banyak',target:'F2'},
      {name:'AND',category:'Logika',syntax:'=AND(kondisi1;kondisi2)',desc:'TRUE hanya jika seluruh kondisi benar.',example:'=AND(C2>=10000;D2>=5)',result:'TRUE',target:'G6'},
      {name:'OR',category:'Logika',syntax:'=OR(kondisi1;kondisi2)',desc:'TRUE jika minimal satu kondisi benar.',example:'=OR(B2="ATK";D2>7)',result:'TRUE',target:'G6'},
      {name:'COUNTIF',category:'Kriteria',syntax:'=COUNTIF(rentang;kriteria)',desc:'Menghitung sel yang memenuhi satu syarat.',example:'=COUNTIF(B2:B7;"ATK")',result:'3',target:'D9'},
      {name:'SUMIF',category:'Kriteria',syntax:'=SUMIF(rentang;kriteria;rentang_jumlah)',desc:'Menjumlahkan angka untuk satu kriteria.',example:'=SUMIF(B2:B7;"Elektronik";E2:E7)',result:'555.000',target:'E9'},
      {name:'COUNTIFS',category:'Kriteria',syntax:'=COUNTIFS(rentang1;syarat1;rentang2;syarat2)',desc:'Menghitung baris yang memenuhi beberapa kondisi.',example:'=COUNTIFS(B2:B7;"ATK";D2:D7;">=5")',result:'3',target:'G6'},
      {name:'SUMIFS',category:'Kriteria',syntax:'=SUMIFS(rentang_jumlah;rentang1;syarat1)',desc:'Menjumlahkan angka dengan beberapa syarat.',example:'=SUMIFS(E2:E7;B2:B7;"ATK")',result:'154.000',target:'G6'},
      {name:'ROUND',category:'Matematika',syntax:'=ROUND(angka;digit)',desc:'Membulatkan angka ke digit tertentu.',example:'=ROUND(E2*0.9;0)',result:'54.000',target:'G2'},
      {name:'CONCAT',category:'Teks',syntax:'=CONCAT(teks1;teks2;...)',desc:'Menggabungkan beberapa bagian teks.',example:'=CONCAT(A3;" - ";B3)',result:'Pulpen - ATK',target:'G3'},
      {name:'LEFT',category:'Teks',syntax:'=LEFT(teks;jumlah)',desc:'Mengambil beberapa karakter paling kiri.',example:'=LEFT(A2;2)',result:'Bu',target:'G6'},
      {name:'RIGHT',category:'Teks',syntax:'=RIGHT(teks;jumlah)',desc:'Mengambil beberapa karakter paling kanan.',example:'=RIGHT(A2;2)',result:'ku',target:'G6'},
      {name:'LEN',category:'Teks',syntax:'=LEN(teks)',desc:'Menghitung jumlah karakter pada teks.',example:'=LEN(A2)',result:'4',target:'G6'},
      {name:'TRIM',category:'Teks',syntax:'=TRIM(teks)',desc:'Menghilangkan spasi berlebih dari teks.',example:'=TRIM("  Data   Excel  ")',result:'Data Excel',target:'G6'},
      {name:'VLOOKUP',category:'Pencarian',syntax:'=VLOOKUP(nilai;tabel;kolom;FALSE)',desc:'Cari pada kolom pertama dan ambil nilai di kolom tujuan.',example:'=VLOOKUP("Mouse";A2:C7;3;FALSE)',result:'75.000',target:'G4'},
      {name:'XLOOKUP',category:'Pencarian',syntax:'=XLOOKUP(nilai;rentang_cari;rentang_hasil)',desc:'Mencari data secara fleksibel dengan pencocokan tepat default.',example:'=XLOOKUP("Flashdisk";A2:A7;C2:C7)',result:'85.000',target:'G5'},
      {name:'IFERROR',category:'Logika',syntax:'=IFERROR(nilai;jika_error)',desc:'Menampilkan hasil pengganti ketika sebuah rumus gagal.',example:'=IFERROR(10/0;"Tidak valid")',result:'Tidak valid',target:'G6'}
    ],
    quiz:[
      ['Alamat B5 merujuk ke ...',['Kolom B, baris 5','Baris B, kolom 5','Lembar B halaman 5','Sel B sampai 5'],0,'Huruf mewakili kolom, angka mewakili baris.'],
      ['Simbol pertama untuk menulis rumus Excel adalah ...',['#','$','=','@'],2,'Excel mengenali rumus ketika diawali tanda sama dengan (=).'],
      ['Manakah rumus yang menjumlahkan E2 hingga E7?',['=ADD(E2:E7)','=SUM(E2:E7)','=COUNT(E2:E7)','=AVG(E2:E7)'],1,'SUM menambahkan nilai numerik dalam rentang.'],
      ['Apa kegunaan tanda dolar pada $A$1?',['Mengubah sel menjadi mata uang','Mengunci baris dan kolom referensi','Mengurutkan data','Menandai sheet rahasia'],1,'$A$1 adalah referensi absolut.'],
      ['Fungsi untuk menghitung rata-rata adalah ...',['MAX','COUNT','AVERAGE','LEN'],2,'AVERAGE adalah fungsi statistik untuk rata-rata.'],
      ['Arti rumus =IF(A1>=75;"Lulus";"Ulang") ketika A1 = 80 ialah ...',['Ulang','75','TRUE','Lulus'],3,'Karena 80 >= 75, argumen kedua dipilih.'],
      ['Fungsi untuk menghitung jumlah baris kategori ATK ialah ...',['COUNTIF','ROUND','CONCAT','MAX'],0,'COUNTIF menghitung sel sesuai kriteria.'],
      ['Shortcut Ctrl + Shift + L pada Excel desktop berguna untuk ...',['Membuat slide','Mengaktifkan/mematikan filter','Mencetak','Membuat rumus baru'],1,'Ctrl+Shift+L mengaktifkan atau menonaktifkan AutoFilter.'],
      ['Perintah yang tepat untuk membuat Excel Table ialah ...',['Ctrl + T','Ctrl + B','Alt + Enter','Shift + F5'],0,'Ctrl+T mengubah rentang menjadi tabel Excel.'],
      ['Kesalahan #DIV/0! menunjukkan ...',['Referensi tak valid','Kesalahan ejaan','Pembagian dengan nol','Nama sheet salah'],2,'#DIV/0! adalah error pembagian dengan nol.'],
      ['Perbedaan utama COUNT dan COUNTA adalah ...',['COUNT untuk angka, COUNTA untuk tidak kosong','COUNT untuk teks, COUNTA untuk grafik','Tidak ada beda','COUNT untuk chart, COUNTA untuk rumus'],0,'COUNT khusus numerik, COUNTA menghitung sel yang tidak kosong.'],
      ['Fungsi VLOOKUP dengan FALSE pada argumen terakhir berarti ...',['Mencari pendekatan','Menonaktifkan formula','Pencarian exact match','Mengurutkan descending'],2,'FALSE berarti pencocokan tepat.'],
      ['Apa fungsi F2 pada Excel desktop?',['Memulai slideshow','Edit sel aktif','Menebalkan isi sel','Mencetak halaman'],1,'F2 membuka mode edit isi sel.'],
      ['Fungsi SUMIF digunakan untuk ...',['Mencari karakter pertama','Menjumlahkan sesuai kriteria','Mengganti warna sel','Menyisipkan filter'],1,'SUMIF menggabungkan penjumlahan dengan satu kriteria.'],
      ['Ketika formula disalin, referensi A2 pada umumnya ...',['Tetap permanen','Berubah mengikuti posisi','Menjadi teks','Menghilang'],1,'A2 adalah referensi relatif.'],
      ['Grafik yang cocok untuk melihat perubahan nilai dari waktu ke waktu ialah ...',['Diagram lingkaran','Grafik garis','Grafik radar','Gambar dekoratif'],1,'Grafik garis lazim untuk menunjukkan tren waktu.']
    ]
  },
  word:{
    name:'Word',short:'W',tagline:'Tulis dokumen lebih rapi, konsisten, dan profesional.',description:'Berlatih pemformatan teks, struktur heading, paragraf, tabel, dan penyusunan dokumen.',
    lessons:[
      {title:'Mengenal dokumen & Ribbon Word',level:'Dasar',min:6,lead:'Microsoft Word merupakan pengolah kata untuk menyiapkan dokumen yang memiliki struktur dan tata letak.',objective:'Memahami dokumen, Ribbon, Quick Access Toolbar, dan panel navigasi.',points:['Dokumen terdiri dari teks, paragraf, halaman, dan objek seperti tabel atau gambar.','Ribbon mengelompokkan perintah dalam tab Home, Insert, Layout, References, Review, dan lainnya.','Simpan berkala dan gunakan nama berkas yang menjelaskan isi serta revisinya.','Gunakan Navigation Pane untuk menelusuri dokumen panjang berdasarkan heading.'],example:'Judul Dokumen\nPendahuluan\nMetode\nHasil dan Pembahasan',exercise:'Masuk ke Praktik Word, pilih dokumen contoh, lalu lakukan pemformatan judul.',tip:'Aktifkan tampilan tanda format (¶) di Word desktop ketika meninjau paragraf dan spasi.'},
      {title:'Format karakter: bold, italic, underline',level:'Dasar',min:7,lead:'Penekanan teks sebaiknya membantu pembaca mengenali gagasan penting.',objective:'Mengatur format teks dan menghindari penekanan yang berlebihan.',points:['Bold menegaskan istilah atau struktur informasi tertentu.','Italic dapat dipakai untuk istilah asing atau penekanan yang sesuai gaya penulisan.','Underline dipakai seperlunya untuk menghindari kesan seperti hyperlink.','Gunakan font yang mudah dibaca, ukuran konsisten, dan warna dengan kontras memadai.'],example:'Ctrl+B → tebal\nCtrl+I → miring\nCtrl+U → garis bawah',exercise:'Di editor praktik, sorot frasa tertentu dan tekan tombol B, I, atau U pada toolbar.',tip:'Jangan memformat setiap judul secara manual. Gunakan Styles untuk struktur yang konsisten.'},
      {title:'Paragraf, alignment, dan spasi',level:'Dasar',min:8,lead:'Tata paragraf memengaruhi kenyamanan baca dan kerapian dokumen.',objective:'Menerapkan rata kiri, kanan, tengah, justify, dan spasi paragraf.',points:['Rata kiri umumnya baik untuk teks panjang dengan spasi kata yang stabil.','Justify meratakan kedua tepi tetapi perlu diperiksa untuk jarak antarkata yang terlalu renggang.','Line spacing dan spacing before/after adalah properti berbeda.','Gunakan pengaturan indentasi, bukan banyak spasi atau tombol Tab berulang.'],example:'Ctrl+L → rata kiri\nCtrl+E → rata tengah\nCtrl+R → rata kanan\nCtrl+J → justify',exercise:'Di editor praktik, pilih sebuah paragraf lalu ubah perataannya.',tip:'Jangan menggunakan Enter berkali-kali untuk membuat jarak antarbab.'},
      {title:'Heading & Styles yang konsisten',level:'Menengah',min:8,lead:'Styles memberikan hierarki teks yang dapat dimanfaatkan oleh daftar isi otomatis.',objective:'Membedakan judul dokumen, Heading 1, Heading 2, dan Normal.',points:['Heading 1 digunakan untuk judul bagian utama; Heading 2 untuk subbagian.','Struktur heading membantu navigasi dokumen dan aksesibilitas.','Modifikasi definisi Style agar perubahan berlaku konsisten di seluruh dokumen.','Daftar isi otomatis menggunakan struktur heading, bukan sekadar teks berukuran besar.'],example:'Heading 1: BAB I PENDAHULUAN\nHeading 2: 1.1 Latar Belakang\nNormal: Teks isi paragraf...',exercise:'Gunakan dropdown Style pada Praktik Word untuk menerapkan Judul 1.',tip:'Teks yang hanya ditulis tebal belum tentu dikenali Word sebagai heading.'},
      {title:'Daftar berpoin, bernomor, dan tabel',level:'Menengah',min:8,lead:'Informasi terstruktur lebih mudah dibaca ketika ditata menjadi daftar atau tabel.',objective:'Menggunakan daftar otomatis serta tabel secara tepat.',points:['Gunakan bullet untuk daftar tanpa urutan prioritas.','Gunakan numbering untuk langkah yang memerlukan urutan.','Tabel cocok untuk data dalam baris dan kolom; selalu berikan header yang jelas.','Hindari menyusun kolom hanya memakai tab atau banyak spasi.'],example:'Tujuan pembelajaran:\n• Mengolah data\n• Menyusun dokumen\n• Menyiapkan presentasi',exercise:'Di Praktik Word, masukkan daftar berpoin berisi minimal dua butir dan tabel.',tip:'Untuk tabel kompleks, pertimbangkan mengolah data di Excel kemudian menyajikannya di Word.'},
      {title:'Page setup, header, footer, page number',level:'Menengah',min:9,lead:'Dokumen formal memerlukan ukuran kertas, margin, dan nomor halaman yang konsisten.',objective:'Mengenal margin, orientasi, section break, serta header/footer.',points:['Page Setup menentukan margin, ukuran kertas, dan orientasi potret atau lanskap.','Header dan footer adalah area berulang di atas dan bawah halaman.','Section break memungkinkan format halaman berbeda dalam satu dokumen.','Nomor halaman dapat diatur berbeda pada halaman pembuka dan isi.'],example:'Layout → Margins\nInsert → Page Number\nLayout → Breaks → Next Page',exercise:'Catat kombinasi margin dan orientasi untuk laporan akademik yang sedang dikerjakan.',tip:'Gunakan Page Break (Ctrl+Enter) ketika ingin memulai halaman baru; jangan menambah Enter kosong.'},
      {title:'Referensi, caption, dan daftar isi',level:'Lanjutan',min:9,lead:'Dokumen akademik sering membutuhkan rujukan, keterangan gambar, dan daftar isi.',objective:'Menggunakan daftar isi otomatis serta caption yang dapat diperbarui.',points:['Terapkan Heading 1/2/3 sebelum membuat daftar isi otomatis.','Gunakan References → Table of Contents untuk membuat dan memperbarui daftar isi.','Caption memudahkan penomoran tabel atau gambar yang konsisten.','Cross-reference membantu merujuk tabel/gambar secara otomatis jika strukturnya berubah.'],example:'References → Table of Contents\nReferences → Insert Caption\nReferences → Cross-reference',exercise:'Susun struktur judul dan subjudul lalu tinjau keteraturan hierarchy heading.',tip:'Setelah mengedit judul, lakukan Update Table untuk memperbarui daftar isi.'},
      {title:'Review: Track Changes & komentar',level:'Lanjutan',min:7,lead:'Siklus revisi membutuhkan jejak perubahan dan komentar yang bisa ditinjau bersama.',objective:'Memahami cara meninjau usulan revisi dokumen.',points:['Track Changes mencatat perubahan teks sebagai saran yang dapat diterima atau ditolak.','Komentar menjelaskan alasan revisi tanpa mengubah isi dokumen secara langsung.','Gunakan Compare untuk membandingkan dokumen ketika diperlukan.','Sebelum finalisasi, periksa apakah terdapat perubahan atau komentar yang belum ditangani.'],example:'Review → Track Changes\nReview → New Comment\nReview → Accept / Reject',exercise:'Tulis satu paragraf revisi di editor dan gunakan format yang jelas untuk membedakan istilah penting.',tip:'Hindari menganggap teks yang tampil rapi sebagai bukti bahwa semua revisi sudah diterima.'},
      {title:'Ekspor, cetak, dan pemeriksaan akhir',level:'Lanjutan',min:7,lead:'Sebelum dibagikan, dokumen perlu diperiksa agar tetap terbaca di perangkat lain.',objective:'Melakukan quality check sebelum ekspor PDF atau cetak.',points:['Periksa heading, nomor halaman, pemenggalan tabel, caption, dan daftar isi.','Gunakan Print Preview untuk memastikan tidak ada halaman kosong yang tidak disengaja.','Ekspor PDF berguna untuk mempertahankan tata letak pada penerima.','Simpan file sumber yang masih dapat diedit selain file PDF final.'],example:'File → Export / Save As → PDF\nCtrl+P → Print Preview',exercise:'Gunakan fitur Cetak pada Praktik Word untuk melihat penataan dokumen mini.',tip:'Baca dokumen sekali lagi pada tampilan halaman, bukan hanya saat sedang mengedit.'}
    ],
    shortcuts:[
      ['Ctrl + N','Dokumen baru','Dokumen'],['Ctrl + O','Buka dokumen','Dokumen'],['Ctrl + S','Simpan dokumen','Dokumen'],['Ctrl + P','Cetak / pratinjau','Dokumen'],['Ctrl + B','Tebal (bold)','Format'],['Ctrl + I','Miring (italic)','Format'],['Ctrl + U','Garis bawah','Format'],['Ctrl + L','Rata kiri','Paragraf'],['Ctrl + E','Rata tengah','Paragraf'],['Ctrl + R','Rata kanan','Paragraf'],['Ctrl + J','Rata kiri-kanan (justify)','Paragraf'],['Ctrl + 1','Spasi baris tunggal','Paragraf'],['Ctrl + 2','Spasi baris ganda','Paragraf'],['Ctrl + 5','Spasi baris 1,5','Paragraf'],['Ctrl + Enter','Page break / halaman baru','Halaman'],['Ctrl + Alt + 1','Terapkan Heading 1','Styles'],['Ctrl + Alt + 2','Terapkan Heading 2','Styles'],['Ctrl + Shift + N','Terapkan gaya Normal','Styles'],['Ctrl + F','Cari teks','Edit'],['Ctrl + H','Cari dan ganti','Edit'],['Ctrl + K','Sisipkan hyperlink','Edit'],['Ctrl + Z','Urungkan','Edit'],['Ctrl + Y','Ulangi','Edit'],['Ctrl + Home','Ke awal dokumen','Navigasi'],['Ctrl + End','Ke akhir dokumen','Navigasi'],['F7','Periksa ejaan dan tata bahasa','Review'],['Shift + F3','Ubah kapitalisasi teks pilihan','Format']
    ],
    quiz:[
      ['Apa fungsi utama Styles pada Word?',['Menghapus file otomatis','Memastikan format dan hierarki konsisten','Menyimpan video','Menghitung angka'],1,'Styles menjaga konsistensi format dan mendukung navigasi/daftar isi.'],
      ['Shortcut untuk menebalkan teks pada Word desktop ialah ...',['Ctrl + I','Ctrl + B','Ctrl + U','Ctrl + K'],1,'Ctrl+B mengaktifkan atau menonaktifkan bold.'],
      ['Yang tepat untuk memulai halaman baru adalah ...',['Menekan Enter 20 kali','Ctrl + Enter','Menambahkan spasi','Ctrl + E'],1,'Ctrl+Enter membuat page break.'],
      ['Heading 2 sebaiknya dipakai untuk ...',['Subbagian dari Heading 1','Membuat gambar','Menyisipkan footer','Mengatur zoom'],0,'Heading 2 menunjukkan tingkat subbagian.'],
      ['Daftar isi otomatis paling bergantung pada ...',['Warna font','Struktur Styles heading','Jumlah gambar','Nama file'],1,'Daftar isi dibuat dari struktur heading.'],
      ['Perintah untuk meratakan kedua sisi paragraf ialah ...',['Ctrl + L','Ctrl + J','Ctrl + E','Ctrl + R'],1,'Ctrl+J mengaktifkan justify.'],
      ['Fungsi Track Changes adalah ...',['Merekam revisi','Menghapus komentar','Mengunci keyboard','Mengubah warna layar'],0,'Track Changes mencatat perubahan yang dapat ditinjau.'],
      ['Dalam dokumen formal, tabel sebaiknya memiliki ...',['Data acak','Header yang jelas','Font hias untuk semua teks','Tanpa border selalu'],1,'Header membantu interpretasi kolom.'],
      ['Perintah untuk membuka pencarian teks ialah ...',['Ctrl + F','Ctrl + P','Ctrl + O','Ctrl + U'],0,'Ctrl+F mencari teks.'],
      ['Apa perbedaan line spacing dan spacing after?',['Tidak berbeda','Satu untuk jarak baris, satu untuk jarak setelah paragraf','Keduanya mengubah font','Keduanya membuat tabel'],1,'Spacing after mengatur jarak setelah paragraf, bukan dalam paragraf.'],
      ['Shortcut Ctrl + Alt + 1 menerapkan ...',['Heading 1','Footnote','Page number','Superscript'],0,'Ctrl+Alt+1 adalah shortcut Heading 1.'],
      ['Mengapa dokumen sering diekspor ke PDF?',['Menghapus seluruh metadata','Menjaga tampilan lebih konsisten saat dibagikan','Agar rumus dapat dihitung','Agar video otomatis muncul'],1,'PDF umum digunakan untuk mempertahankan tata letak.'],
      ['Cara terbaik memberi jarak antarbab ialah ...',['Enter berulang','Spacing paragraph / page break yang tepat','Spasi 50 kali','Menambah tab'],1,'Gunakan pengaturan paragraf atau break sesuai struktur.'],
      ['Shortcut pemeriksaan ejaan pada Word ialah ...',['F7','F2','F5','F12'],0,'F7 membuka pemeriksaan ejaan/tata bahasa di Word desktop.'],
      ['Header berbeda dari footer karena ...',['Header di atas halaman, footer di bawah','Header hanya untuk tabel','Footer selalu huruf miring','Tidak ada perbedaan'],0,'Header berada di bagian atas halaman, footer di bawah.']
    ]
  },
  powerpoint:{
    name:'PowerPoint',short:'P',tagline:'Bangun cerita visual yang jelas dan presentasikan dengan percaya diri.',description:'Pelajari layout, desain slide, visual, transisi, dan praktik membuat tayangan sendiri.',
    lessons:[
      {title:'Merencanakan alur presentasi',level:'Dasar',min:7,lead:'Slide yang baik mengikuti pesan, bukan sekadar memindahkan isi dokumen ke layar.',objective:'Menyusun tujuan, audiens, dan narasi presentasi.',points:['Tentukan satu pesan utama untuk keseluruhan presentasi.','Gunakan alur pembuka → masalah → temuan/solusi → kesimpulan.','Satu slide sebaiknya membawa satu gagasan utama.','Buat outline sebelum memilih warna atau animasi.'],example:'Slide 1: Judul & tujuan\nSlide 2: Latar masalah\nSlide 3: Hasil kunci\nSlide 4: Rekomendasi',exercise:'Di Praktik PowerPoint, buat setidaknya tiga slide dan beri masing-masing judul berbeda.',tip:'Judul slide yang menyatakan kesimpulan lebih informatif dibanding judul yang terlalu umum.'},
      {title:'Slide, layout & placeholder',level:'Dasar',min:7,lead:'Layout menentukan posisi judul, teks, dan objek secara konsisten.',objective:'Mengatur layout serta menambah, menggandakan, dan memindahkan slide.',points:['Placeholder adalah area khusus untuk judul, teks, atau objek.','Layout sebaiknya dipilih berdasarkan jenis isi, bukan asal dekorasi.','Slide dapat ditambah, diurutkan ulang, dan digandakan.','Gunakan panel thumbnail untuk meninjau alur narasi.'],example:'Ctrl+M → slide baru\nCtrl+D → duplikasi objek\nCtrl+Shift+D → duplikasi slide',exercise:'Coba tombol Tambah Slide, Duplikasi, dan Pindahkan pada simulator.',tip:'Duplikasi slide untuk mempertahankan elemen yang konsisten tanpa memulai dari awal.'},
      {title:'Tema, warna, dan Slide Master',level:'Menengah',min:9,lead:'Identitas visual yang konsisten memudahkan audiens mengikuti isi presentasi.',objective:'Memilih palet, font, dan mengenal Slide Master.',points:['Gunakan palet warna dengan peran jelas: utama, aksen, teks, dan latar.','Pastikan kontras teks dengan latar cukup tinggi.','Slide Master memusatkan pengaturan desain yang diterapkan ke banyak slide.','Jangan mengganti terlalu banyak jenis font dan ukuran tanpa alasan.'],example:'View → Slide Master\nDesign → Themes / Variants',exercise:'Coba mengubah tema warna di simulator dan perhatikan semua slide ikut menyesuaikan.',tip:'Terapkan variasi yang cukup untuk orientasi visual, bukan variasi acak pada setiap slide.'},
      {title:'Tipografi & hierarki visual',level:'Dasar',min:8,lead:'Audiens biasanya melihat slide dari jarak yang lebih jauh daripada membaca dokumen.',objective:'Menjaga ukuran teks, hierarki judul, dan ruang kosong.',points:['Gunakan judul besar, ringkas, dan spesifik.','Kurangi blok teks panjang; utamakan poin yang membantu penyampaian lisan.','Gunakan alignment dan ruang kosong agar kelompok informasi mudah dipahami.','Hindari mencampur terlalu banyak gaya tipografi dalam satu slide.'],example:'Judul: "Penjualan meningkat 18%"\nIsi: 2–4 bukti singkat yang relevan',exercise:'Ringkas salah satu slide menjadi judul dan maksimal empat poin inti.',tip:'Jangan memaksakan angka ukuran font tertentu untuk semua situasi; pertimbangkan ukuran layar dan jarak audiens.'},
      {title:'Gambar, ikon, dan grafik',level:'Menengah',min:9,lead:'Elemen visual sebaiknya menyampaikan informasi, bukan hanya mengisi ruang.',objective:'Memilih visual yang relevan dan membuat grafik mudah ditafsirkan.',points:['Pilih gambar resolusi memadai dan perhatikan izin penggunaan.','Gunakan grafik batang untuk membandingkan kategori dan garis untuk tren waktu.','Berikan label yang jelas; hindari dekorasi 3D yang mengganggu interpretasi.','Gunakan alternatif teks (alt text) pada visual penting untuk aksesibilitas.'],example:'Hasil Survei\nSangat puas 60%\nPuas 30%\nLainnya 10%',exercise:'Buat sebuah slide ringkasan hasil berisi angka dan satu pesan utama.',tip:'Visualisasi sebaiknya bisa dipahami meskipun penjelasan lisan terlewat.'},
      {title:'Transisi & animasi secukupnya',level:'Menengah',min:7,lead:'Gerakan yang terarah dapat membantu fokus, tetapi efek berlebihan membuat pesan terpecah.',objective:'Membedakan transisi antar-slide dan animasi pada objek.',points:['Transition memengaruhi perpindahan dari satu slide ke slide lain.','Animation memengaruhi kemunculan/pergerakan objek di dalam slide.','Pilih durasi seperlunya dan hindari terlalu banyak efek.','Gunakan urutan kemunculan jika itu membantu menjelaskan langkah atau proses.'],example:'Transitions → Fade\nAnimations → Appear / Fade',exercise:'Pada aplikasi PowerPoint desktop, bandingkan sebuah slide tanpa transisi dan dengan transisi sederhana.',tip:'Efek tidak boleh menggantikan kualitas narasi dan desain.'},
      {title:'Speaker notes & Presenter View',level:'Menengah',min:8,lead:'Catatan pembicara dapat menyimpan detail yang tidak perlu tampil pada layar audiens.',objective:'Membedakan isi slide dengan speaker notes dan mengenal Presenter View.',points:['Isi slide menyampaikan poin inti; speaker notes menyimpan detail yang diucapkan.','Presenter View dapat menampilkan catatan, slide berikutnya, dan timer kepada pembicara.','Coba presentasi beberapa kali untuk memeriksa alur dan estimasi durasi.','Ketahui shortcut F5 dan Shift+F5 pada PowerPoint desktop.'],example:'F5 → mulai dari awal\nShift+F5 → mulai dari slide aktif\nEsc → keluar slideshow',exercise:'Isi kolom Catatan Pembicara di simulator, lalu jalankan Mode Presentasi.',tip:'Uji pengaturan tampilan proyektor sebelum acara dimulai.'},
      {title:'Presentasi efektif & navigasi',level:'Lanjutan',min:8,lead:'Saat menyajikan, kontrol waktu dan kemampuan navigasi membantu menjaga perhatian audiens.',objective:'Mengoperasikan slideshow dan memperjelas pesan saat berbicara.',points:['Gunakan tombol panah kanan/kiri untuk berpindah slide dalam slideshow.','Tombol B dapat menampilkan layar hitam sementara dalam slideshow PowerPoint desktop.','Jangan membacakan seluruh isi slide; jelaskan hubungan data dengan pesan.','Periksa keterbacaan slide dari sudut terjauh ruangan.'],example:'F5 → mulai\n→ → slide berikutnya\n← → slide sebelumnya\nEsc → selesai',exercise:'Coba mode presentasi penuh di simulator dan navigasi menggunakan panah kiri/kanan.',tip:'Lakukan simulasi presentasi dengan pengatur waktu sebelum presentasi resmi.'},
      {title:'Meninjau, berbagi, dan mengekspor',level:'Lanjutan',min:7,lead:'Versi presentasi final perlu diperiksa baik dari sisi isi maupun teknis penayangannya.',objective:'Mengecek bahasa, sumber, alignment, dan format ekspor.',points:['Periksa kesesuaian font, tema, dan konsistensi judul antar-slide.','Cantumkan sumber untuk data, gambar, dan kutipan yang digunakan.','Ekspor PDF bila audiens memerlukan salinan noninteraktif.','Simpan file sumber yang tetap dapat diedit.'],example:'File → Export → PDF\nFile → Save As → PPTX',exercise:'Tinjau seluruh slide pada simulator, lalu presentasikan dari awal sampai akhir.',tip:'File hasil latihan OfficeLab dapat diekspor sebagai HTML, bukan file .pptx native.'}
    ],
    shortcuts:[
      ['Ctrl + M','Tambah slide baru','Slide'],['Ctrl + D','Duplikasi objek terpilih','Objek'],['Ctrl + Shift + D','Duplikasi slide terpilih','Slide'],['Ctrl + N','Buat presentasi baru','Dokumen'],['Ctrl + S','Simpan presentasi','Dokumen'],['Ctrl + P','Cetak presentasi','Dokumen'],['Ctrl + Z','Urungkan perubahan','Edit'],['Ctrl + Y','Ulangi perubahan','Edit'],['Ctrl + C','Salin objek/teks','Edit'],['Ctrl + V','Tempel objek/teks','Edit'],['Ctrl + A','Pilih semua','Edit'],['Ctrl + K','Sisipkan hyperlink','Edit'],['Ctrl + G','Kelompokkan bentuk terpilih','Objek'],['Ctrl + Shift + G','Pisahkan kelompok bentuk','Objek'],['F5','Mulai slideshow dari awal','Presentasi'],['Shift + F5','Mulai dari slide saat ini','Presentasi'],['Esc','Keluar dari slideshow','Presentasi'],['N / →','Animasi/slide berikutnya','Presentasi'],['P / ←','Animasi/slide sebelumnya','Presentasi'],['B','Layar hitam saat slideshow','Presentasi'],['W','Layar putih saat slideshow','Presentasi'],['Alt + F5','Mulai Presenter View','Presentasi'],['Ctrl + B','Tebalkan teks pilihan','Format'],['Ctrl + I','Miringkan teks pilihan','Format'],['Ctrl + U','Garis bawahi teks pilihan','Format']
    ],
    quiz:[
      ['Shortcut untuk menambah slide baru ialah ...',['Ctrl + M','Ctrl + T','Ctrl + E','Shift + F5'],0,'Ctrl+M menambahkan slide di PowerPoint desktop.'],
      ['Untuk memulai presentasi dari awal digunakan ...',['F2','F5','F7','F12'],1,'F5 memulai slideshow dari slide pertama.'],
      ['Untuk memulai presentasi dari slide aktif digunakan ...',['Ctrl+N','Shift+F5','Alt+Enter','Ctrl+F'],1,'Shift+F5 memulai dari slide terpilih.'],
      ['Transisi berbeda dengan animasi karena ...',['Transisi antar-slide, animasi pada objek','Transisi untuk dokumen Word','Animasi menghapus slide','Tidak ada perbedaan'],0,'Transition untuk perpindahan slide, animation untuk objek.'],
      ['Speaker notes sebaiknya digunakan untuk ...',['Menutupi judul','Catatan detail pembicara','Mengganti semua isi slide','Menyimpan kata sandi'],1,'Catatan pembicara dapat berisi detail yang tak perlu diproyeksikan.'],
      ['Satu slide idealnya memiliki ...',['Sepuluh ide sekaligus','Satu pesan utama yang jelas','Semua teks laporan','Banyak warna tanpa fungsi'],1,'Satu pesan per slide memudahkan fokus.'],
      ['Grafik yang tepat untuk menunjukkan tren waktu adalah ...',['Grafik garis','Diagram 3D dekoratif','Ikon acak','Foto sampul'],0,'Grafik garis cocok untuk data deret waktu.'],
      ['Fungsi Slide Master adalah ...',['Menjaga elemen desain konsisten','Mengubah file menjadi video','Menyembunyikan seluruh angka','Menghitung rumus'],0,'Slide Master memberi kontrol terhadap desain di banyak slide.'],
      ['Apa fungsi tombol Esc ketika slideshow berlangsung?',['Beralih ke Word','Mengakhiri slideshow','Menambahkan slide','Menggandakan objek'],1,'Esc keluar dari slideshow.'],
      ['Bagaimana memastikan visual mudah dibaca?',['Kontras cukup dan teks ringkas','Font sekecil mungkin','Warna dan animasi sebanyak mungkin','Semua konten dalam satu slide'],0,'Kontras dan keterbacaan sangat penting.'],
      ['Untuk menduplikasi slide yang dipilih dapat menggunakan ...',['Ctrl + Shift + D','Ctrl + K','F7','Ctrl + Shift + L'],0,'Ctrl+Shift+D menduplikasi slide yang dipilih.'],
      ['Sebelum menyajikan slide sebaiknya ...',['Menguji proyektor, alur, dan durasi','Menghapus semua sumber','Mengubah tema di tengah presentasi','Menonaktifkan speaker notes permanen'],0,'Simulasi mengurangi kendala teknis dan narasi.'],
      ['Tombol B ketika slideshow PowerPoint desktop berfungsi untuk ...',['Membuat grafik','Menampilkan layar hitam sementara','Menebalkan semua slide','Mengunci presenter'],1,'B menyembunyikan slide dengan layar hitam sementara.'],
      ['Mengapa sumber data perlu ditampilkan?',['Agar slide ramai','Untuk transparansi dan keterlacakan','Agar animasi berjalan','Untuk menghemat jumlah slide'],1,'Sumber membantu audiens memverifikasi klaim dan data.'],
      ['Layout slide adalah ...',['Susunan placeholder dan konten','Jenis file audio','Nomor seri perangkat','Menu rumus Excel'],0,'Layout menentukan susunan ruang judul, teks, dan elemen lain.']
    ]
  }
};
const EXCEL_INITIAL = JSON.parse(JSON.stringify(EXCEL_CASES.keuangan.cells));
const EXCEL_TASKS = [
  {
    "id": "gap",
    "cell": "E8",
    "fn": "SUM",
    "title": "Jumlahkan selisih seluruh kegiatan",
    "desc": "Jumlahkan semua selisih E2:E7 menggunakan SUM.",
    "hint": "=SUM(E2:E7)"
  },
  {
    "id": "sumTarget",
    "cell": "C8",
    "fn": "SUM",
    "title": "Total target / pagu",
    "desc": "Jumlahkan C2:C7 sebagai total target atau pagu.",
    "hint": "=SUM(C2:C7)"
  },
  {
    "id": "sumActual",
    "cell": "D8",
    "fn": "SUM",
    "title": "Total realisasi / capaian",
    "desc": "Jumlahkan D2:D7 sebagai total realisasi.",
    "hint": "=SUM(D2:D7)"
  },
  {
    "id": "if",
    "cell": "F2",
    "fn": "IF",
    "title": "Status pemantauan kegiatan",
    "desc": "Jika D2 lebih besar dari C2, tampilkan \"Melebihi\"; selain itu \"Dalam batas\".",
    "hint": "=IF(D2>C2;\"Melebihi\";\"Dalam batas\")"
  },
  {
    "id": "average",
    "cell": "C9",
    "fn": "AVERAGE",
    "title": "Rata-rata realisasi",
    "desc": "Hitung rata-rata nilai D2:D7.",
    "hint": "=AVERAGE(D2:D7)"
  },
  {
    "id": "countif",
    "cell": "D9",
    "fn": "COUNTIF",
    "title": "Hitung data kelompok pilihan",
    "desc": "Hitung jumlah baris dengan B2:B7 = \"{criterion}\".",
    "hint": "=COUNTIF(B2:B7;\"{criterion}\")"
  },
  {
    "id": "sumif",
    "cell": "E9",
    "fn": "SUMIF",
    "title": "Total realisasi kelompok",
    "desc": "Jumlahkan D2:D7 untuk B2:B7 = \"{criterion}\".",
    "hint": "=SUMIF(B2:B7;\"{criterion}\";D2:D7)"
  },
  {
    "id": "round",
    "cell": "G2",
    "fn": "ROUND",
    "title": "Capaian persen satu desimal",
    "desc": "Hitung D2/C2 × 100, dibulatkan satu desimal.",
    "hint": "=ROUND(D2/C2*100;1)"
  },
  {
    "id": "concat",
    "cell": "G3",
    "fn": "CONCAT",
    "title": "Label kegiatan + kategori",
    "desc": "Gabungkan nama A3 dan kategori B3 dengan pemisah \" - \".",
    "hint": "=CONCAT(A3;\" - \";B3)"
  },
  {
    "id": "xlookup",
    "cell": "G4",
    "fn": "XLOOKUP",
    "title": "Cari realisasi dengan XLOOKUP",
    "desc": "Cari kegiatan \"{lookup}\" dari kolom A dan ambil nilainya dari kolom D.",
    "hint": "=XLOOKUP(\"{lookup}\";A2:A7;D2:D7)"
  },
  {
    "id": "vlookup",
    "cell": "G5",
    "fn": "VLOOKUP",
    "title": "Cari realisasi dengan VLOOKUP",
    "desc": "Cari kegiatan \"{lookup}\" pada A2:D7, ambil kolom ke-4 dengan FALSE.",
    "hint": "=VLOOKUP(\"{lookup}\";A2:D7;4;FALSE)"
  },
  {
    "id": "max",
    "cell": "G6",
    "fn": "MAX",
    "title": "Realisasi tertinggi",
    "desc": "Temukan nilai realisasi terbesar D2:D7.",
    "hint": "=MAX(D2:D7)"
  },
  {
    "id": "countifs",
    "cell": "G7",
    "fn": "COUNTIFS",
    "title": "Hitung dua kriteria",
    "desc": "Hitung kategori \"{criterion}\" dengan realisasi D2:D7 >=100.",
    "hint": "=COUNTIFS(B2:B7;\"{criterion}\";D2:D7;\">=100\")"
  },
  {
    "id": "sumifs",
    "cell": "G8",
    "fn": "SUMIFS",
    "title": "Jumlah dua kriteria",
    "desc": "Jumlahkan realisasi kategori \"{criterion}\" bila target C2:C7 >=100.",
    "hint": "=SUMIFS(D2:D7;B2:B7;\"{criterion}\";C2:C7;\">=100\")"
  }
];
const WORD_INITIAL = WORD_TEMPLATES.keuangan.html;
const WORD_TASKS = [
  {id:'heading',title:'Buat Heading 1',desc:'Ubah salah satu paragraf menjadi Heading 1.'},
  {id:'bold',title:'Gunakan teks tebal',desc:'Sorot sebuah frasa dan gunakan tombol B.'},
  {id:'italic',title:'Gunakan teks miring',desc:'Sorot sebuah istilah dan gunakan tombol I.'},
  {id:'bullets',title:'Buat daftar berpoin',desc:'Sisipkan daftar dengan minimal dua butir.'},
  {id:'table',title:'Sisipkan tabel',desc:'Buat tabel sederhana 2 × 2.'}
];
const PPT_INITIAL = JSON.parse(JSON.stringify(PPT_TEMPLATES.keuangan.slides));
const PPT_TASKS = [
  {id:'slides',title:'Minimal tiga slide',desc:'Susun sekurang-kurangnya tiga slide.'},
  {id:'titles',title:'Judul yang berbeda',desc:'Minimal tiga slide memiliki judul unik.'},
  {id:'bullets',title:'Slide dengan tiga poin',desc:'Tuliskan minimal tiga baris isi pada sebuah slide.'},
  {id:'theme',title:'Ubah tema warna',desc:'Pilih salah satu tema selain biru.'},
  {id:'notes',title:'Tulis catatan pembicara',desc:'Isi speaker notes pada salah satu slide.'},
  {id:'present',title:'Latih presentasi',desc:'Jalankan tayangan menggunakan tombol Presentasikan.'}
];
const COURSE_SECTIONS = [
  {id:'home',label:'Beranda',icon:'⌂'},
  {id:'lessons',label:'Materi',icon:'▤'},
  {id:'cases',label:'Studi Kasus Unit',icon:'▦'},
  {id:'shortcuts',label:'Shortcut',icon:'⌨'},
  {id:'formulas',label:'Rumus Excel',icon:'∑',only:'excel'},
  {id:'practice',label:'Praktik',icon:'⚙'},
  {id:'quiz',label:'Kuis',icon:'◈'},
  {id:'settings',label:'Pengaturan',icon:'⚑'},
  {id:'about',label:'Tentang OfficeLab',icon:'ⓘ'}
];
