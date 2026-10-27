Modeling Aplikasi AI Sekretariat:



Tujuan: Untuk menggantikan beberapa tugas sekretaris yang bisa di otomasi dengan suatu program/AI



Tugas yang menurut saya bisa digantikan:

* Merangkum rapat dan mencatatnya dengan input suara (Notulen) // **POIN 1**
* Mencatat sesuatu dengan input suara // **POIN 2**
* Membuat suatu surat tertentu // **POIN 3**
* Mencatat acara secara kronologis/berurutan dengan input tamplate manual // **POIN 4**
* Mengedit/menghapus data pada database yang ada // **POIN 5**



Cara kerja **POIN 1**:

* Pengguna menginput suatu file suara
* File suara ditranskrip ke teks
* Teks hasil transkrip dirangkum dan disusun secara rapih oleh AI
* Hasilnya ditampilkan kepada user, lalu user diminta untuk memberikan judul rapatnya, dan diberi pilihan untuk save/beri tambahan
* Jika user memutuskan untuk "save", teks hasilnya dimasukan ke database SQLite yang berisi table(ID; JUDUL\_RAPAT; HASIL\_RAPAT; TIME\_STAMP)
* User bisa mengakses hasilnya di suatu bagian khusus data rangkuman rapat



Cara kerja **POIN 2**:

* Pengguna menginput suatu file suara/menginput prompt manual (jika file suara digunakan, transkrip file menjadi teks)
* Model AI mengerjakan perintah user
* Hasilnya ditampilkan kepada user, lalu user diminta untuk memberi judul pada catatannya, lalu pengguna diberi pilihan untuk save/beri tambahan
* Jika user memutuskan untuk "save", teks hasilnya dimasukan ke database SQLite yang berisi table(ID; JUDUL; CATATAN)
* User bisa mengakses hasilnya di suatu bagian khusus data catatan pribadi



Cara kerja **POIN 3**:

* Pengguna menginput prompt yang menjelaskan isi dari suratnya
* Model AI mengerjakan surat sesuai dengan perintah user
* Hasil surat berupa teks HTML di ubah ke file dokumen
* Nama file dokumen ditampilkan kepada user lalu user diminta untuk mendownload hasilnya untuk melihat
* User diberi pilihan untuk save/tidak \*jika tidak, program selesai
* Jika user memutuskan untuk "SAVE", maka program mencari lokasi dengan nama file seperti sebelumnya, lalu mencatat alamat path-nya
* Alamat file dokumen dan nama file dokumen tersebut disimpan ke database SQLite berisi tabel(ID; NAMA\_FILE; ALAMAT\_FILE; TIME\_STAMP)
* User bisa mengakses hasilnya di suatu bagian khusus data surat



Cara kerja **POIN 4**:

* Pengguna diberikan menu input NAMA\_ACARA, WAKTU\_MULAI, WAKTU\_SELESAI, TANGGAL\_ACARA.
* Pengguna mengisi input tersebut
* Jika pengguna memencet tombol save, data tersebut disimpan di database yang berisi table(NAMA\_ACARA; WAKTU\_MULAI; WAKTU\_SELESAI; TANGGAL\_ACARA)
* Pada bagian tertentu pada program, data tersebut diurutkan sesuai dengan hari, tanggal, dan tahun



Cara kerja **POIN 5**:

* Pengguna memilih database dan tabel mana yang ingin dihapus
* Pengguna ditampilkan semua data yang ada pada table tersebut
* Pengguna mamasukan nama ID/NAMA data yang ingin dihapus
* Pengguna diberikan pertanyaan ulang untuk meyakinkan pengguna untuk menghapus data tersebut
* Data dengan ID/NAMA tersebut akan didrop/hapus dari table



Cara kerja **Bagian Review**:

* Pengguna memilih untuk melihat ke data apa (EX: DATA\_NOTULEN/DATA\_CATATAN\_PRIBADI/DATA\_SURAT/DATA\_ACARA)
* Program menampilkan data yang telah dipilih oleh user.

  * **DATA NOTULEN**, ditampilkan list judul rapat yang tersedia, pengguna memilih judul mana yang mau dia lihat, membuka hasil rapat dengan judul yang dipilih
  * **DATA CATATAN PRIBADI**, ditampilkan lis judul catatan yang tersedia, pengguna memilih judul mana yang mau dia lihat, membuka hasil rapat dengan judul yang dipilih
  * **DATA SURAT**, ditampilkan list nama file dokumen, pengguna memilih file dokumen mana yang mau dia lihat, membuka hasil file dokumen berdasarkan alamat file yang tercatat
  * **DATA ACARA**, Pengguna diberi pilihan SEMUA\_ACARA/ACARA\_TAHUN\_INI/ACARA\_BULAN\_INI, pengguna memilih rentang waktu mana yang mau dia lihat, membuka data acara pada rentang Waktu tersebut



STRUKTUR DATA:



📁APP AI SEKRETARIAT

|\\\_Modeling Aplikasi AI Sekretariat.md

&#x20;\\\_mainproces.py

&#x20;\\\_📁Assets

&#x20;  |\\\_hasil\_rapat.db

&#x20;  |\\\_catatan\_pribadi.db

&#x20;  |\\\_hasil\_surat.db

&#x20;   \\\_data\_acara.db

