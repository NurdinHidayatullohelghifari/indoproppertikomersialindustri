# IndoProperti Komersial & Industri

Portal informasi properti komersial dan industri Indonesia untuk lahan, gudang, pabrik, dan aset bisnis.

## Fitur

- Katalog listing dengan pencarian, filter kategori, dan pengurutan
- Form pengajuan publik untuk titip jual atau mencari properti
- Dashboard admin untuk pengajuan dan CRUD listing (tambah, edit, hapus)
- Upload foto dan dokumen pada formulir dan listing
- Info properti dari pengajuan dapat ditampilkan otomatis di katalog publik tanpa nama dan nomor WhatsApp
- Backup/restore JSON dan ekspor CSV dari dashboard
- Penyimpanan demo lokal menggunakan `localStorage`
- Integrasi opsional Google Apps Script untuk menyimpan data dan lampiran di Drive/Sheets

## Menjalankan secara lokal

Buka `index.html` di browser. Listing dan pengajuan demo disimpan di `localStorage` pada browser tersebut, bukan dibagikan ke perangkat lain.

## Integrasi Google Apps Script

`google-apps-script.gs` adalah backend opsional untuk menyimpan data dan lampiran ke Google Drive dan Sheets.

1. Buat project di Google Apps Script dan salin isi `google-apps-script.gs`.
2. Jalankan `setupIndoProperti()` dari editor Apps Script dan berikan izin yang diminta.
3. Deploy sebagai Web app dengan eksekusi sebagai pemilik project dan akses `Siapa saja`. URL deployment aktif sudah dikonfigurasi sebagai URL default di `index.html`; pengaturan pada Dashboard Admin hanya menjadi override lokal pada browser itu.
4. Uji endpoint, penyimpanan data, dan lampiran dengan data non-produksi sebelum digunakan.

Setelah mengubah `google-apps-script.gs`, perbarui deployment Web app. Pengajuan baru dari formulir publik disimpan di Sheet; informasi properti dari pengajuan yang belum diposting admin dapat muncul di katalog publik. Hanya judul/kategori, status jual atau cari, lokasi, harga, luas, dan deskripsi yang dikirim sebagai listing publik; nama, nomor WhatsApp, dan lampiran tidak disertakan. Pengajuan yang sudah diposting admin tidak ditampilkan dua kali.

**Peringatan keamanan dan batas sinkronisasi:** URL Web app bersifat publik. Endpoint ini menerima pengajuan publik tanpa autentikasi, sehingga dapat disalahgunakan untuk spam atau memenuhi Drive/kuota email. Jangan gunakan untuk data klien nyata sebelum menambahkan perlindungan server-side seperti autentikasi admin, pembatasan penyalahgunaan, dan kebijakan privasi. Kredensial demo sisi browser bukan perlindungan untuk deployment publik. Sinkronisasi backend saat ini hanya menerima pengajuan publik; operasi admin untuk mengubah status, posting, mengedit, atau menghapus data belum didukung backend dan tidak boleh dianggap tersinkron ke Drive.

## GitHub Pages

Repository telah terhubung ke GitHub Pages. URL situs mengikuti pola:
`https://<username>.github.io/<nama-repository>/`.

Untuk mengatur atau memeriksa deployment, buka **Settings → Pages** di repository. Pilih branch `main` dan folder root (`/`) sebagai source bila belum dikonfigurasi. Setelah perubahan masuk ke branch tersebut, tunggu workflow Pages selesai lalu muat ulang URL situs.

GitHub Pages hanya melayani file statis dan tidak menyediakan database atau autentikasi admin. Tanpa URL Google Apps Script yang aktif, data tersimpan hanya pada browser/perangkat itu sendiri dan tidak dapat terlihat oleh perangkat lain. Di situs publik, dashboard admin disembunyikan dan formulir tidak mengklaim berhasil bila backend belum dikonfigurasi. Untuk menerima pengajuan lintas perangkat, siapkan backend dengan autentikasi server-side dan kebijakan privasi sebelum menghubungkan formulir.
