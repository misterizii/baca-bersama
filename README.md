# Baca Bersama

Aplikasi belajar membaca untuk **SD Negeri Geger, Kedungadem - Bojonegoro**.
Guru dan orang tua murid memantau perkembangan membaca anak bersama-sama. Kepala sekolah menjadi **admin**.

## Cara kerjanya

```
HP (PWA di Vercel)  ──fetch JSON──▶  Google Apps Script (Web App)  ──▶  Google Sheets + Google Drive
   index.html                           Code.gs                          database & file lampiran
```

- **Frontend** (`index.html`, `sw.js`, `manifest.webmanifest`): satu halaman, tanpa library, tanpa font luar, tanpa Tailwind CDN. Halaman utamanya hanya sekitar 58 KB (jauh lebih kecil saat dikompres Vercel), sehingga cepat dibuka di sinyal lemah. Bisa dipasang ke layar utama HP.
- **Backend** (`Code.gs`): Google Apps Script sebagai API. Semua aturan peran dan **isolasi kelas dicek di server**.
- **Database**: Google Sheets. **File lampiran**: folder Drive `LMS_Membaca_Files`.

## Peran

| Peran | Akses |
|---|---|
| `admin` (kepala sekolah) | Semua kelas: membuat tugas, melihat laporan, hasil soal, diskusi. Mengelola akun dan memetakan guru/orang tua ke kelas. |
| `guru` | Hanya kelasnya sendiri. |
| `ortu` | Hanya kelasnya sendiri. Balas diskusi hanya pada pesan yang menyebut dirinya. |

## Langkah 1 — Pasang backend (Google Apps Script)

1. Buat **Google Sheet baru** (kosong) di akun Google sekolah.
2. Buka **Extensions > Apps Script**.
3. Hapus isi `Code.gs` bawaan, lalu tempel seluruh isi file `Code.gs` dari proyek ini. Simpan.
4. Pilih fungsi **`setupDatabase`** di menu fungsi, klik **Run**. Setujui izin (Sheets dan Drive). Ini membuat semua tab, folder Drive, dan akun contoh.
5. Klik **Deploy > New deployment > Web app**:
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
6. Salin **Web app URL** (berakhiran `/exec`).

> Setiap kali `Code.gs` diubah, pilih **Deploy > Manage deployments > Edit > New version** agar perubahan aktif. URL tidak berubah.

## Langkah 2 — Hubungkan frontend

Buka `index.html`, cari baris ini dan ganti dengan URL dari langkah 1:

```js
var API_URL = 'GANTI_DENGAN_URL_WEB_APP_ANDA';
```

## Langkah 3 — Deploy ke GitHub lalu Vercel

1. Buat repository GitHub baru dan unggah semua file proyek ini. `Code.gs` dan `README.md` ikut diunggah tidak masalah, karena Vercel mengabaikannya lewat `.vercelignore`.
2. Di [vercel.com](https://vercel.com): **Add New > Project**, pilih repository tersebut.
3. *Framework Preset*: **Other**. Kosongkan Build Command dan Output Directory. Klik **Deploy**.
4. Buka alamat `.vercel.app` yang diberikan di HP, lalu **Pasang aplikasi** (menu ☰ di aplikasi, atau menu browser "Tambahkan ke layar utama").

## Akun contoh (segera ganti!)

| Username | PIN | Peran | Kelas |
|---|---|---|---|
| `kepsek` | `123456` | admin | semua |
| `guru1` | `111111` | guru | 1A |
| `ortu1` | `222222` | ortu (anak: Budi) | 1A |
| `ortu2` | `333333` | ortu (anak: Sari) | 2B |

Login sebagai `kepsek`, buka menu **Pengguna**, lalu ganti PIN semua akun contoh (atau hapus akun contoh dan buat akun asli). Setiap pengguna dapat mengganti PIN sendiri lewat menu ☰.

## Membuat akun dan memetakan kelas (admin)

Menu **Pengguna > + Tambah Pengguna**:
- **Guru**: isi nama, kelas, username, PIN.
- **Orang tua**: isi nama, **nama anak**, kelas, username, PIN.
- Memindahkan guru/orang tua ke kelas lain: **Ubah / Reset PIN**, ganti kelasnya.
- Lupa PIN: **Ubah / Reset PIN**, isi PIN baru.

## Catatan keamanan

- PIN disimpan sebagai hash SHA-256 (dengan kunci rahasia yang dibuat otomatis di Script Properties). Salah PIN 5 kali mengunci login 10 menit.
- Sesi berupa token bertanda tangan, berlaku 14 hari. Pilihan "Ingat saya" menyimpan sesi di HP; tanpa itu sesi hilang saat tab ditutup.
- Kunci jawaban pilihan ganda **tidak pernah dikirim** ke orang tua.
- Lampiran bisa dibuka siapa saja yang memegang tautannya (sesuai permintaan: *anyone with link*).
- Jangan bagikan akses edit Google Sheet kepada siapa pun selain admin.

## Struktur file

```
index.html            aplikasi (HTML + CSS + JS dalam satu file)
sw.js                 service worker (cache halaman agar cepat & tahan sinyal lemah)
manifest.webmanifest  identitas aplikasi untuk dipasang di HP
logo.png, icon-*.png  logo SDN Geger dan ikon aplikasi
vercel.json           header cache & keamanan
Code.gs               backend (ditempel ke Apps Script, bukan diunggah ke Vercel)
```
