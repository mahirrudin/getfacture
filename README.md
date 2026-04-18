# GetFacture

CLI tool untuk mengekstrak data dari file PDF **Faktur Pajak** Indonesia secara batch, lalu mengekspornya ke file Excel (`.xlsx`) dengan dua sheet.

## Fitur

- Membaca semua file PDF dalam satu folder secara otomatis
- Mengekstrak informasi **penjual** dan **pembeli** (nama, alamat, NPWP, NIK, email, dll.)
- Mengekstrak **nomor seri faktur pajak**
- Mengekstrak **detail barang/jasa** (nama, harga, kuantitas, kode, potongan, PPnBM)
- Mengekstrak **total** (harga jual, potongan, DPP, uang muka, PPN, PPnBM)
- Mengekstrak **informasi penandatangan** (tempat, tanggal, nama, referensi)
- Output ke Excel dengan 2 sheet:
  - **Summary Faktur** — satu baris per file PDF
  - **Detail Barang** — satu baris per item barang/jasa

## Prasyarat

### Versi Go

- [Go](https://go.dev/dl/) 1.21 atau lebih baru
- Tool eksternal: `pdftotext.exe` (dari [XpdfReader](https://www.xpdfreader.com/)) harus berada satu folder dengan `getfacture.exe` atau `main.go`. Alat ini dibutuhkan untuk mengekstrak teks PDF dengan fungsionalitas `-layout` agar struktur dan format e-Faktur tidak rusak.

---

## Instalasi

### Go

```bash
# Clone atau download repository
cd getfacture

# Download dependencies
go mod tidy

# kompilasi program menjadi exe file
go build -o getfacture.exe main.go
```

## Cara Penggunaan

### Menjalankan Program

**Go:**
```bash
# menjalankan program langsung dari code
go run main.go

# atau ketika sudah melakukan kompilasi program
./getfacture.exe
```

### Memverifikasi Output (Opsional)

Untuk membandingkan apakah file output dari implementasi Go sama dengan hasil ekspektasi (`result.xlsx`), Anda dapat menjalankan script verifikasi secara langsung:

```bash
go run verify_output.go
```

> **Catatan:** Tool akan melihat secara langsung `output.xlsx` dan membandingkannya baris per baris hingga tingkat sel melawan `result.xlsx`.

### Langkah-langkah

1. **Masukkan path folder PDF**

   Program akan meminta path folder yang berisi file-file PDF faktur pajak.

   ```
   Masukkan path folder PDF: C:\Users\user\Documents\faktur
   ```

   > Program akan memvalidasi bahwa folder tersebut ada. Jika tidak valid, akan diminta ulang.

2. **Masukkan path file output Excel**

   Tentukan lokasi dan nama file output `.xlsx`.

   ```
   Masukkan path file output Excel (.xlsx): C:\Users\user\Documents\result.xlsx
   ```

   > File harus berekstensi `.xlsx` dan folder tujuan harus sudah ada.

3. **Proses berjalan otomatis**

   Program akan memproses setiap file PDF satu per satu dan menampilkan progress:

   ```
   Total file PDF ditemukan: 15

   [1/15] Memproses: faktur_001.pdf
      -> Jumlah halaman: 2
   [2/15] Memproses: faktur_002.pdf
      -> Jumlah halaman: 1
   ...

   Proses selesai. Data disimpan di C:\Users\user\Documents\result.xlsx
   ```

4. **Tekan Enter untuk keluar**

---

## Output Excel

File Excel yang dihasilkan memiliki **2 sheet**:

### Sheet 1: Summary Faktur

| Kolom | Deskripsi |
|-------|-----------|
| `nama_file` | Nama file PDF sumber |
| `nomor_faktur` | Kode dan nomor seri faktur pajak |
| `nama_penjual` | Nama PKP penjual |
| `alamat_penjual` | Alamat penjual |
| `npwp_penjual` | NPWP penjual |
| `nama_pembeli` | Nama pembeli |
| `alamat_pembeli` | Alamat pembeli |
| `npwp_pembeli` | NPWP pembeli |
| `nik_pembeli` | NIK pembeli |
| `paspor_pembeli` | Nomor paspor pembeli |
| `iden_pembeli` | Identitas lain pembeli |
| `email_pembeli` | Email pembeli |
| `harga_jual_total` | Total harga jual / penggantian |
| `potongan_harga_total` | Total potongan harga |
| `uang_muka` | Uang muka yang telah diterima |
| `dpp` | Dasar Pengenaan Pajak |
| `ppn_total` | Jumlah PPN |
| `ppnbm_total` | Jumlah PPnBM |
| `tempat` | Tempat penandatanganan |
| `tanggal` | Tanggal penandatanganan |
| `penandatangan` | Nama penandatangan |
| `referensi` | Nomor referensi |

### Sheet 2: Detail Barang

| Kolom | Deskripsi |
|-------|-----------|
| `nomor_faktur` | Nomor faktur (relasi ke Sheet 1) |
| `nama_barang` | Nama barang/jasa |
| `harga` | Harga satuan (Rp) |
| `qty` | Kuantitas (Diformat otomatis sebagai angka desimal) |
| `kode` | Kode barang |
| `total` | Total harga (harga × qty) |
| `potongan_harga` | Potongan harga |
| `tarif_ppnbm` | Tarif PPnBM |
| `besaran_ppnbm` | Besaran PPnBM |

---

## Format PDF yang Didukung

Program ini dirancang untuk membaca **Faktur Pajak elektronik** (e-Faktur) dari DJP yang memiliki struktur standar, termasuk:

- Header: "Kode dan Nomor Seri Faktur Pajak"
- Blok "Pengusaha Kena Pajak" (penjual)
- Blok "Pembeli Barang Kena Pajak / Penerima Jasa Kena Pajak"
- Tabel barang/jasa dengan format `Rp <harga> x <qty>`
- Footer dengan total, DPP, PPN, dan informasi penandatangan

> **Catatan:** PDF yang dihasilkan dari sumber selain e-Faktur DJP mungkin tidak dapat diekstrak dengan benar.

---

## Struktur Proyek

```
getfacture/
├── main.go            # Source code versi Go utama
├── verify_output.go   # Source code utility untuk membandingkan output Go dan Python
├── getfacture.py      # Versi Python
├── pdftotext.exe      # Binary ekstensi Xpdf parser yang diperlukan oleh Go
├── getfacture.exe     # Binary executable ekstrak faktur yang sudah dicompile
├── go.mod             # Go module
├── go.sum             # Go dependencies checksum
└── README.md          # Dokumentasi
```

## Lisensi

Proyek ini bersifat private dan untuk penggunaan internal.
