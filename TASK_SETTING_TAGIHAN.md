# Roadmap — Setting Tagihan Antar Tahun Pelajaran

Status: PLANNED
Prioritas: Tinggi

## Tujuan

Memudahkan operator membuat setting/tagihan untuk tahun pelajaran baru tanpa mengatur ulang satu per satu, sekaligus menyediakan mekanisme Excel sebagai backup dan pertukaran template.

## Prinsip Desain

- Tidak membuat Sheet/database baru pada tahap pertama.
- Sumber pola setting diambil dari Tagihan_Siswa.
- Setting disalin berdasarkan pola Kelas + Pos + Periode.
- Tagihan tahun tujuan dibuat sebagai data baru.
- Data pembayaran tahun sumber tidak boleh ikut tersalin.
- Tagihan yang sudah ada di tahun tujuan tidak ditimpa secara default.
- Siswa aktif pada tahun tujuan mengikuti kelas/data tahun tujuan.

## Tahap 1 — Audit & Desain

- [x] Audit struktur Setting Tagihan
- [x] Pastikan tidak diperlukan schema baru
- [x] Identifikasi Tagihan_Siswa sebagai sumber pola setting
- [x] Tentukan data pembayaran tidak ikut disalin
- [x] Tentukan mapping berdasarkan kelas

## Tahap 2 — Salin Setting Tahun Pelajaran

- [ ] Pilih tahun pelajaran sumber
- [ ] Pilih tahun pelajaran tujuan
- [ ] Buat fungsi preview
- [ ] Tampilkan jumlah kelas, pola setting, dan siswa
- [ ] Deteksi setting yang sudah ada
- [ ] Default duplikasi: skip
- [ ] Buat tagihan baru dengan terbayar = 0
- [ ] Set diskon_tambahan_total = 0
- [ ] Hitung ulang sisa_tunggakan
- [ ] Set status awal sesuai nominal tagihan
- [ ] Jangan menyalin ID tagihan/transaksi/pembayaran
- [ ] Terapkan ke siswa aktif pada kelas tujuan
- [ ] Test 2026/2027 → 2027/2028
- [ ] clasp push
- [ ] Test di GAS
- [ ] Git commit/checkpoint

## Tahap 3 — Export Template Setting ke Excel

- [ ] Tombol Export Template Setting
- [ ] Export pola setting, bukan pembayaran
- [ ] Kelas
- [ ] Pos pembayaran
- [ ] Periode
- [ ] Tarif/nominal
- [ ] Potongan
- [ ] Opsi cicilan
- [ ] Keterangan
- [ ] Download .xlsx
- [ ] Test hasil Excel

## Tahap 4 — Import Template Setting dari Excel

- [ ] Tombol Import Template Setting
- [ ] Validasi struktur Excel
- [ ] Validasi tahun tujuan
- [ ] Validasi kelas
- [ ] Validasi pos pembayaran
- [ ] Validasi nominal dan potongan
- [ ] Preview sebelum import
- [ ] Deteksi duplikasi
- [ ] Skip/overwrite/cancel bila diperlukan
- [ ] Generate tagihan baru
- [ ] Pastikan pembayaran dimulai dari nol
- [ ] Test import
- [ ] Git checkpoint

## Tahap 5 — Finalisasi

- [ ] Audit Setting Tagihan
- [ ] Audit alur pembayaran
- [ ] Audit Laporan Tunggakan
- [ ] Audit Laporan Siswa Lunas
- [ ] Audit Riwayat Setting
- [ ] Dokumentasikan format Excel
- [ ] Checkpoint/tag Git final

## Urutan Implementasi

1. Salin Setting Antar Tahun
2. Test dan checkpoint
3. Export Template Excel
4. Test dan checkpoint
5. Import Template Excel
6. Test dan checkpoint
7. Audit dan final checkpoint

## Catatan Keamanan

### Boleh disalin

- kelas
- pos pembayaran
- periode
- nominal/tarif
- potongan
- aturan cicilan
- keterangan yang relevan

### Tidak boleh disalin

- id_tagihan
- id_transaksi
- id_detail
- terbayar
- diskon_tambahan_total
- sisa_tunggakan tahun sumber
- status pembayaran tahun sumber
- histori transaksi

## Checkpoint Git

Setiap tahap yang sudah berhasil diuji harus disimpan dengan commit/checkpoint sebelum melanjutkan tahap berikutnya.
