/**
 * DASHBOARD
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. CARI SISWA (untuk Dashboard)
// ============================================================

function cariSiswa(keyword) {
  const kw = normalizeText(keyword).toLowerCase();

  if (!kw) {
    return [];
  }

  const siswaList = getSheetDataAsObjects("Siswa");

  const potonganList = getSheetDataAsObjects("Master_Potongan");

  return siswaList
    .filter(function (s) {
      const status = normalizeText(s.status || s.status_siswa).toUpperCase();

      return status === "AKTIF";
    })
    .filter(function (s) {
      return (
        normalizeText(s.nama_lengkap || s.nama_siswa)
          .toLowerCase()
          .includes(kw) ||
        normalizeText(s.nisn).toLowerCase().includes(kw) ||
        normalizeText(s.nis).toLowerCase().includes(kw)
      );
    })
    .map(function (s) {
      const potongan = potonganList.find(function (p) {
        return (
          normalizeText(p.id_potongan) === normalizeText(s.id_potongan_default)
        );
      });

      return {
        id_siswa: s.id_siswa,
        nisn: s.nisn,
        nis: s.nis,
        nama_lengkap: s.nama_lengkap || s.nama_siswa || "",
        kelas: s.kelas,
        tahun_pelajaran: s.tahun_pelajaran,
        id_potongan_default: s.id_potongan_default || "",
        nama_potongan: potongan ? potongan.nama_potongan : "Tidak Ada",
        tipe_potongan: potongan ? potongan.tipe_nilai : "NONE",
        nilai_potongan: potongan ? toNumber(potongan.nilai_potongan) : 0,
      };
    });
}

// ============================================================
// 2. DAFTAR POS PEMBAYARAN
// ============================================================

function getPosPembayaranList() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pos_Pembayaran") ||
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Pos");

  if (!sheet) {
    throw new Error(
      "Sheet 'Pos_Pembayaran' atau 'Master_Pos' tidak ditemukan.",
    );
  }

  return getSheetDataAsObjects(sheet.getName()).filter(function (row) {
    return row.aktif === undefined || toBoolean(row.aktif);
  });
}