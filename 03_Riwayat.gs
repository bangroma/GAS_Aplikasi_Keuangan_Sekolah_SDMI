/**
 * MENU RIWAYAT TRANSAKSI
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. RIWAYAT TRANSAKSI
// ============================================================

function getRiwayatTransaksi(filter) {
  filter = filter || {};

  const keyword = normalizeText(filter.keyword).toLowerCase();
  const tanggalMulai = normalizeText(filter.tanggal_mulai);
  const tanggalSelesai = normalizeText(filter.tanggal_selesai);
  const metode = normalizeText(filter.metode).toUpperCase();

  // Pagination
  const page = Math.max(1, parseInt(filter.page, 10) || 1);
  const limit = Math.min(Math.max(parseInt(filter.limit, 10) || 10, 5), 50);

  const headerList = getSheetDataAsObjects("Transaksi_Header");
  const detailList = getSheetDataAsObjects("Transaksi_Detail");
  const siswaList = getSheetDataAsObjects("Siswa");

  const hasil = [];

  headerList.forEach(function (trx) {
    const idTransaksi = normalizeText(trx.id_transaksi);
    const noKuitansi = normalizeText(trx.no_kuitansi || trx.nomor_kuitansi);

    if (!idTransaksi && !noKuitansi) return;

    const idSiswa = normalizeText(trx.id_siswa);

    const siswa = siswaList.find(function (s) {
      return normalizeText(s.id_siswa) === idSiswa;
    });

    const namaSiswa = siswa
      ? normalizeText(siswa.nama_lengkap || siswa.nama_siswa)
      : "-";
    const nisn = siswa ? normalizeText(siswa.nisn) : "";
    const nis = siswa ? normalizeText(siswa.nis) : "";
    const kelas = siswa ? normalizeText(siswa.kelas) : "";

    const tanggal = normalizeText(trx.tanggal_transaksi || trx.tanggal);
    const metodeTransaksi = normalizeText(
      trx.metode_pembayaran || trx.metode_bayar || trx.metode,
    ).toUpperCase();

    // Filter metode
    if (metode && metode !== "ALL" && metodeTransaksi !== metode) return;

    // Filter tanggal
    if (tanggalMulai || tanggalSelesai) {
      const tanggalOnly = tanggal.substring(0, 10);
      if (tanggalMulai && tanggalOnly < tanggalMulai) return;
      if (tanggalSelesai && tanggalOnly > tanggalSelesai) return;
    }

    // Filter keyword
    if (keyword) {
      const searchable = [
        idTransaksi,
        noKuitansi,
        namaSiswa,
        nisn,
        nis,
        kelas,
        metodeTransaksi,
        normalizeText(trx.nama_kasir),
      ]
        .join(" ")
        .toLowerCase();

      if (searchable.indexOf(keyword) === -1) return;
    }

    // Hitung jumlah item
    const jumlahItem = detailList.filter(function (d) {
      return normalizeText(d.id_transaksi) === idTransaksi;
    }).length;

    hasil.push({
      id_transaksi: idTransaksi,
      no_kuitansi: noKuitansi,
      tanggal: tanggal,
      id_siswa: idSiswa,
      nama_siswa: namaSiswa,
      nisn: nisn,
      nis: nis,
      kelas: kelas,
      total_bayar: toNumber(trx.total_bayar),
      metode_pembayaran: metodeTransaksi || "TUNAI",
      nama_kasir: normalizeText(trx.nama_kasir) || "Admin Kasir",
      catatan: normalizeText(trx.catatan),
      status: normalizeText(trx.status) || "SELESAI",
      jumlah_item: jumlahItem,
    });
  });

  // Urutkan terbaru di atas
  hasil.sort(function (a, b) {
    return String(b.tanggal).localeCompare(String(a.tanggal));
  });

  const total = hasil.length;
  const start = (page - 1) * limit;
  const data = hasil.slice(start, start + limit);

  return {
    success: true,
    data: data,
    total: total,
    page: page,
    limit: limit,
    total_pages: Math.ceil(total / limit) || 1,
  };
}

// ============================================================
// 2. DETAIL TRANSAKSI
// ============================================================

function getDetailTransaksi(idTransaksi) {
  const id = normalizeText(idTransaksi);

  if (!id) {
    throw new Error("ID transaksi wajib diisi.");
  }

  const headerList = getSheetDataAsObjects("Transaksi_Header");
  const detailList = getSheetDataAsObjects("Transaksi_Detail");
  const siswaList = getSheetDataAsObjects("Siswa");

  const trx = headerList.find(function (row) {
    return normalizeText(row.id_transaksi) === id;
  });

  if (!trx) {
    throw new Error("Transaksi " + id + " tidak ditemukan.");
  }

  const idSiswa = normalizeText(trx.id_siswa);

  const siswa = siswaList.find(function (s) {
    return normalizeText(s.id_siswa) === idSiswa;
  });

  const details = detailList
    .filter(function (detail) {
      return normalizeText(detail.id_transaksi) === id;
    })
    .map(function (detail) {
      return {
        id_detail: normalizeText(detail.id_detail),
        id_tagihan: normalizeText(detail.id_tagihan),
        nama_item: normalizeText(
          detail.nama_item_snapshot || detail.nama_item || detail.nama_tagihan,
        ),
        nominal_dibayar: toNumber(
          detail.nominal_dibayar || detail.jumlah_bayar || detail.nominal,
        ),
        diskon_tambahan: toNumber(detail.diskon_tambahan || detail.diskon),
        tanggal_transaksi: normalizeText(
          detail.tanggal_transaksi || detail.tanggal,
        ),
      };
    });

  return {
    success: true,
    transaksi: {
      id_transaksi: id,
      no_kuitansi: normalizeText(trx.no_kuitansi || trx.nomor_kuitansi),
      tanggal: normalizeText(trx.tanggal_transaksi || trx.tanggal),
      id_siswa: idSiswa,
      nama_siswa: siswa
        ? normalizeText(siswa.nama_lengkap || siswa.nama_siswa)
        : "-",
      nisn: siswa ? normalizeText(siswa.nisn) : "",
      nis: siswa ? normalizeText(siswa.nis) : "",
      kelas: siswa ? normalizeText(siswa.kelas) : "",
      total_bayar: toNumber(trx.total_bayar),
      metode_pembayaran:
        normalizeText(
          trx.metode_pembayaran || trx.metode_bayar || trx.metode,
        ) || "TUNAI",
      nama_kasir: normalizeText(trx.nama_kasir) || "Admin Kasir",
      catatan: normalizeText(trx.catatan),
      status: normalizeText(trx.status) || "SELESAI",
    },
    detail: details,
  };
}