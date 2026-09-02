/**
 * MENU LAPORAN TUNGGAKAN
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. REKAP TUNGGAKAN PER KELAS
// ============================================================

function getRekapTunggakanPerKelas(idPos) {
  const siswaList = getSheetDataAsObjects("Siswa").filter(function (s) {
    const status = normalizeText(s.status_siswa).toUpperCase();
    return status === "AKTIF";
  });

  let tagihanList = getSheetDataAsObjects("Tagihan_Siswa");

  if (idPos && normalizeText(idPos).toUpperCase() !== "ALL") {
    tagihanList = tagihanList.filter(function (t) {
      return normalizeText(t.id_pos) === normalizeText(idPos);
    });
  }

  const mapSiswaTunggakan = {};

  tagihanList.forEach(function (t) {
    const nominalAkhir = toNumber(t.nominal_akhir);
    const terbayar = toNumber(t.terbayar);
    const diskon = toNumber(t.diskon_tambahan_total);

    const pelunasanEfektif = Math.min(
      nominalAkhir,
      Math.max(0, terbayar + diskon)
    );

    const sisa = Math.max(
      0,
      nominalAkhir - pelunasanEfektif
    );

    if (sisa > 0) {
      const idSiswa = normalizeText(t.id_siswa);
      mapSiswaTunggakan[idSiswa] = (mapSiswaTunggakan[idSiswa] || 0) + sisa;
    }
  });

  const rekapKelas = {};

  siswaList.forEach(function (s) {
    const kelas = normalizeText(s.kelas) || "Tanpa Kelas";

    if (!rekapKelas[kelas]) {
      rekapKelas[kelas] = {
        kelas: kelas,
        total_siswa: 0,
        siswa_nunggak: 0,
        total_tunggakan: 0,
      };
    }

    rekapKelas[kelas].total_siswa++;

    const tunggakanSiswa = mapSiswaTunggakan[normalizeText(s.id_siswa)] || 0;

    if (tunggakanSiswa > 0) {
      rekapKelas[kelas].siswa_nunggak++;
      rekapKelas[kelas].total_tunggakan += tunggakanSiswa;
    }
  });

  return Object.values(rekapKelas).sort(function (a, b) {
    return a.kelas.localeCompare(b.kelas, undefined, {
      numeric: true,
      sensitivity: "base",
    });
  });
}

// ============================================================
// 2. GET LIST KELAS UNTUK FILTER LAPORAN
// ============================================================

function getListKelasForLaporan() {
  const siswaList = getSheetDataAsObjects("Siswa").filter(function (s) {
    const status = normalizeText(s.status_siswa).toUpperCase();
    return status === "AKTIF";
  });

  const kelasSet = new Set();

  siswaList.forEach(function (s) {
    const kelas = normalizeText(s.kelas);
    if (kelas) {
      kelasSet.add(kelas);
    }
  });

  return Array.from(kelasSet).sort(function (a, b) {
    return a.localeCompare(b, undefined, {
      numeric: true,
      sensitivity: "base",
    });
  });
}

// ============================================================
// 3. GET POS LIST UNTUK FILTER LAPORAN
// ============================================================

function getPosListForLaporan() {
  try {
    const posList = getSheetDataAsObjects("Master_Pos");
    return posList
      .filter(function (p) {
        return toBoolean(p.aktif);
      })
      .map(function (p) {
        return {
          id_pos: normalizeText(p.id_pos),
          kode_pos: normalizeText(p.kode_pos),
          nama_pos: normalizeText(p.nama_pos)
        };
      });
  } catch (err) {
    return [];
  }
}

// ============================================================
// 4. GET SISWA BY KELAS UNTUK LAPORAN
// ============================================================

function getSiswaByKelasForLaporan(kelas, idPos) {
  kelas = normalizeText(kelas);
  idPos = normalizeText(idPos);

  if (!kelas) {
    return [];
  }

  const siswaList = getSheetDataAsObjects("Siswa");
  let tagihanList = getSheetDataAsObjects("Tagihan_Siswa");

  // Filter siswa aktif
  const siswaAktif = siswaList.filter(function (s) {
    return normalizeText(s.status_siswa).toUpperCase() === "AKTIF";
  });

  const siswaKelas = siswaAktif.filter(function (s) {
    return normalizeText(s.kelas) === kelas;
  });

  const siswaKelasIds = {};
  siswaKelas.forEach(function (siswa) {
    siswaKelasIds[normalizeText(siswa.id_siswa)] = true;
  });

  // Filter tagihan berdasarkan pos
  if (idPos && idPos !== "ALL") {
    tagihanList = tagihanList.filter(function (t) {
      return normalizeText(t.id_pos) === idPos;
    });
  }

  const tagihanBySiswa = {};

  tagihanList.forEach(function (tagihan) {
    const idSiswa = normalizeText(tagihan.id_siswa);

    if (!siswaKelasIds[idSiswa]) return;

    if (!tagihanBySiswa[idSiswa]) {
      tagihanBySiswa[idSiswa] = [];
    }

    tagihanBySiswa[idSiswa].push(tagihan);
  });

  const hasil = [];

  siswaKelas.forEach(function (siswa) {
    const idSiswa = normalizeText(siswa.id_siswa);
    const tagihanSiswa = tagihanBySiswa[idSiswa] || [];

    if (tagihanSiswa.length === 0) return;

    let totalTagihan = 0;
    let totalDibayar = 0;
    let totalDiskon = 0;
    let totalEfektif = 0;
    let totalSisa = 0;

    const detailTagihan = tagihanSiswa.map(function (t) {
      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const diskon = toNumber(t.diskon_tambahan_total);

      const efektif = Math.min(
        nominalAkhir,
        Math.max(0, terbayar + diskon)
      );

      const sisa = Math.max(
        0,
        nominalAkhir - efektif
      );

      const progress =
        nominalAkhir > 0
          ? Math.min(100, Math.round((efektif / nominalAkhir) * 100))
          : 100;

      const status =
        sisa <= 0
          ? "LUNAS"
          : efektif > 0
            ? "CICILAN"
            : "BELUM BAYAR";

      totalTagihan += nominalAkhir;
      totalDibayar += terbayar;
      totalDiskon += diskon;
      totalEfektif += efektif;
      totalSisa += sisa;

      return {
        id_tagihan: normalizeText(t.id_tagihan),
        id_pos: normalizeText(t.id_pos),
        nama_item: normalizeText(
          t.nama_item_snapshot || t.nama_item || "Tagihan",
        ),
        periode: normalizeText(t.periode),
        nominal_akhir: nominalAkhir,
        terbayar: terbayar,
        diskon_tambahan_total: diskon,
        pelunasan_efektif: efektif,
        sisa_tunggakan: sisa,
        progress: progress,
        status: status,
      };
    });

    const progressTotal =
      totalTagihan > 0
        ? Math.min(100, Math.round((totalEfektif / totalTagihan) * 100))
        : 100;

    const statusTotal =
      totalSisa <= 0
        ? "LUNAS"
        : totalEfektif > 0
          ? "CICILAN"
          : "BELUM BAYAR";

    hasil.push({
      id_siswa: idSiswa,
      nisn: normalizeText(siswa.nisn),
      nis: normalizeText(siswa.nis),
      nama_lengkap: normalizeText(siswa.nama_lengkap),
      kelas: normalizeText(siswa.kelas),
      total_tagihan: totalTagihan,
      total_dibayar: totalDibayar,
      total_diskon: totalDiskon,
      total_pelunasan_efektif: totalEfektif,
      total_sisa: totalSisa,
      progress: progressTotal,
      status: statusTotal,
      tagihan: detailTagihan,
    });
  });

  hasil.sort(function (a, b) {
    return a.nama_lengkap.localeCompare(b.nama_lengkap);
  });

  return hasil;
}

// ============================================================
// 5. EKSPOR REKAP TUNGGAKAN
// ============================================================

// ============================================================
// 5. EKSPOR REKAP TUNGGAKAN (DOWNLOAD EXCEL)
// ============================================================

function exportLaporanTunggakanToExcel(idPos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var timezone = Session.getScriptTimeZone() || "Asia/Jakarta";
    var dateStr = Utilities.formatDate(now, timezone, "yyyyMMdd_HHmmss");
    var fileName = "Rekap_Tunggakan_" + dateStr;
    var namaPosText = getNamaPos(idPos);
    var rekapKelas = getRekapTunggakanPerKelas(idPos);

    if (!rekapKelas || rekapKelas.length === 0) {
      throw new Error("Tidak ada data tunggakan untuk diekspor.");
    }

    // Buat spreadsheet temporer
    var tempSS = SpreadsheetApp.create(fileName);
    var sheet = tempSS.getActiveSheet();
    sheet.setName("Rekap Tunggakan");

    // Header
    var titleRange = sheet.getRange("A1:D1");
    titleRange.merge();
    titleRange.setValue("REKAPITULASI TUNGGAKAN PEMBAYARAN PER KELAS");
    titleRange.setFontWeight("bold");
    titleRange.setFontSize(13);
    titleRange.setHorizontalAlignment("center");

    var subRange = sheet.getRange("A2:D2");
    subRange.merge();
    subRange.setValue("Pos Pembayaran: " + namaPosText);
    subRange.setFontWeight("bold");
    subRange.setFontSize(10);
    subRange.setHorizontalAlignment("center");

    var dateRange = sheet.getRange("A3:D3");
    dateRange.merge();
    dateRange.setValue(
      "Tanggal Cetak: " + Utilities.formatDate(now, timezone, "dd MMMM yyyy HH:mm:ss")
    );
    dateRange.setFontSize(9);
    dateRange.setFontStyle("italic");
    dateRange.setHorizontalAlignment("center");

    // Header Tabel
    var headerRange = sheet.getRange(5, 1, 1, 4);
    headerRange.setValues([
      ["Kelas", "Total Siswa", "Siswa Menunggak", "Total Nominal Tunggakan"],
    ]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#e2e8f0");
    headerRange.setHorizontalAlignment("center");

    // Data
    var rowData = rekapKelas.map(function (r) {
      return [r.kelas, r.total_siswa, r.siswa_nunggak, r.total_tunggakan];
    });

    var grandTotal = 0;
    var grandSiswaNunggak = 0;

    rekapKelas.forEach(function (r) {
      grandTotal += r.total_tunggakan;
      grandSiswaNunggak += r.siswa_nunggak;
    });

    if (rowData.length > 0) {
      var dataRange = sheet.getRange(6, 1, rowData.length, 4);
      dataRange.setValues(rowData);
      
      var nominalRange = sheet.getRange(6, 4, rowData.length, 1);
      nominalRange.setNumberFormat('"Rp" #,##0');
      nominalRange.setHorizontalAlignment("right");
    }

    var totalRow = 6 + rowData.length;

    // Total baris
    var totalRange = sheet.getRange(totalRow, 1, 1, 2);
    totalRange.merge();
    totalRange.setValue("TOTAL KESELURUHAN");
    totalRange.setFontWeight("bold");
    totalRange.setBackground("#cbd5e1");
    totalRange.setHorizontalAlignment("center");

    var totalSiswaRange = sheet.getRange(totalRow, 3);
    totalSiswaRange.setValue(grandSiswaNunggak);
    totalSiswaRange.setFontWeight("bold");
    totalSiswaRange.setBackground("#cbd5e1");
    totalSiswaRange.setHorizontalAlignment("center");

    var totalNominalRange = sheet.getRange(totalRow, 4);
    totalNominalRange.setValue(grandTotal);
    totalNominalRange.setFontWeight("bold");
    totalNominalRange.setNumberFormat('"Rp" #,##0');
    totalNominalRange.setBackground("#cbd5e1");
    totalNominalRange.setHorizontalAlignment("right");

    sheet.autoResizeColumns(1, 4);

    // Flush untuk memastikan semua data tersimpan
    SpreadsheetApp.flush();

    // Dapatkan ID file dan buat URL download
    var fileId = tempSS.getId();
    var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: fileName + ".xlsx",
      message: "Laporan rekap tunggakan siap diunduh.",
      totalData: rekapKelas.length
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}



// ============================================================
// 6. GET NAMA POS
// ============================================================

function getNamaPos(idPos) {
  if (!idPos || normalizeText(idPos).toUpperCase() === "ALL") {
    return "Semua Jenis Iuran";
  }

  const listPos = getPosPembayaranList();

  const posObj = listPos.find(function (p) {
    return normalizeText(p.id_pos) === normalizeText(idPos);
  });

  return posObj ? normalizeText(posObj.nama_pos) : "Pos " + idPos;
}

// ============================================================
// 7. DETAIL TUNGGAKAN PER KELAS
// ============================================================

function getDetailTunggakanPerKelas(kelas, idPos) {
  kelas = normalizeText(kelas);
  idPos = normalizeText(idPos);

  if (!kelas) throw new Error("Kelas wajib diisi.");

  const siswaList = getSheetDataAsObjects("Siswa");
  let tagihanList = getSheetDataAsObjects("Tagihan_Siswa");

  const siswaAktif = siswaList.filter(function (s) {
    return normalizeText(s.status_siswa).toUpperCase() === "AKTIF";
  });

  const siswaKelas = siswaAktif.filter(function (s) {
    return normalizeText(s.kelas) === kelas;
  });

  const siswaKelasIds = {};

  siswaKelas.forEach(function (siswa) {
    siswaKelasIds[normalizeText(siswa.id_siswa)] = true;
  });

  if (idPos && idPos !== "ALL") {
    tagihanList = tagihanList.filter(function (t) {
      return normalizeText(t.id_pos) === idPos;
    });
  }

  const tagihanBySiswa = {};

  tagihanList.forEach(function (tagihan) {
    const idSiswa = normalizeText(tagihan.id_siswa);

    if (!siswaKelasIds[idSiswa]) return;

    if (!tagihanBySiswa[idSiswa]) {
      tagihanBySiswa[idSiswa] = [];
    }

    tagihanBySiswa[idSiswa].push(tagihan);
  });

  const hasil = [];

  siswaKelas.forEach(function (siswa) {
    const idSiswa = normalizeText(siswa.id_siswa);
    const tagihanSiswa = tagihanBySiswa[idSiswa] || [];

    if (tagihanSiswa.length === 0) return;

    let totalTagihan = 0;
    let totalDibayar = 0;
    let totalEfektif = 0;
    let totalSisa = 0;

    const detailTagihan = tagihanSiswa.map(function (t) {
      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const diskon = toNumber(t.diskon_tambahan_total);

      const efektif = Math.min(
        nominalAkhir,
        Math.max(0, terbayar + diskon)
      );

      const sisa = Math.max(
        0,
        nominalAkhir - efektif
      );

      totalTagihan += nominalAkhir;
      totalDibayar += terbayar;
      totalEfektif += efektif;
      totalSisa += sisa;

      const progress =
        nominalAkhir > 0
          ? Math.min(100, Math.round((efektif / nominalAkhir) * 100))
          : 100;

      const status =
        sisa <= 0
          ? "LUNAS"
          : efektif > 0
            ? "CICILAN"
            : "BELUM BAYAR";

      return {
        id_tagihan: normalizeText(t.id_tagihan),
        id_pos: normalizeText(t.id_pos),
        nama_item: normalizeText(
          t.nama_item_snapshot || t.nama_item || "Tagihan",
        ),
        periode: normalizeText(t.periode),
        nominal_akhir: nominalAkhir,
        terbayar: terbayar,
        diskon_tambahan_total: diskon,
        pelunasan_efektif: efektif,
        sisa_tunggakan: sisa,
        progress: progress,
        status: status,
      };
    });

    if (totalSisa <= 0) return;

    const progressTotal =
      totalTagihan > 0
        ? Math.min(100, Math.round((totalEfektif / totalTagihan) * 100))
        : 100;

    const statusTotal =
      totalSisa <= 0
        ? "LUNAS"
        : totalEfektif > 0
          ? "CICILAN"
          : "BELUM BAYAR";

    hasil.push({
      id_siswa: idSiswa,
      nisn: normalizeText(siswa.nisn),
      nis: normalizeText(siswa.nis),
      nama_lengkap: normalizeText(siswa.nama_lengkap),
      kelas: normalizeText(siswa.kelas),
      total_tagihan: totalTagihan,
      total_dibayar: totalDibayar,
      total_diskon: Math.max(0, totalEfektif - totalDibayar),
      total_pelunasan_efektif: totalEfektif,
      total_sisa: totalSisa,
      progress: progressTotal,
      status: statusTotal,
      tagihan: detailTagihan,
    });
  });

  hasil.sort(function (a, b) {
    return a.nama_lengkap.localeCompare(b.nama_lengkap);
  });

  return hasil;
}

// ============================================================
// 8. EXPORT DETAIL TUNGGAKAN TO SHEET
// ============================================================
// ============================================================
// EXPORT DETAIL TUNGGAKAN KE EXCEL (DOWNLOAD LANGSUNG)
// ============================================================

function exportDetailTunggakanToExcel(kelas, idPos) {
  try {
    kelas = normalizeText(kelas);
    idPos = normalizeText(idPos) || "ALL";

    if (!kelas) {
      throw new Error("Kelas wajib diisi.");
    }

    var data = getDetailTunggakanPerKelas(kelas, idPos);

    if (!data || data.length === 0) {
      throw new Error("Tidak ada data siswa menunggak untuk diekspor.");
    }

    var now = new Date();
    var timestamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "yyyyMMdd_HHmmss"
    );

    var safeKelas = kelas.replace(/[\\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
    var fileName = "Tunggakan_" + safeKelas + "_" + timestamp;

    // Buat spreadsheet temporer
    var tempSS = SpreadsheetApp.create(fileName);
    var sheet = tempSS.getActiveSheet();
    sheet.setName("Detail Tunggakan");

    // Header Laporan
    var titleRange = sheet.getRange("A1:H1");
    titleRange.merge();
    titleRange.setValue("LAPORAN DETAIL TUNGGAKAN SISWA");
    titleRange.setFontWeight("bold");
    titleRange.setFontSize(16);
    titleRange.setHorizontalAlignment("center");

    var subRange = sheet.getRange("A2:H2");
    subRange.merge();
    subRange.setValue("Kelas: " + kelas);
    subRange.setFontWeight("bold");
    subRange.setFontSize(12);
    subRange.setHorizontalAlignment("center");

    var dateRange = sheet.getRange("A3:H3");
    dateRange.merge();
    dateRange.setValue(
      "Tanggal Cetak: " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    );
    dateRange.setFontSize(9);
    dateRange.setFontStyle("italic");
    dateRange.setHorizontalAlignment("center");

    // Header Tabel Siswa
    var headers = [
      "No",
      "NISN",
      "NIS",
      "Nama Siswa",
      "Total Tagihan",
      "Sudah Dibayar",
      "Sisa Tunggakan",
      "Progress",
    ];

    var headerRange = sheet.getRange(5, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#d1fae5");
    headerRange.setHorizontalAlignment("center");

    // Data Siswa
    var rows = [];
    data.forEach(function (siswa, index) {
      rows.push([
        index + 1,
        siswa.nisn || "-",
        siswa.nis || "-",
        siswa.nama_lengkap || "-",
        toNumber(siswa.total_tagihan),
        toNumber(siswa.total_dibayar),
        toNumber(siswa.total_sisa),
        toNumber(siswa.progress) / 100,
      ]);
    });

    if (rows.length > 0) {
      var dataRange = sheet.getRange(6, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Format nominal
      sheet.getRange(6, 5, rows.length, 3).setNumberFormat('"Rp" #,##0');
      sheet.getRange(6, 8, rows.length, 1).setNumberFormat("0%");
    }

    // Detail Tagihan
    var detailStartRow = 6 + rows.length + 3;

    var detailTitleRange = sheet.getRange(detailStartRow, 1, 1, 7);
    detailTitleRange.merge();
    detailTitleRange.setValue("DETAIL TAGIHAN SISWA");
    detailTitleRange.setFontWeight("bold");
    detailTitleRange.setFontSize(12);
    detailTitleRange.setBackground("#e2e8f0");
    detailTitleRange.setHorizontalAlignment("center");

    detailStartRow++;

    var detailHeaders = [
      "No",
      "Nama Siswa",
      "ID Tagihan",
      "Item Tagihan",
      "Total Tagihan",
      "Sudah Dibayar",
      "Sisa Tunggakan",
    ];

    var detailHeaderRange = sheet.getRange(detailStartRow, 1, 1, detailHeaders.length);
    detailHeaderRange.setValues([detailHeaders]);
    detailHeaderRange.setFontWeight("bold");
    detailHeaderRange.setBackground("#e2e8f0");
    detailHeaderRange.setHorizontalAlignment("center");

    detailStartRow++;

    var detailRows = [];
    var detailNo = 1;

    data.forEach(function (siswa) {
      (siswa.tagihan || []).forEach(function (t) {
        detailRows.push([
          detailNo++,
          siswa.nama_lengkap || "-",
          t.id_tagihan || "-",
          t.nama_item || "-",
          toNumber(t.nominal_akhir),
          toNumber(t.terbayar),
          toNumber(t.sisa_tunggakan),
        ]);
      });
    });

    if (detailRows.length > 0) {
      var detailDataRange = sheet.getRange(detailStartRow, 1, detailRows.length, detailHeaders.length);
      detailDataRange.setValues(detailRows);
      
      var detailNominalRange = sheet.getRange(detailStartRow, 5, detailRows.length, 3);
      detailNominalRange.setNumberFormat('"Rp" #,##0');
      detailNominalRange.setHorizontalAlignment("right");
    }

    // Auto resize columns
    sheet.autoResizeColumns(1, headers.length);
    sheet.setColumnWidth(4, 250);

    // Flush
    SpreadsheetApp.flush();

    // URL Download
    var fileId = tempSS.getId();
    var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: fileName + ".xlsx",
      message: "Laporan detail tunggakan kelas " + kelas + " siap diunduh.",
      totalData: data.length
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}