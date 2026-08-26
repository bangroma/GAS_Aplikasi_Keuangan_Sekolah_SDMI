/**
 * MENU KASIR / PEMBAYARAN
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. AMBIL TAGIHAN SISWA
// ============================================================

function getTagihanSiswa(idSiswa) {
  const siswaId = normalizeText(idSiswa);

  if (!siswaId) {
    return [];
  }

  const tagihanList = getSheetDataAsObjects("Tagihan_Siswa");
  const posList = getSheetDataAsObjects("Master_Pos");
  const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

  return tagihanList
    .filter(function (t) {
      return normalizeText(t.id_siswa) === siswaId;
    })
    .map(function (t) {
      const pos = posList.find(function (p) {
        return normalizeText(p.id_pos) === normalizeText(t.id_pos);
      });

      const nominalAsal = toNumber(t.nominal_asal);
      const nominalPotongan = toNumber(t.nominal_potongan);
      const nominalAkhirStored = toNumber(t.nominal_akhir);

      const nominalAkhir =
        nominalAkhirStored > 0
          ? nominalAkhirStored
          : Math.max(0, nominalAsal - nominalPotongan);

      const terbayar = toNumber(t.terbayar);
      const sisaTunggakan = Math.max(0, nominalAkhir - terbayar);

      // PERIODE BERSIH
      let periodeBersih = "";
      const rawPeriode = t.periode;

      if (rawPeriode instanceof Date && !isNaN(rawPeriode.getTime())) {
        periodeBersih = Utilities.formatDate(rawPeriode, timezone, "yyyy-MM");
      } else {
        periodeBersih = normalizeText(rawPeriode);
      }

      // NAMA ITEM BERSIH
      let namaPosBersih = "";
      if (pos) {
        namaPosBersih = normalizeText(pos.nama_pos);
      } else {
        let snapshot = normalizeText(t.nama_item_snapshot);
        snapshot = snapshot.replace(/\s*\([^)]*GMT[^)]*\)/gi, "").trim();
        namaPosBersih = snapshot || normalizeText(t.id_pos);
      }

      let status = normalizeText(t.status).toUpperCase();
      if (sisaTunggakan <= 0) {
        status = "LUNAS";
      } else if (terbayar > 0) {
        status = "CICILAN";
      } else {
        status = "BELUM_BAYAR";
      }

      return {
        id_tagihan: normalizeText(t.id_tagihan),
        id_siswa: normalizeText(t.id_siswa),
        id_pos: normalizeText(t.id_pos),
        nama_pos: namaPosBersih,
        periode: periodeBersih,
        tipe_pos: pos ? normalizeText(pos.tipe_pos) : "BULANAN",
        bisa_dicicil: pos
          ? toBoolean(pos.bisa_dicicil)
          : toBoolean(t.bisa_dicicil),
        nominal_asal: nominalAsal,
        nominal_potongan: nominalPotongan,
        nominal_akhir: nominalAkhir,
        terbayar: terbayar,
        sisa_tunggakan: sisaTunggakan,
        status: status,
      };
    })
    .filter(function (t) {
      return t.sisa_tunggakan > 0;
    });
}

function bersihkanPeriodeTagihan() {
  const sheet = getSheet("Tagihan_Siswa");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) {
    return normalizeText(h);
  });

  const colPeriode = headers.indexOf("periode");
  const colSnapshot = headers.indexOf("nama_item_snapshot");

  if (colPeriode === -1) {
    Logger.log("Kolom periode tidak ditemukan");
    return;
  }

  const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const raw = data[i][colPeriode];

    if (raw instanceof Date && !isNaN(raw.getTime())) {
      const bersih = Utilities.formatDate(raw, timezone, "yyyy-MM");
      sheet.getRange(i + 1, colPeriode + 1).setValue(bersih);
      updated++;
    }

    if (colSnapshot !== -1) {
      let snapshot = data[i][colSnapshot];
      if (typeof snapshot === "string" && snapshot.indexOf("GMT") !== -1) {
        const bersihSnapshot = snapshot
          .replace(/\s*\([^)]*GMT[^)]*\)/gi, "")
          .trim();
        sheet.getRange(i + 1, colSnapshot + 1).setValue(bersihSnapshot);
        updated++;
      }
    }
  }

  Logger.log("Berhasil membersihkan " + updated + " baris.");
}

// ============================================================
// 2. SIMPAN TRANSAKSI
// ============================================================

function simpanTransaksi(payload) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  let headerRow = null;
  let detailStartRow = null;
  let detailRowCount = 0;

  const originalTagihan = [];

  try {
    // VALIDASI PAYLOAD
    if (!payload || typeof payload !== "object") {
      throw new Error("Payload transaksi kosong atau tidak valid.");
    }

    const idSiswa = normalizeText(payload.id_siswa);

    if (!idSiswa) {
      throw new Error("ID siswa wajib diisi.");
    }

    if (!Array.isArray(payload.items) || payload.items.length === 0) {
      throw new Error("Data transaksi kosong.");
    }

    const items = payload.items.map(function (item, index) {
      if (!item || typeof item !== "object") {
        throw new Error("Item transaksi ke-" + (index + 1) + " tidak valid.");
      }

      const idTagihan = normalizeText(item.id_tagihan);

      const nominalDibayar = toNumber(item.nominal_dibayar);

      if (!idTagihan) {
        throw new Error(
          "ID tagihan pada item ke-" + (index + 1) + " wajib diisi.",
        );
      }

      if (nominalDibayar <= 0) {
        throw new Error(
          "Nominal pembayaran pada tagihan " +
            idTagihan +
            " harus lebih dari 0.",
        );
      }

      return {
        id_tagihan: idTagihan,

        nama_item_snapshot: normalizeText(item.nama_item_snapshot),

        nominal_dibayar: nominalDibayar,

        diskon_tambahan: Math.max(0, toNumber(item.diskon_tambahan)),
      };
    });

    const uniqueIds = new Set(
      items.map(function (item) {
        return item.id_tagihan;
      }),
    );

    if (uniqueIds.size !== items.length) {
      throw new Error(
        "Satu tagihan tidak boleh muncul lebih dari satu kali dalam transaksi.",
      );
    }

    // SHEET
    const headerSheet = getSheet("Transaksi_Header");

    const detailSheet = getSheet("Transaksi_Detail");

    const tagihanSheet = getSheet("Tagihan_Siswa");

    const siswaSheet = getSheet("Siswa");

    const posSheet = getSheet("Master_Pos");

    // STRUKTUR HEADER
    const headerMap = getHeaderMap(headerSheet);

    const detailMap = getHeaderMap(detailSheet);

    const tagihanMap = getHeaderMap(tagihanSheet);

    const headerCol = {
      idTransaksi: resolveColumn(headerMap, ["id_transaksi"], "id_transaksi"),

      noKuitansi: resolveColumn(
        headerMap,
        ["no_kuitansi", "nomor_kuitansi"],
        "no_kuitansi",
      ),

      tanggal: resolveColumn(
        headerMap,
        ["tanggal_transaksi", "tanggal", "tanggal_tx"],
        "tanggal transaksi",
      ),

      idSiswa: resolveColumn(headerMap, ["id_siswa"], "id_siswa"),

      totalBayar: resolveColumn(
        headerMap,
        ["total_bayar", "nominal_total", "total"],
        "total_bayar",
      ),

      metode: resolveColumn(
        headerMap,
        ["metode_pembayaran", "metode_bayar", "metode"],
        "metode_pembayaran",
      ),

      kasir: resolveColumn(
        headerMap,
        ["nama_kasir", "kasir", "petugas"],
        "nama_kasir",
      ),

      catatan: resolveColumn(headerMap, ["catatan", "keterangan"], "catatan"),

      status: resolveColumn(headerMap, ["status"], "status transaksi"),
    };

    const detailCol = {
      idDetail: resolveColumn(detailMap, ["id_detail"], "id_detail"),

      idTransaksi: resolveColumn(
        detailMap,
        ["id_transaksi"],
        "id_transaksi detail",
      ),

      idTagihan: resolveColumn(detailMap, ["id_tagihan"], "id_tagihan detail"),

      idSiswa: resolveColumn(detailMap, ["id_siswa"], "id_siswa detail"),

      snapshot: resolveColumn(
        detailMap,
        ["nama_item_snapshot", "nama_item", "nama_tagihan"],
        "nama item snapshot",
      ),

      nominal: resolveColumn(
        detailMap,
        ["nominal_dibayar", "jumlah_bayar", "nominal"],
        "nominal_dibayar",
      ),

      diskon: resolveColumn(
        detailMap,
        ["diskon_tambahan", "diskon"],
        "diskon_tambahan",
      ),

      tanggal: resolveColumn(
        detailMap,
        ["tanggal_transaksi", "tanggal"],
        "tanggal transaksi detail",
      ),
    };

    const requiredTagihan = [
      "id_tagihan",
      "id_siswa",
      "id_pos",
      "nominal_akhir",
      "terbayar",
      "sisa_tunggakan",
      "status",
    ];

    requireColumns(Object.keys(tagihanMap), requiredTagihan, "Tagihan_Siswa");

    // VALIDASI SISWA
    const siswaList = getSheetDataAsObjects("Siswa");

    const siswa = siswaList.find(function (s) {
      return normalizeText(s.id_siswa) === idSiswa;
    });

    if (!siswa) {
      throw new Error("Siswa dengan ID " + idSiswa + " tidak ditemukan.");
    }

    const statusSiswa = normalizeText(
      siswa.status || siswa.status_siswa,
    ).toUpperCase();

    if (statusSiswa !== "AKTIF") {
      throw new Error(
        "Siswa " +
          normalizeText(siswa.nama_lengkap || siswa.nama_siswa) +
          " tidak berstatus AKTIF.",
      );
    }

    // BACA TAGIHAN
    const tagihanValues = tagihanSheet.getDataRange().getValues();

    if (tagihanValues.length < 2) {
      throw new Error("Sheet Tagihan_Siswa belum memiliki data tagihan.");
    }

    const tagihanHeaders = tagihanValues[0].map(function (h) {
      return normalizeText(h);
    });

    const col = {
      idTagihan: tagihanHeaders.indexOf("id_tagihan"),

      idSiswa: tagihanHeaders.indexOf("id_siswa"),

      idPos: tagihanHeaders.indexOf("id_pos"),

      nominalAkhir: tagihanHeaders.indexOf("nominal_akhir"),

      terbayar: tagihanHeaders.indexOf("terbayar"),

      sisa: tagihanHeaders.indexOf("sisa_tunggakan"),

      status: tagihanHeaders.indexOf("status"),
    };

    Object.keys(col).forEach(function (key) {
      if (col[key] === -1) {
        throw new Error("Kolom tagihan tidak ditemukan: " + key);
      }
    });

    const tagihanMapById = new Map();

    for (let r = 1; r < tagihanValues.length; r++) {
      const row = tagihanValues[r];

      const idTagihan = normalizeText(row[col.idTagihan]);

      if (idTagihan) {
        tagihanMapById.set(idTagihan, {
          rowIndex: r + 1,
          row: row,
        });
      }
    }

    // VALIDASI SEMUA ITEM
    const posList = getSheetDataAsObjects("Master_Pos");

    const updatePlan = [];

    let totalBayar = 0;

    items.forEach(function (item) {
      const found = tagihanMapById.get(item.id_tagihan);

      if (!found) {
        throw new Error("Tagihan " + item.id_tagihan + " tidak ditemukan.");
      }

      const row = found.row;

      const rowSiswa = normalizeText(row[col.idSiswa]);

      if (rowSiswa !== idSiswa) {
        throw new Error(
          "Tagihan " + item.id_tagihan + " bukan milik siswa yang dipilih.",
        );
      }

      const nominalAkhir = toNumber(row[col.nominalAkhir]);

      const currentTerbayar = toNumber(row[col.terbayar]);

      const currentSisa = calculateSisa(nominalAkhir, currentTerbayar);

      if (
        currentSisa <= 0 ||
        normalizeText(row[col.status]).toUpperCase() === "LUNAS"
      ) {
        throw new Error("Tagihan " + item.id_tagihan + " sudah LUNAS.");
      }

      if (item.nominal_dibayar > currentSisa) {
        throw new Error(
          "Pembayaran tagihan " +
            item.id_tagihan +
            " melebihi sisa tagihan. Sisa: " +
            formatRupiah(currentSisa),
        );
      }

      const pos = posList.find(function (p) {
        return normalizeText(p.id_pos) === normalizeText(row[col.idPos]);
      });

      const isPartial = item.nominal_dibayar < currentSisa;

      if (isPartial && pos && !toBoolean(pos.bisa_dicicil)) {
        throw new Error("Tagihan " + item.id_tagihan + " tidak dapat dicicil.");
      }

      const newTerbayar = currentTerbayar + item.nominal_dibayar;

      const newSisa = calculateSisa(nominalAkhir, newTerbayar);

      const newStatus = newSisa <= 0 ? "LUNAS" : "CICILAN";

      if (!item.nama_item_snapshot) {
        item.nama_item_snapshot = pos
          ? normalizeText(pos.nama_pos)
          : item.id_tagihan;
      }

      originalTagihan.push({
        rowIndex: found.rowIndex,

        terbayar: row[col.terbayar],

        sisa: row[col.sisa],

        status: row[col.status],
      });

      updatePlan.push({
        rowIndex: found.rowIndex,

        newTerbayar: newTerbayar,

        newSisa: newSisa,

        newStatus: newStatus,
      });

      totalBayar += item.nominal_dibayar;
    });

    if (totalBayar <= 0) {
      throw new Error("Total pembayaran harus lebih dari 0.");
    }

    // GENERATE IDENTITAS TRANSAKSI
    const now = new Date();

    const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

    const yearMonth = Utilities.formatDate(now, timezone, "yyyyMM");

    const tanggalTx = Utilities.formatDate(
      now,
      timezone,
      "yyyy-MM-dd HH:mm:ss",
    );

    const idTransaksi = generateNextTransactionId(headerSheet);

    const noKuitansi = generateNextReceiptNumber(headerSheet, yearMonth);

    const detailIds = generateNextDetailIds(detailSheet, items.length);

    // TULIS HEADER
    const headerLastColumn = headerSheet.getLastColumn();

    const headerRowValues = new Array(headerLastColumn).fill("");

    headerRowValues[headerCol.idTransaksi] = idTransaksi;

    headerRowValues[headerCol.noKuitansi] = noKuitansi;

    headerRowValues[headerCol.tanggal] = tanggalTx;

    headerRowValues[headerCol.idSiswa] = idSiswa;

    headerRowValues[headerCol.totalBayar] = totalBayar;

    headerRowValues[headerCol.metode] =
      normalizeText(payload.metode_pembayaran) || "TUNAI";

    headerRowValues[headerCol.kasir] =
      normalizeText(payload.nama_kasir) || "Admin Kasir";

    headerRowValues[headerCol.catatan] = normalizeText(payload.catatan);

    headerRowValues[headerCol.status] = "SELESAI";

    headerRow = headerSheet.getLastRow() + 1;

    headerSheet
      .getRange(headerRow, 1, 1, headerLastColumn)
      .setValues([headerRowValues]);

    // TULIS DETAIL
    const detailLastColumn = detailSheet.getLastColumn();

    const detailRows = items.map(function (item, index) {
      const row = new Array(detailLastColumn).fill("");

      row[detailCol.idDetail] = detailIds[index];

      row[detailCol.idTransaksi] = idTransaksi;

      row[detailCol.idTagihan] = item.id_tagihan;

      row[detailCol.idSiswa] = idSiswa;

      row[detailCol.snapshot] = item.nama_item_snapshot;

      row[detailCol.nominal] = item.nominal_dibayar;

      row[detailCol.diskon] = item.diskon_tambahan;

      row[detailCol.tanggal] = tanggalTx;

      return row;
    });

    detailStartRow = detailSheet.getLastRow() + 1;

    detailRowCount = detailRows.length;

    detailSheet
      .getRange(detailStartRow, 1, detailRows.length, detailLastColumn)
      .setValues(detailRows);

    // UPDATE TAGIHAN
    updatePlan.forEach(function (update) {
      tagihanSheet
        .getRange(update.rowIndex, col.terbayar + 1)
        .setValue(update.newTerbayar);

      tagihanSheet
        .getRange(update.rowIndex, col.sisa + 1)
        .setValue(update.newSisa);

      tagihanSheet
        .getRange(update.rowIndex, col.status + 1)
        .setValue(update.newStatus);
    });

    SpreadsheetApp.flush();

    // RETURN
    return {
      success: true,

      message: "Transaksi berhasil disimpan!",

      id_transaksi: idTransaksi,

      no_kuitansi: noKuitansi,

      total_bayar: totalBayar,

      tanggal: tanggalTx,

      jumlah_item: items.length,
    };
  } catch (err) {
    // ROLLBACK
    try {
      const tagihanSheet = getSheet("Tagihan_Siswa");

      const tagihanMap = getHeaderMap(tagihanSheet);

      originalTagihan.forEach(function (original) {
        if (tagihanMap.terbayar !== undefined) {
          tagihanSheet
            .getRange(original.rowIndex, tagihanMap.terbayar + 1)
            .setValue(original.terbayar);
        }

        if (tagihanMap.sisa_tunggakan !== undefined) {
          tagihanSheet
            .getRange(original.rowIndex, tagihanMap.sisa_tunggakan + 1)
            .setValue(original.sisa);
        }

        if (tagihanMap.status !== undefined) {
          tagihanSheet
            .getRange(original.rowIndex, tagihanMap.status + 1)
            .setValue(original.status);
        }
      });

      if (detailStartRow && detailRowCount > 0) {
        const detailSheet = getSheet("Transaksi_Detail");

        if (detailStartRow <= detailSheet.getLastRow()) {
          detailSheet.deleteRows(
            detailStartRow,
            Math.min(
              detailRowCount,
              detailSheet.getLastRow() - detailStartRow + 1,
            ),
          );
        }
      }

      if (headerRow) {
        const headerSheet = getSheet("Transaksi_Header");

        if (headerRow <= headerSheet.getLastRow()) {
          headerSheet.deleteRow(headerRow);
        }
      }
    } catch (rollbackError) {
      // Rollback best effort.
    }

    return {
      success: false,

      message: "Gagal memproses transaksi: " + err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 3. GENERATOR ID TRANSAKSI
// ============================================================

function generateNextTransactionId(headerSheet) {
  if (!headerSheet) {
    throw new Error("Sheet Transaksi_Header tidak ditemukan.");
  }

  const values = headerSheet.getDataRange().getValues();

  const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

  const tanggal = Utilities.formatDate(new Date(), timezone, "yyyyMMdd");

  const prefix = "TRX-" + tanggal + "-";

  if (values.length < 2) {
    return prefix + "0001";
  }

  const headers = values[0].map(function (header) {
    return normalizeText(header).toLowerCase();
  });

  const idCol = headers.indexOf("id_transaksi");

  if (idCol === -1) {
    throw new Error("Kolom id_transaksi tidak ditemukan di Transaksi_Header.");
  }

  let maxNumber = 0;

  for (let i = 1; i < values.length; i++) {
    const id = normalizeText(values[i][idCol]);

    if (!id) {
      continue;
    }

    const match = id.match(new RegExp("^TRX-" + tanggal + "-(\\d+)$", "i"));

    if (match) {
      const nomor = Number(match[1]);

      if (Number.isFinite(nomor) && nomor > maxNumber) {
        maxNumber = nomor;
      }
    }
  }

  return prefix + String(maxNumber + 1).padStart(4, "0");
}

// ============================================================
// 4. GENERATOR NOMOR KUITANSI
// ============================================================

function generateNextReceiptNumber(headerSheet, yearMonth) {
  if (!headerSheet) {
    throw new Error("Sheet Transaksi_Header tidak ditemukan.");
  }

  const values = headerSheet.getDataRange().getValues();

  if (values.length < 1) {
    throw new Error("Struktur Transaksi_Header tidak valid.");
  }

  const headers = values[0].map(function (header) {
    return normalizeText(header).toLowerCase();
  });

  const kuitansiCol = headers.indexOf("no_kuitansi");

  if (kuitansiCol === -1) {
    throw new Error("Kolom no_kuitansi tidak ditemukan di Transaksi_Header.");
  }

  const prefix = "KW-" + yearMonth + "-";

  let max = 0;

  for (let i = 1; i < values.length; i++) {
    const value = normalizeText(values[i][kuitansiCol]);

    if (!value.startsWith(prefix)) {
      continue;
    }

    const match = value.match(/(\d+)$/);

    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }

  return prefix + String(max + 1).padStart(4, "0");
}

// ============================================================
// 5. GENERATOR ID DETAIL
// ============================================================

function generateNextDetailIds(detailSheet, count) {
  if (!detailSheet) {
    throw new Error("Sheet Transaksi_Detail tidak ditemukan.");
  }

  const values = detailSheet.getDataRange().getValues();

  if (values.length === 0) {
    throw new Error("Struktur Transaksi_Detail tidak valid.");
  }

  const headers = values[0].map(function (header) {
    return normalizeText(header).toLowerCase();
  });

  const idCol = headers.indexOf("id_detail");

  if (idCol === -1) {
    throw new Error("Kolom id_detail tidak ditemukan di Transaksi_Detail.");
  }

  let max = 0;

  for (let i = 1; i < values.length; i++) {
    const value = normalizeText(values[i][idCol]);

    if (!value.startsWith("DET-")) {
      continue;
    }

    const match = value.match(/(\d+)$/);

    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }

  const ids = [];

  for (let i = 1; i <= count; i++) {
    ids.push("DET-" + String(max + i).padStart(6, "0"));
  }

  return ids;
}