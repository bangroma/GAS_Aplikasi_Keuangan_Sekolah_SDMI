/**
 * SISTEM KEUANGAN SEKOLAH
 * GOOGLE APPS SCRIPT BACKEND
 *
 * Versi Final:
 * - Setup database otomatis
 * - Struktur sheet otomatis
 * - Validasi transaksi
 * - Batch write
 * - Script Lock
 * - Rollback
 * - ID transaksi terpisah dari nomor kuitansi
 * - Generator ID aman berdasarkan header
 * - Cicilan pembayaran
 * - Generate tagihan SPP
 * - Rekap tunggakan
 * - Export laporan
 * - Tanpa log percobaan/debug
 */

// ============================================================
// 1. WEB APP
// ============================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Sistem Keuangan MI Muhammadiyah Kertonatan")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ============================================================
// 2. KONFIGURASI DATABASE
// ============================================================

const DATABASE_SCHEMA = {
  Siswa: [
    "id_siswa",
    "nisn",
    "nis",
    "nama_lengkap",
    "jenis_kelamin",
    "kelas",
    "tahun_pelajaran",
    "status",
    "id_potongan_default",
  ],

  Master_Pos: [
    "id_pos",
    "kode_pos",
    "nama_pos",
    "tipe_pos",
    "tarif_default",
    "bisa_dicicil",
    "aktif",
  ],

  Master_Potongan: [
    "id_potongan",
    "kode_potongan",
    "nama_potongan",
    "tipe_nilai",
    "nilai_potongan",
    "aktif",
  ],

  Tagihan_Siswa: [
    "id_tagihan",
    "id_siswa",
    "id_pos",
    "id_potongan",
    "tahun_pelajaran",
    "periode",
    "nama_item_snapshot",
    "nominal_asal",
    "terbayar",
    "sisa_tunggakan",
    "bisa_dicicil",
    "tanggal_tagihan",
    "keterangan",
    "nominal_potongan",
    "nominal_akhir",
    "status",
  ],

  Transaksi_Header: [
    "id_transaksi",
    "no_kuitansi",
    "tanggal_transaksi",
    "id_siswa",
    "total_bayar",
    "metode_pembayaran",
    "nama_kasir",
    "catatan",
    "status",
  ],

  Transaksi_Detail: [
    "id_detail",
    "id_transaksi",
    "id_tagihan",
    "id_siswa",
    "nama_item_snapshot",
    "nominal_dibayar",
    "diskon_tambahan",
    "tanggal_transaksi",
  ],

  Pos_Pembayaran: ["id_pos", "kode_pos", "nama_pos", "aktif"],
};

// ============================================================
// 3. HELPER DASAR
// ============================================================

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "Sheet '" +
        sheetName +
        "' tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.",
    );
  }

  return sheet;
}

function getSheetDataAsObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function (h) {
    return normalizeText(h);
  });

  return values.slice(1).map(function (row) {
    const obj = {};

    headers.forEach(function (header, index) {
      if (header) {
        obj[header] = row[index];
      }
    });

    return obj;
  });
}

function normalizeText(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function toNumber(value, defaultValue) {
  const n = Number(value);
  return Number.isFinite(n) ? n : defaultValue || 0;
}

function toBoolean(value) {
  const text = normalizeText(value).toUpperCase();

  return (
    value === true ||
    ["TRUE", "1", "YA", "Y", "YES", "AKTIF"].indexOf(text) !== -1
  );
}

function resolveColumn(map, aliases, label) {
  for (const alias of aliases) {
    if (map[alias] !== undefined) {
      return map[alias];
    }
  }

  throw new Error(
    "Kolom " +
      label +
      " tidak ditemukan. Alias yang didukung: " +
      aliases.join(", "),
  );
}

function requireColumns(headers, requiredColumns, sheetName) {
  const missing = requiredColumns.filter(function (col) {
    return headers.indexOf(col) === -1;
  });

  if (missing.length > 0) {
    throw new Error(
      "Kolom wajib tidak ditemukan pada sheet '" +
        sheetName +
        "': " +
        missing.join(", "),
    );
  }
}

function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return {};
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  const map = {};

  headers.forEach(function (header, index) {
    const name = normalizeText(header);

    if (name) {
      map[name] = index;
    }
  });

  return map;
}

function calculateSisa(nominalAkhir, terbayar) {
  return Math.max(0, toNumber(nominalAkhir) - toNumber(terbayar));
}

function formatRupiah(value) {
  return "Rp " + toNumber(value).toLocaleString("id-ID");
}

// ============================================================
// 4. SETUP DATABASE
// ============================================================

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hasil = {
    success: true,
    dibuat: [],
    diperbaiki: [],
    data_master: [],
  };

  Object.keys(DATABASE_SCHEMA).forEach(function (sheetName) {
    const requiredColumns = DATABASE_SCHEMA[sheetName];

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);

      sheet
        .getRange(1, 1, 1, requiredColumns.length)
        .setValues([requiredColumns]);

      hasil.dibuat.push(sheetName);
    } else {
      const lastColumn = sheet.getLastColumn();

      let headers = [];

      if (lastColumn > 0) {
        headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function (h) {
            return normalizeText(h);
          });
      }

      if (
        headers.length === 0 ||
        headers.every(function (h) {
          return !h;
        })
      ) {
        sheet
          .getRange(1, 1, 1, requiredColumns.length)
          .setValues([requiredColumns]);

        hasil.diperbaiki.push(sheetName);
      } else {
        const existing = {};

        headers.forEach(function (header) {
          if (header) {
            existing[header] = true;
          }
        });

        const missing = requiredColumns.filter(function (column) {
          return !existing[column];
        });

        if (missing.length > 0) {
          const startColumn = sheet.getLastColumn() + 1;

          sheet
            .getRange(1, startColumn, 1, missing.length)
            .setValues([missing]);

          hasil.diperbaiki.push(sheetName + ": +" + missing.join(", "));
        }
      }
    }

    formatDatabaseSheet(sheet);
  });

  seedMasterDatabase();

  hasil.data_master.push("Master_Pos");
  hasil.data_master.push("Master_Potongan");
  hasil.data_master.push("Pos_Pembayaran");

  return hasil;
}

// ============================================================
// 5. FORMAT SHEET DATABASE
// ============================================================

function formatDatabaseSheet(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return;
  }

  sheet
    .getRange(1, 1, 1, lastColumn)
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  sheet.setFrozenRows(1);

  try {
    sheet.autoResizeColumns(1, lastColumn);
  } catch (e) {
    // Tidak perlu log.
  }
}

// ============================================================
// 6. SEED DATA MASTER
// ============================================================

function seedMasterDatabase() {
  seedMasterPos();
  seedMasterPotongan();
  seedPosPembayaran();
}

function seedMasterPos() {
  const sheet = getSheet("Master_Pos");

  const data = getSheetDataAsObjects("Master_Pos");

  const defaults = [
    ["POS-SPP", "SPP", "SPP Bulanan", "BULANAN", 150000, true, true],
    ["POS-SRG", "SRG", "Seragam Siswa", "SEKALI", 750000, true, true],
    ["POS-KGT", "KGT", "Kegiatan Tahunan", "KEGIATAN", 1000000, true, true],
  ];

  const existing = new Set(
    data.map(function (row) {
      return normalizeText(row.id_pos);
    }),
  );

  const rows = defaults.filter(function (row) {
    return !existing.has(row[0]);
  });

  if (rows.length === 0) {
    return;
  }

  const map = getHeaderMap(sheet);
  const lastColumn = sheet.getLastColumn();

  const output = rows.map(function (dataRow) {
    const row = new Array(lastColumn).fill("");

    row[map.id_pos] = dataRow[0];
    row[map.kode_pos] = dataRow[1];
    row[map.nama_pos] = dataRow[2];
    row[map.tipe_pos] = dataRow[3];
    row[map.tarif_default] = dataRow[4];
    row[map.bisa_dicicil] = dataRow[5];
    row[map.aktif] = dataRow[6];

    return row;
  });

  sheet
    .getRange(sheet.getLastRow() + 1, 1, output.length, lastColumn)
    .setValues(output);
}

function seedMasterPotongan() {
  const sheet = getSheet("Master_Potongan");

  const data = getSheetDataAsObjects("Master_Potongan");

  const defaults = [
    ["POT-UMUM", "UMUM", "Potongan Umum", "NOMINAL", 0, true],
    ["POT-SIBLING", "SIBLING", "Potongan Saudara", "PERSEN", 10, true],
  ];

  const existing = new Set(
    data.map(function (row) {
      return normalizeText(row.id_potongan);
    }),
  );

  const rows = defaults.filter(function (row) {
    return !existing.has(row[0]);
  });

  if (rows.length === 0) {
    return;
  }

  const map = getHeaderMap(sheet);
  const lastColumn = sheet.getLastColumn();

  const output = rows.map(function (dataRow) {
    const row = new Array(lastColumn).fill("");

    row[map.id_potongan] = dataRow[0];
    row[map.kode_potongan] = dataRow[1];
    row[map.nama_potongan] = dataRow[2];
    row[map.tipe_nilai] = dataRow[3];
    row[map.nilai_potongan] = dataRow[4];
    row[map.aktif] = dataRow[5];

    return row;
  });

  sheet
    .getRange(sheet.getLastRow() + 1, 1, output.length, lastColumn)
    .setValues(output);
}

function seedPosPembayaran() {
  const sheet = getSheet("Pos_Pembayaran");

  const data = getSheetDataAsObjects("Pos_Pembayaran");

  const master = getSheetDataAsObjects("Master_Pos");

  const existing = new Set(
    data.map(function (row) {
      return normalizeText(row.id_pos);
    }),
  );

  const rows = master
    .filter(function (pos) {
      return toBoolean(pos.aktif);
    })
    .filter(function (pos) {
      return !existing.has(normalizeText(pos.id_pos));
    });

  if (rows.length === 0) {
    return;
  }

  const map = getHeaderMap(sheet);
  const lastColumn = sheet.getLastColumn();

  const output = rows.map(function (pos) {
    const row = new Array(lastColumn).fill("");

    row[map.id_pos] = pos.id_pos;
    row[map.kode_pos] = pos.kode_pos;
    row[map.nama_pos] = pos.nama_pos;
    row[map.aktif] = true;

    return row;
  });

  sheet
    .getRange(sheet.getLastRow() + 1, 1, output.length, lastColumn)
    .setValues(output);
}

// ============================================================
// 7. CARI SISWA
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
// 8. AMBIL TAGIHAN SISWA
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

      // ---------- PERIODE BERSIH ----------
      let periodeBersih = "";
      const rawPeriode = t.periode;

      if (rawPeriode instanceof Date && !isNaN(rawPeriode.getTime())) {
        // Kalau terlanjur tersimpan sebagai Date
        periodeBersih = Utilities.formatDate(rawPeriode, timezone, "yyyy-MM");
      } else {
        periodeBersih = normalizeText(rawPeriode);
      }

      // ---------- NAMA ITEM BERSIH ----------
      let namaPosBersih = "";
      if (pos) {
        namaPosBersih = normalizeText(pos.nama_pos);
      } else {
        // Ambil dari snapshot, tapi bersihkan kalau ada string Date di dalamnya
        let snapshot = normalizeText(t.nama_item_snapshot);
        // Hapus bagian (Wed Aug ... GMT...) kalau ada
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

    // Kalau periode berisi Date object
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      const bersih = Utilities.formatDate(raw, timezone, "yyyy-MM");
      sheet.getRange(i + 1, colPeriode + 1).setValue(bersih);
      updated++;
    }

    // Bersihkan juga nama_item_snapshot yang mengandung string GMT
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
// 9. SIMPAN TRANSAKSI
// ============================================================

function simpanTransaksi(payload) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  let headerRow = null;
  let detailStartRow = null;
  let detailRowCount = 0;

  const originalTagihan = [];

  try {
    // --------------------------------------------------------
    // A. VALIDASI PAYLOAD
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // B. SHEET
    // --------------------------------------------------------

    const headerSheet = getSheet("Transaksi_Header");

    const detailSheet = getSheet("Transaksi_Detail");

    const tagihanSheet = getSheet("Tagihan_Siswa");

    const siswaSheet = getSheet("Siswa");

    const posSheet = getSheet("Master_Pos");

    // --------------------------------------------------------
    // C. STRUKTUR HEADER
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // D. VALIDASI SISWA
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // E. BACA TAGIHAN
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // F. VALIDASI SEMUA ITEM
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // G. GENERATE IDENTITAS TRANSAKSI
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // H. TULIS HEADER
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // I. TULIS DETAIL
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // J. UPDATE TAGIHAN
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // K. RETURN
    // --------------------------------------------------------

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
    // --------------------------------------------------------
    // L. ROLLBACK
    // --------------------------------------------------------

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
// 10. GENERATOR ID TRANSAKSI
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
// 11. GENERATOR NOMOR KUITANSI
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
// 12. GENERATOR ID DETAIL
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

// ============================================================
// 13. GENERATE TAGIHAN SPP BULANAN
// ============================================================

function generateTagihanSPPBulanan() {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const sheetSiswa = getSheet("Siswa");

    const sheetTagihan = getSheet("Tagihan_Siswa");

    const posList = getSheetDataAsObjects("Master_Pos");

    const potonganList = getSheetDataAsObjects("Master_Potongan");

    const now = new Date();

    const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

    const periodeFormatted = Utilities.formatDate(now, timezone, "yyyy-MM");

    const tagihanPrefix =
      "TAG-SPP-" + Utilities.formatDate(now, timezone, "yyyyMM");

    const posSPP = posList.find(function (p) {
      const id = normalizeText(p.id_pos).toUpperCase();

      const kode = normalizeText(p.kode_pos).toUpperCase();

      const nama = normalizeText(p.nama_pos).toUpperCase();

      return id === "POS-SPP" || kode === "SPP" || nama.includes("SPP");
    });

    if (!posSPP) {
      throw new Error("Master POS SPP tidak ditemukan.");
    }

    const tarifDefaultSPP = toNumber(
      posSPP.tarif_default || posSPP.nominal_default,
    );

    if (tarifDefaultSPP <= 0) {
      throw new Error("Tarif default SPP harus lebih dari 0.");
    }

    const existingTagihan = getSheetDataAsObjects("Tagihan_Siswa");

    const existingMap = new Set(
      existingTagihan
        .filter(function (t) {
          return normalizeText(t.id_pos) === normalizeText(posSPP.id_pos);
        })
        .filter(function (t) {
          return normalizeText(t.periode) === periodeFormatted;
        })
        .map(function (t) {
          return normalizeText(t.id_siswa);
        }),
    );

    const siswaList = getSheetDataAsObjects("Siswa").filter(function (s) {
      const status = normalizeText(s.status || s.status_siswa).toUpperCase();

      return status === "AKTIF";
    });

    if (siswaList.length === 0) {
      return {
        success: false,
        message: "Tidak ada siswa aktif ditemukan.",
      };
    }

    const tagihanMap = getHeaderMap(sheetTagihan);

    const required = [
      "id_tagihan",
      "id_siswa",
      "id_pos",
      "periode",
      "nominal_asal",
      "nominal_potongan",
      "nominal_akhir",
      "terbayar",
      "sisa_tunggakan",
      "status",
    ];

    requireColumns(Object.keys(tagihanMap), required, "Tagihan_Siswa");

    const lastColumn = sheetTagihan.getLastColumn();

    const newRows = [];

    let skippedCount = 0;

    let maxSequence = 0;

    existingTagihan.forEach(function (t) {
      const id = normalizeText(t.id_tagihan);

      if (!id.startsWith(tagihanPrefix + "-")) {
        return;
      }

      const match = id.match(/(\d+)$/);

      if (match) {
        maxSequence = Math.max(maxSequence, Number(match[1]));
      }
    });

    siswaList.forEach(function (siswa) {
      const idSiswa = normalizeText(siswa.id_siswa);

      if (existingMap.has(idSiswa)) {
        skippedCount++;
        return;
      }

      maxSequence++;

      const idTagihan =
        tagihanPrefix + "-" + String(maxSequence).padStart(4, "0");

      let nominalPotongan = 0;

      const idPotongan = normalizeText(siswa.id_potongan_default);

      if (idPotongan && idPotongan !== "NONE") {
        const pData = potonganList.find(function (p) {
          return normalizeText(p.id_potongan) === idPotongan;
        });

        if (pData) {
          const tipe = normalizeText(pData.tipe_nilai).toUpperCase();

          if (tipe === "PERSEN") {
            nominalPotongan =
              (tarifDefaultSPP * toNumber(pData.nilai_potongan)) / 100;
          } else if (tipe === "NOMINAL") {
            nominalPotongan = toNumber(pData.nilai_potongan);
          }
        }
      }

      nominalPotongan = Math.min(Math.max(0, nominalPotongan), tarifDefaultSPP);

      const nominalAkhir = Math.max(0, tarifDefaultSPP - nominalPotongan);

      const row = new Array(lastColumn).fill("");

      row[tagihanMap.id_tagihan] = idTagihan;

      row[tagihanMap.id_siswa] = idSiswa;

      row[tagihanMap.id_pos] = posSPP.id_pos;

      if (tagihanMap.id_potongan !== undefined) {
        row[tagihanMap.id_potongan] = idPotongan;
      }

      if (tagihanMap.tahun_pelajaran !== undefined) {
        row[tagihanMap.tahun_pelajaran] = siswa.tahun_pelajaran || "";
      }

      row[tagihanMap.periode] = periodeFormatted;

      if (tagihanMap.nama_item_snapshot !== undefined) {
        row[tagihanMap.nama_item_snapshot] = posSPP.nama_pos;
      }

      row[tagihanMap.nominal_asal] = tarifDefaultSPP;

      row[tagihanMap.terbayar] = 0;

      row[tagihanMap.sisa_tunggakan] = nominalAkhir;

      if (tagihanMap.bisa_dicicil !== undefined) {
        row[tagihanMap.bisa_dicicil] = toBoolean(posSPP.bisa_dicicil);
      }

      if (tagihanMap.tanggal_tagihan !== undefined) {
        row[tagihanMap.tanggal_tagihan] = now;
      }

      if (tagihanMap.keterangan !== undefined) {
        row[tagihanMap.keterangan] = "Tagihan SPP " + periodeFormatted;
      }

      row[tagihanMap.nominal_potongan] = nominalPotongan;

      row[tagihanMap.nominal_akhir] = nominalAkhir;

      row[tagihanMap.status] = nominalAkhir <= 0 ? "LUNAS" : "BELUM_BAYAR";

      newRows.push(row);
    });

    if (newRows.length > 0) {
      const startRow = sheetTagihan.getLastRow() + 1;

      sheetTagihan
        .getRange(startRow, 1, newRows.length, lastColumn)
        .setValues(newRows);
    }

    return {
      success: true,

      message:
        "Selesai. Berhasil generate: " +
        newRows.length +
        " tagihan. Dilewati: " +
        skippedCount +
        " siswa.",

      generated: newRows.length,

      skipped: skippedCount,
    };
  } catch (err) {
    return {
      success: false,

      message: "Gagal generate tagihan SPP: " + err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 14. DAFTAR POS PEMBAYARAN
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

// ============================================================
// 15. REKAP TUNGGAKAN PER KELAS
// ============================================================

function getRekapTunggakanPerKelas(idPos) {
  const siswaList = getSheetDataAsObjects("Siswa").filter(function (s) {
    const status = normalizeText(s.status || s.status_siswa).toUpperCase();

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
    const sisa = calculateSisa(t.nominal_akhir, t.terbayar);

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
// 16. EKSPOR REKAP TUNGGAKAN
// ============================================================

function exportLaporanTunggakanToSheet(idPos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const now = new Date();

  const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

  const dateStr = Utilities.formatDate(now, timezone, "yyyyMMdd_HHmmss");

  const sheetName = "Rekap_Tunggakan_" + dateStr;

  const sheet = ss.insertSheet(sheetName);

  const namaPosText = getNamaPos(idPos);

  const rekapKelas = getRekapTunggakanPerKelas(idPos);

  sheet
    .getRange("A1:D1")
    .merge()
    .setValue("REKAPITULASI TUNGGAKAN PEMBAYARAN PER KELAS")
    .setFontWeight("bold")
    .setFontSize(13);

  sheet
    .getRange("A2:D2")
    .merge()
    .setValue("Pos Pembayaran: " + namaPosText)
    .setFontWeight("bold")
    .setFontSize(10);

  sheet
    .getRange("A3:D3")
    .merge()
    .setValue(
      "Tanggal Cetak: " +
        Utilities.formatDate(now, timezone, "dd MMMM yyyy HH:mm:ss"),
    )
    .setFontSize(9)
    .setFontItalic(true);

  sheet
    .getRange(5, 1, 1, 4)
    .setValues([
      ["Kelas", "Total Siswa", "Siswa Menunggak", "Total Nominal Tunggakan"],
    ])
    .setFontWeight("bold")
    .setBackground("#e2e8f0");

  const rowData = rekapKelas.map(function (r) {
    return [r.kelas, r.total_siswa, r.siswa_nunggak, r.total_tunggakan];
  });

  let grandTotal = 0;

  let grandSiswaNunggak = 0;

  rekapKelas.forEach(function (r) {
    grandTotal += r.total_tunggakan;

    grandSiswaNunggak += r.siswa_nunggak;
  });

  if (rowData.length > 0) {
    sheet.getRange(6, 1, rowData.length, 4).setValues(rowData);

    sheet.getRange(6, 4, rowData.length, 1).setNumberFormat("Rp #,##0");
  }

  const totalRow = 6 + rowData.length;

  sheet
    .getRange(totalRow, 1, 1, 2)
    .merge()
    .setValue("TOTAL KESELURUHAN")
    .setFontWeight("bold");

  sheet.getRange(totalRow, 3).setValue(grandSiswaNunggak).setFontWeight("bold");

  sheet
    .getRange(totalRow, 4)
    .setValue(grandTotal)
    .setFontWeight("bold")
    .setNumberFormat("Rp #,##0");

  sheet.getRange(totalRow, 1, 1, 4).setBackground("#cbd5e1");

  sheet.autoResizeColumns(1, 4);

  return {
    success: true,

    sheetName: sheetName,

    message:
      "Laporan filter '" +
      namaPosText +
      "' berhasil diekspor ke tab sheet '" +
      sheetName +
      "'",
  };
}

// ============================================================
// 17. NAMA POS
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
// 18. RIWAYAT TRANSAKSI
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
// 19. DETAIL TRANSAKSI
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
/**
 * DETAIL SISWA MENUNGGAK PER KELAS
 */
function getDetailTunggakanPerKelas(kelas, idPos) {
  kelas = normalizeText(kelas);
  idPos = normalizeText(idPos);

  if (!kelas) throw new Error("Kelas wajib diisi.");

  const siswaList = getSheetDataAsObjects("Siswa");
  let tagihanList = getSheetDataAsObjects("Tagihan_Siswa");

  // Filter siswa aktif
  const siswaAktif = siswaList.filter(function (s) {
    return normalizeText(s.status || s.status_siswa).toUpperCase() === "AKTIF";
  });

  const siswaKelas = siswaAktif.filter(function (s) {
    return normalizeText(s.kelas) === kelas;
  });
  const siswaKelasIds = {};

  siswaKelas.forEach(function (siswa) {
    siswaKelasIds[normalizeText(siswa.id_siswa)] = true;
  });

  // Filter tagihan sesuai pos jika dipilih
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
      let totalSisa = 0;

      const detailTagihan = tagihanSiswa.map(function (t) {
        const nominalAkhir = toNumber(t.nominal_akhir);
        const terbayar = toNumber(t.terbayar);
        const sisa = Math.max(
          0,
          toNumber(t.sisa_tunggakan) || nominalAkhir - terbayar,
        );

        totalTagihan += nominalAkhir;
        totalDibayar += terbayar;
        totalSisa += sisa;

        const progress =
          nominalAkhir > 0
            ? Math.min(100, Math.round((terbayar / nominalAkhir) * 100))
            : 0;

        return {
          id_tagihan: normalizeText(t.id_tagihan),
          id_pos: normalizeText(t.id_pos),
          nama_item: normalizeText(
            t.nama_item_snapshot || t.nama_item || "Tagihan",
          ),
          periode: normalizeText(t.periode),
          nominal_akhir: nominalAkhir,
          terbayar: terbayar,
          sisa_tunggakan: sisa,
          progress: progress,
          status: normalizeText(t.status) || "BELUM_BAYAR",
        };
      });

      if (totalSisa <= 0) return;

      const progressTotal =
        totalTagihan > 0
          ? Math.min(100, Math.round((totalDibayar / totalTagihan) * 100))
          : 0;

      hasil.push({
        id_siswa: idSiswa,
        nisn: normalizeText(siswa.nisn),
        nis: normalizeText(siswa.nis),
        nama_lengkap: normalizeText(siswa.nama_lengkap || siswa.nama_siswa),
        kelas: normalizeText(siswa.kelas),
        total_tagihan: totalTagihan,
        total_dibayar: totalDibayar,
        total_sisa: totalSisa,
        progress: progressTotal,
        status: totalDibayar > 0 ? "CICILAN" : "BELUM BAYAR",
        tagihan: detailTagihan,
      });
    });

  hasil.sort(function (a, b) {
    return a.nama_lengkap.localeCompare(b.nama_lengkap);
  });

  return hasil;
}

/**
 * EXPORT DETAIL TUNGGAKAN SISWA KE GOOGLE SHEET
 * Hasil sheet dapat diunduh sebagai Excel (.xlsx)
 */
function exportDetailTunggakanToSheet(kelas, idPos) {
  kelas = normalizeText(kelas);
  idPos = normalizeText(idPos) || "ALL";

  if (!kelas) {
    throw new Error("Kelas wajib diisi.");
  }

  const data = getDetailTunggakanPerKelas(kelas, idPos);

  if (!data || data.length === 0) {
    throw new Error("Tidak ada data siswa menunggak untuk diekspor.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss",
  );

  const safeKelas = kelas.replace(/[\\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");

  const baseName = "Tunggakan_" + safeKelas + "_" + timestamp;

  let sheetName = baseName.substring(0, 100);
  let counter = 1;

  while (ss.getSheetByName(sheetName)) {
    sheetName = (baseName + "_" + counter).substring(0, 100);
    counter++;
  }

  const sheet = ss.insertSheet(sheetName);

  // ================================
  // HEADER LAPORAN
  // ================================
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").setValue("LAPORAN DETAIL TUNGGAKAN SISWA");

  sheet.getRange("A2:H2").merge();
  sheet.getRange("A2").setValue("Kelas: " + kelas);

  sheet.getRange("A3:H3").merge();
  sheet
    .getRange("A3")
    .setValue(
      "Tanggal Cetak: " +
        Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm",
        ),
    );

  // ================================
  // HEADER TABEL
  // ================================
  const headers = [
    "No",
    "NISN",
    "NIS",
    "Nama Siswa",
    "Total Tagihan",
    "Sudah Dibayar",
    "Sisa Tunggakan",
    "Progress",
  ];

  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

  // ================================
  // DATA SISWA
  // ================================
  const rows = [];

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

  sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);

  // ================================
  // DETAIL TAGIHAN
  // ================================
  let detailStartRow = 6 + rows.length + 3;

  sheet.getRange(detailStartRow, 1, 1, 7).merge();
  sheet.getRange(detailStartRow, 1).setValue("DETAIL TAGIHAN SISWA");

  detailStartRow++;

  const detailHeaders = [
    "No",
    "Nama Siswa",
    "ID Tagihan",
    "Item Tagihan",
    "Total Tagihan",
    "Sudah Dibayar",
    "Sisa Tunggakan",
  ];

  sheet
    .getRange(detailStartRow, 1, 1, detailHeaders.length)
    .setValues([detailHeaders]);

  detailStartRow++;

  const detailRows = [];
  let detailNo = 1;

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
    sheet
      .getRange(detailStartRow, 1, detailRows.length, detailHeaders.length)
      .setValues(detailRows);

    sheet
      .getRange(detailStartRow, 5, detailRows.length, 3)
      .setNumberFormat('"Rp" #,##0');
  }

  // ================================
  // FORMAT
  // ================================
  sheet
    .getRange("A1:H1")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center");

  sheet.getRange("A2:H3").setHorizontalAlignment("center");

  sheet
    .getRange(5, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d1fae5")
    .setHorizontalAlignment("center");

  sheet.getRange(6, 5, rows.length, 3).setNumberFormat('"Rp" #,##0');

  sheet.getRange(6, 8, rows.length, 1).setNumberFormat("0%");

  sheet
    .getRange(detailStartRow - 1, 1, 1, detailHeaders.length)
    .setFontWeight("bold")
    .setBackground("#e2e8f0")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(5);

  // Border
  sheet
    .getRange(5, 1, rows.length + 1, headers.length)
    .setBorder(true, true, true, true, true, true);

  if (detailRows.length > 0) {
    sheet
      .getRange(
        detailStartRow - 1,
        1,
        detailRows.length + 1,
        detailHeaders.length,
      )
      .setBorder(true, true, true, true, true, true);
  }

  sheet.autoResizeColumns(1, 8);

  // Lebar minimum agar nyaman dibaca
  sheet.setColumnWidth(4, 250);

  // ================================
  // RETURN
  // ================================
  return {
    success: true,
    sheetName: sheetName,
    spreadsheetUrl: ss.getUrl(),
    message: "Laporan detail tunggakan kelas " + kelas + " berhasil dibuat.",
  };
}

// ============================================================
// 20. GET LIST KELAS
// ============================================================

function getListKelas() {
  const siswaList = getSheetDataAsObjects("Siswa").filter(function (s) {
    const status = normalizeText(s.status || s.status_siswa).toUpperCase();
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
// 21. GET SISWA BY KELAS
// ============================================================

function getSiswaByKelas(kelas) {
  kelas = normalizeText(kelas);

  if (!kelas) {
    return [];
  }

  const siswaList = getSheetDataAsObjects("Siswa");
  const potonganList = getSheetDataAsObjects("Master_Potongan");

  return siswaList
    .filter(function (s) {
      const status = normalizeText(s.status || s.status_siswa).toUpperCase();
      return status === "AKTIF" && normalizeText(s.kelas) === kelas;
    })
    .map(function (s) {
      const potongan = potonganList.find(function (p) {
        return (
          normalizeText(p.id_potongan) === normalizeText(s.id_potongan_default)
        );
      });

      return {
        id_siswa: normalizeText(s.id_siswa),
        nisn: normalizeText(s.nisn),
        nis: normalizeText(s.nis),
        nama_lengkap: normalizeText(s.nama_lengkap || s.nama_siswa),
        kelas: normalizeText(s.kelas),
        tahun_pelajaran: normalizeText(s.tahun_pelajaran),
        id_potongan_default: normalizeText(s.id_potongan_default),
        nama_potongan: potongan
          ? normalizeText(potongan.nama_potongan)
          : "Tidak Ada",
      };
    })
    .sort(function (a, b) {
      return a.nama_lengkap.localeCompare(b.nama_lengkap);
    });
}
// ============================================================
// 22. GENERATE TAGIHAN MASSAL (FLEXIBLE)
// ============================================================

function generateTagihanMassal(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!payload || typeof payload !== "object") {
      throw new Error("Payload tidak valid.");
    }

    const idPos = normalizeText(payload.id_pos);
    const periode = normalizeText(payload.periode);
    const nominalInput = toNumber(payload.nominal);
    const idPotonganInput = normalizeText(payload.id_potongan);
    const bisaDicicil =
      payload.bisa_dicicil === true || toBoolean(payload.bisa_dicicil);
    const catatan = normalizeText(payload.catatan);
    const idSiswaList = Array.isArray(payload.id_siswa_list)
      ? payload.id_siswa_list.map(normalizeText).filter(Boolean)
      : [];

    if (!idPos) throw new Error("Pos pembayaran wajib dipilih.");
    if (!periode) throw new Error("Periode wajib diisi.");
    if (idSiswaList.length === 0) throw new Error("Pilih minimal satu siswa.");

    // Ambil data master
    const posList = getSheetDataAsObjects("Master_Pos");
    const potonganList = getSheetDataAsObjects("Master_Potongan");
    const siswaList = getSheetDataAsObjects("Siswa");
    const existingTagihan = getSheetDataAsObjects("Tagihan_Siswa");

    const pos = posList.find(function (p) {
      return normalizeText(p.id_pos) === idPos;
    });

    if (!pos) {
      throw new Error("Pos pembayaran tidak ditemukan.");
    }

    if (idPotonganInput && idPotonganInput !== "NONE") {
      const potonganInput = potonganList.find(function (p) {
        return normalizeText(p.id_potongan) === idPotonganInput;
      });
      if (!potonganInput) {
        throw new Error("Potongan yang dipilih tidak ditemukan.");
      }
    }

    // Nominal: prioritaskan input user, kalau kosong pakai tarif_default
    let nominalAsal =
      nominalInput > 0 ? nominalInput : toNumber(pos.tarif_default);

    if (nominalAsal <= 0) {
      throw new Error("Nominal tagihan harus lebih dari 0.");
    }

    const sheetTagihan = getSheet("Tagihan_Siswa");
    const tagihanMap = getHeaderMap(sheetTagihan);
    const lastColumn = sheetTagihan.getLastColumn();

    // Pastikan kolom wajib ada
    const required = [
      "id_tagihan",
      "id_siswa",
      "id_pos",
      "periode",
      "nominal_asal",
      "nominal_potongan",
      "nominal_akhir",
      "terbayar",
      "sisa_tunggakan",
      "status",
    ];
    requireColumns(Object.keys(tagihanMap), required, "Tagihan_Siswa");

    // Existing map: siswa + pos + periode (untuk skip duplikat)
    const existingSet = new Set(
      existingTagihan
        .filter(function (t) {
          return (
            normalizeText(t.id_pos) === idPos &&
            normalizeText(t.periode) === periode
          );
        })
        .map(function (t) {
          return normalizeText(t.id_siswa);
        }),
    );

    // Hitung max sequence untuk ID tagihan
    const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";
    const now = new Date();
    const prefix = "TAG-" + Utilities.formatDate(now, timezone, "yyyyMM") + "-";

    let maxSequence = 0;
    existingTagihan.forEach(function (t) {
      const id = normalizeText(t.id_tagihan);
      if (id.startsWith(prefix)) {
        const match = id.match(/(\d+)$/);
        if (match) {
          maxSequence = Math.max(maxSequence, Number(match[1]));
        }
      }
    });

    const newRows = [];
    let skippedCount = 0;
    let successCount = 0;

    idSiswaList.forEach(function (idSiswa) {
      // Skip kalau sudah ada tagihan yang sama (pos + periode)
      if (existingSet.has(idSiswa)) {
        skippedCount++;
        return;
      }

      const siswa = siswaList.find(function (s) {
        return normalizeText(s.id_siswa) === idSiswa;
      });

      if (!siswa) {
        skippedCount++;
        return;
      }

      // Gunakan potongan dari form jika dipilih; jika tidak, gunakan potongan default siswa.
      let nominalPotongan = 0;
      const idPotongan = idPotonganInput === "NONE"
        ? ""
        : idPotonganInput || normalizeText(siswa.id_potongan_default);

      if (idPotongan && idPotongan !== "NONE") {
        const pData = potonganList.find(function (p) {
          return normalizeText(p.id_potongan) === idPotongan;
        });

        if (pData) {
          const tipe = normalizeText(pData.tipe_nilai).toUpperCase();
          if (tipe === "PERSEN") {
            nominalPotongan =
              (nominalAsal * toNumber(pData.nilai_potongan)) / 100;
          } else if (tipe === "NOMINAL") {
            nominalPotongan = toNumber(pData.nilai_potongan);
          }
        }
      }

      nominalPotongan = Math.min(Math.max(0, nominalPotongan), nominalAsal);
      const nominalAkhir = Math.max(0, nominalAsal - nominalPotongan);

      maxSequence++;
      const idTagihan = prefix + String(maxSequence).padStart(4, "0");

      // Buat baris sesuai urutan header yang ada
      const row = new Array(lastColumn).fill("");

      // Mapping berdasarkan header yang Anda miliki
      if (tagihanMap.id_tagihan !== undefined)
        row[tagihanMap.id_tagihan] = idTagihan;
      if (tagihanMap.id_siswa !== undefined) row[tagihanMap.id_siswa] = idSiswa;
      if (tagihanMap.id_pos !== undefined) row[tagihanMap.id_pos] = idPos;
      if (tagihanMap.id_potongan !== undefined)
        row[tagihanMap.id_potongan] = idPotongan || "";
      if (tagihanMap.tahun_pelajaran !== undefined)
        row[tagihanMap.tahun_pelajaran] = siswa.tahun_pelajaran || "";
      if (tagihanMap.periode !== undefined) row[tagihanMap.periode] = periode;
      if (tagihanMap.nama_item_snapshot !== undefined) {
        row[tagihanMap.nama_item_snapshot] =
          pos.nama_pos + (periode ? " (" + periode + ")" : "");
      }
      if (tagihanMap.nominal_asal !== undefined)
        row[tagihanMap.nominal_asal] = nominalAsal;
      if (tagihanMap.terbayar !== undefined) row[tagihanMap.terbayar] = 0;
      if (tagihanMap.sisa_tunggakan !== undefined)
        row[tagihanMap.sisa_tunggakan] = nominalAkhir;
      if (tagihanMap.bisa_dicicil !== undefined)
        row[tagihanMap.bisa_dicicil] = bisaDicicil;

      // status_siswa → isi dengan status siswa saat ini (AKTIF)
      if (tagihanMap.status_siswa !== undefined) {
        row[tagihanMap.status_siswa] =
          normalizeText(siswa.status || siswa.status_siswa) || "AKTIF";
      }

      if (tagihanMap.tanggal_tagihan !== undefined)
        row[tagihanMap.tanggal_tagihan] = now;
      if (tagihanMap.keterangan !== undefined) {
        row[tagihanMap.keterangan] =
          catatan || "Tagihan " + pos.nama_pos + " " + periode;
      }
      if (tagihanMap.nominal_potongan !== undefined)
        row[tagihanMap.nominal_potongan] = nominalPotongan;
      if (tagihanMap.nominal_akhir !== undefined)
        row[tagihanMap.nominal_akhir] = nominalAkhir;
      if (tagihanMap.status !== undefined) {
        row[tagihanMap.status] = nominalAkhir <= 0 ? "LUNAS" : "BELUM_BAYAR";
      }

      newRows.push(row);
      successCount++;
    });

    // Tulis batch
    if (newRows.length > 0) {
      const startRow = sheetTagihan.getLastRow() + 1;
      sheetTagihan
        .getRange(startRow, 1, newRows.length, lastColumn)
        .setValues(newRows);
    }

    return {
      success: true,
      message:
        "Berhasil generate " +
        successCount +
        " tagihan. " +
        (skippedCount > 0
          ? "Dilewati: " + skippedCount + " (sudah ada / siswa tidak valid)."
          : ""),
      generated: successCount,
      skipped: skippedCount,
    };
  } catch (err) {
    return {
      success: false,
      message: "Gagal generate tagihan: " + err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 23. RIWAYAT GENERATE TAGIHAN (PAGINATION) - FINAL
// ============================================================

function getRiwayatGenerateTagihan(page, limit) {
  try {
    page = Math.max(1, parseInt(page, 10) || 1);
    limit = Math.min(Math.max(parseInt(limit, 10) || 10, 5), 50);

    const tagihanList = getSheetDataAsObjects("Tagihan_Siswa");
    const siswaList = getSheetDataAsObjects("Siswa");
    const posList = getSheetDataAsObjects("Master_Pos");

    const mapSiswa = {};
    siswaList.forEach(function (s) {
      mapSiswa[normalizeText(s.id_siswa)] = s;
    });

    const mapPos = {};
    posList.forEach(function (p) {
      mapPos[normalizeText(p.id_pos)] = p;
    });

    const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

    const semua = tagihanList.map(function (t) {
      const idSiswa = normalizeText(t.id_siswa);
      const siswa = mapSiswa[idSiswa] || {};
      const pos = mapPos[normalizeText(t.id_pos)] || {};

      // ---------- TANGGAL TAGIHAN ----------
      let tanggalMs = 0;
      let tanggalTampil = "-";

      const rawTgl = t.tanggal_tagihan;
      if (rawTgl instanceof Date && !isNaN(rawTgl.getTime())) {
        tanggalMs = rawTgl.getTime();
        tanggalTampil = Utilities.formatDate(rawTgl, timezone, "dd/MM/yyyy");
      } else if (rawTgl) {
        const d = new Date(rawTgl);
        if (!isNaN(d.getTime())) {
          tanggalMs = d.getTime();
          tanggalTampil = Utilities.formatDate(d, timezone, "dd/MM/yyyy");
        }
      }

      // ---------- PERIODE (pastikan selalu teks bersih) ----------
      let periodeTampil = "";
      const rawPeriode = t.periode;

      if (rawPeriode instanceof Date && !isNaN(rawPeriode.getTime())) {
        // Kalau tidak sengaja terisi Date, format jadi yyyy-MM
        periodeTampil = Utilities.formatDate(rawPeriode, timezone, "yyyy-MM");
      } else {
        periodeTampil = normalizeText(rawPeriode);
      }

      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const sisa = Math.max(
        0,
        toNumber(t.sisa_tunggakan) || nominalAkhir - terbayar,
      );

      return {
        id_tagihan: normalizeText(t.id_tagihan),
        id_siswa: idSiswa,
        nama_siswa:
          normalizeText(siswa.nama_lengkap || siswa.nama_siswa) || "-",
        kelas: normalizeText(siswa.kelas) || "-",
        id_pos: normalizeText(t.id_pos),
        nama_pos:
          normalizeText(pos.nama_pos) ||
          normalizeText(t.nama_item_snapshot) ||
          "-",
        periode: periodeTampil || "-",
        nominal_akhir: nominalAkhir,
        terbayar: terbayar,
        sisa_tunggakan: sisa,
        status: normalizeText(t.status) || "BELUM_BAYAR",
        tanggal_tagihan: tanggalTampil, // sudah string "dd/MM/yyyy"
        tanggal_ms: tanggalMs,
        bisa_dicicil: toBoolean(t.bisa_dicicil),
      };
    });

    // Urutkan terbaru di atas
    semua.sort(function (a, b) {
      if (b.tanggal_ms !== a.tanggal_ms) {
        return b.tanggal_ms - a.tanggal_ms;
      }
      return String(b.id_tagihan).localeCompare(String(a.id_tagihan));
    });

    const total = semua.length;
    const start = (page - 1) * limit;
    const data = semua.slice(start, start + limit);

    // Bersihkan field internal
    data.forEach(function (row) {
      delete row.tanggal_ms;
    });

    return {
      success: true,
      data: data,
      total: total,
      page: page,
      limit: limit,
      total_pages: Math.ceil(total / limit) || 1,
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
      data: [],
      total: 0,
      page: 1,
      limit: 10,
      total_pages: 1,
    };
  }
}

// ============================================================
// 24. MASTER DATA (MODULAR)
// ============================================================

const MASTER_CONFIG = {
  Master_Pos: {
    primaryKey: "id_pos",
    label: "Master Pos",
  },
  Master_Potongan: {
    primaryKey: "id_potongan",
    label: "Master Potongan",
  },
  Pos_Pembayaran: {
    primaryKey: "id_pos",
    label: "Pos Pembayaran",
  },
};

function getMasterData(sheetName) {
  sheetName = normalizeText(sheetName);

  if (!MASTER_CONFIG[sheetName]) {
    throw new Error("Sheet master tidak dikenali: " + sheetName);
  }

  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();

  if (values.length === 0) {
    return {
      success: true,
      headers: [],
      data: [],
      primaryKey: MASTER_CONFIG[sheetName].primaryKey,
    };
  }

  const headers = values[0]
    .map(function (h) {
      return normalizeText(h);
    })
    .filter(Boolean);

  const data = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = { _rowIndex: i + 1 }; // untuk update/delete

    headers.forEach(function (header, idx) {
      let value = row[idx];

      // Konversi Date ke string biar aman
      if (value instanceof Date) {
        value = Utilities.formatDate(
          value,
          Session.getScriptTimeZone() || "Asia/Jakarta",
          "yyyy-MM-dd",
        );
      }

      obj[header] = value;
    });

    data.push(obj);
  }

  return {
    success: true,
    headers: headers,
    data: data,
    primaryKey: MASTER_CONFIG[sheetName].primaryKey,
    sheetName: sheetName,
  };
}

function saveMasterData(sheetName, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    sheetName = normalizeText(sheetName);

    if (!MASTER_CONFIG[sheetName]) {
      throw new Error("Sheet master tidak dikenali: " + sheetName);
    }

    if (!payload || typeof payload !== "object") {
      throw new Error("Data tidak valid.");
    }

    const sheet = getSheet(sheetName);
    const headerMap = getHeaderMap(sheet);
    const headers = Object.keys(headerMap);
    const primaryKey = MASTER_CONFIG[sheetName].primaryKey;
    const lastColumn = sheet.getLastColumn();

    const isEdit = payload._rowIndex && Number(payload._rowIndex) > 1;
    const rowIndex = isEdit
      ? Number(payload._rowIndex)
      : sheet.getLastRow() + 1;

    // Validasi primary key
    const pkValue = normalizeText(payload[primaryKey]);
    if (!pkValue) {
      throw new Error("Primary key (" + primaryKey + ") wajib diisi.");
    }

    // Cek duplikat primary key (kecuali saat edit baris yang sama)
    const existing = getSheetDataAsObjects(sheetName);
    const duplicate = existing.find(function (row, idx) {
      const existingPk = normalizeText(row[primaryKey]);
      const existingRowIndex = idx + 2; // karena data mulai baris 2
      return existingPk === pkValue && existingRowIndex !== rowIndex;
    });

    if (duplicate) {
      throw new Error(primaryKey + ' "' + pkValue + '" sudah digunakan.');
    }

    // Susun baris sesuai urutan header
    const rowValues = new Array(lastColumn).fill("");

    headers.forEach(function (header) {
      if (header === "_rowIndex") return;

      let value = payload[header];

      // Handle boolean
      if (typeof value === "boolean") {
        value = value;
      } else if (value === "true" || value === "TRUE" || value === "1") {
        value = true;
      } else if (value === "false" || value === "FALSE" || value === "0") {
        value = false;
      }

      if (headerMap[header] !== undefined) {
        rowValues[headerMap[header]] =
          value !== undefined && value !== null ? value : "";
      }
    });

    sheet.getRange(rowIndex, 1, 1, lastColumn).setValues([rowValues]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: isEdit
        ? "Data berhasil diupdate."
        : "Data berhasil ditambahkan.",
      rowIndex: rowIndex,
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterData(sheetName, rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    sheetName = normalizeText(sheetName);

    if (!MASTER_CONFIG[sheetName]) {
      throw new Error("Sheet master tidak dikenali: " + sheetName);
    }

    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) {
      throw new Error("Baris tidak valid.");
    }

    const sheet = getSheet(sheetName);
    if (rowIndex > sheet.getLastRow()) {
      throw new Error("Baris tidak ditemukan.");
    }

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Data berhasil dihapus.",
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 25. DATA SISWA (CRUD + EXPORT + IMPORT)
// ============================================================

function getSiswaList() {
  const siswaList = getSheetDataAsObjects("Siswa");
  const potonganList = getSheetDataAsObjects("Master_Potongan");

  const mapPotongan = {};
  potonganList.forEach(function (p) {
    mapPotongan[normalizeText(p.id_potongan)] = p;
  });

  const data = siswaList.map(function (s, idx) {
    const idPotongan = normalizeText(s.id_potongan_default);
    const potongan = mapPotongan[idPotongan];

    return {
      _rowIndex: idx + 2, // baris di sheet (header = 1)
      id_siswa: normalizeText(s.id_siswa),
      nisn: normalizeText(s.nisn),
      nis: normalizeText(s.nis),
      nama_lengkap: normalizeText(s.nama_lengkap || s.nama_siswa),
      jenis_kelamin: normalizeText(s.jenis_kelamin),
      kelas: normalizeText(s.kelas),
      tahun_pelajaran: normalizeText(s.tahun_pelajaran),
      status_siswa: normalizeText(s.status_siswa || s.status) || "AKTIF",
      id_potongan_default: idPotongan,
      nama_potongan: potongan
        ? normalizeText(potongan.nama_potongan)
        : "Tidak Ada",
    };
  });

  // Urutkan berdasarkan kelas lalu nama
  data.sort(function (a, b) {
    const k = a.kelas.localeCompare(b.kelas, undefined, { numeric: true });
    if (k !== 0) return k;
    return a.nama_lengkap.localeCompare(b.nama_lengkap);
  });

  return {
    success: true,
    data: data,
    total: data.length,
  };
}

function saveSiswa(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    if (!payload || typeof payload !== "object") {
      throw new Error("Data siswa tidak valid.");
    }

    const sheet = getSheet("Siswa");
    const headerMap = getHeaderMap(sheet);
    const lastColumn = sheet.getLastColumn();
    const primaryKey = "id_siswa";

    const idSiswa = normalizeText(payload.id_siswa);
    if (!idSiswa) {
      throw new Error("id_siswa wajib diisi.");
    }

    const isEdit = payload._rowIndex && Number(payload._rowIndex) > 1;
    let rowIndex = isEdit ? Number(payload._rowIndex) : sheet.getLastRow() + 1;

    // Cek duplikat id_siswa
    const existing = getSheetDataAsObjects("Siswa");
    const duplicate = existing.find(function (row, idx) {
      return normalizeText(row.id_siswa) === idSiswa && idx + 2 !== rowIndex;
    });

    if (duplicate) {
      throw new Error('id_siswa "' + idSiswa + '" sudah digunakan.');
    }

    const rowValues = new Array(lastColumn).fill("");

    // Mapping field
    const fields = [
      "id_siswa",
      "nisn",
      "nis",
      "nama_lengkap",
      "jenis_kelamin",
      "kelas",
      "tahun_pelajaran",
      "status_siswa",
      "id_potongan_default",
    ];

    fields.forEach(function (field) {
      if (headerMap[field] !== undefined) {
        let val = payload[field];
        if (val === undefined || val === null) val = "";
        rowValues[headerMap[field]] = val;
      }
    });

    // Handle status (beberapa sheet pakai "status")
    if (
      headerMap.status !== undefined &&
      headerMap.status_siswa === undefined
    ) {
      rowValues[headerMap.status] = payload.status_siswa || "AKTIF";
    }

    sheet.getRange(rowIndex, 1, 1, lastColumn).setValues([rowValues]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: isEdit
        ? "Data siswa berhasil diupdate."
        : "Data siswa berhasil ditambahkan.",
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteSiswa(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) {
      throw new Error("Baris tidak valid.");
    }

    const sheet = getSheet("Siswa");
    if (rowIndex > sheet.getLastRow()) {
      throw new Error("Baris tidak ditemukan.");
    }

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Data siswa berhasil dihapus.",
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Export data siswa ke sheet baru (bisa diunduh sebagai Excel)
 */
/**
 * Export data siswa → langsung file Excel (.xlsx)
 */
function exportSiswaToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getSiswaList().data || [];

  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone() || "Asia/Jakarta",
    "yyyyMMdd_HHmmss",
  );

  // Buat spreadsheet sementara khusus untuk export
  const tempSS = SpreadsheetApp.create("Export_Siswa_" + timestamp);
  const sheet = tempSS.getActiveSheet();
  sheet.setName("Data Siswa");

  const headers = [
    "id_siswa",
    "nisn",
    "nis",
    "nama_lengkap",
    "jenis_kelamin",
    "kelas",
    "tahun_pelajaran",
    "status_siswa",
    "id_potongan_default",
  ];

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#d1fae5");

  if (data.length > 0) {
    const rows = data.map(function (s) {
      return [
        s.id_siswa,
        s.nisn,
        s.nis,
        s.nama_lengkap,
        s.jenis_kelamin,
        s.kelas,
        s.tahun_pelajaran,
        s.status_siswa,
        s.id_potongan_default,
      ];
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);

  // Buat link download Excel
  const fileId = tempSS.getId();
  const url =
    "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

  // Opsional: hapus file sementara setelah beberapa jam (bisa pakai trigger nanti)
  // Untuk sekarang biarkan user yang download

  return {
    success: true,
    downloadUrl: url,
    fileName: "Data_Siswa_" + timestamp + ".xlsx",
    message: "File Excel siap diunduh (" + data.length + " data).",
  };
}

/**
 * Import data siswa
 * - Jika id_siswa sudah ada → Update
 * - Jika id_siswa belum ada → Insert
 */
function importSiswa(dataArray) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!Array.isArray(dataArray) || dataArray.length === 0) {
      throw new Error("Data import kosong.");
    }

    const sheet = getSheet("Siswa");
    const headerMap = getHeaderMap(sheet);
    const lastColumn = sheet.getLastColumn();
    const existing = getSheetDataAsObjects("Siswa");

    // Map id_siswa → rowIndex
    const mapExisting = {};
    existing.forEach(function (row, idx) {
      const id = normalizeText(row.id_siswa);
      if (id) {
        mapExisting[id] = idx + 2;
      }
    });

    let inserted = 0;
    let updated = 0;
    let skipped = 0;

    dataArray.forEach(function (item) {
      const idSiswa = normalizeText(item.id_siswa);

      if (!idSiswa) {
        skipped++;
        return;
      }

      const rowValues = new Array(lastColumn).fill("");

      const fields = [
        "id_siswa",
        "nisn",
        "nis",
        "nama_lengkap",
        "jenis_kelamin",
        "kelas",
        "tahun_pelajaran",
        "status_siswa",
        "id_potongan_default",
      ];

      fields.forEach(function (field) {
        if (headerMap[field] !== undefined) {
          let val = item[field];
          if (val === undefined || val === null) val = "";
          rowValues[headerMap[field]] = val;
        }
      });

      // status fallback
      if (
        headerMap.status !== undefined &&
        headerMap.status_siswa === undefined
      ) {
        rowValues[headerMap.status] = item.status_siswa || "AKTIF";
      }

      if (mapExisting[idSiswa]) {
        // UPDATE
        const rowIndex = mapExisting[idSiswa];
        sheet.getRange(rowIndex, 1, 1, lastColumn).setValues([rowValues]);
        updated++;
      } else {
        // INSERT
        const newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1, 1, lastColumn).setValues([rowValues]);
        mapExisting[idSiswa] = newRow; // hindari duplikat dalam 1 batch
        inserted++;
      }
    });

    SpreadsheetApp.flush();

    return {
      success: true,
      message:
        "Import selesai. " +
        "Insert: " +
        inserted +
        " | Update: " +
        updated +
        " | Dilewati: " +
        skipped,
      inserted: inserted,
      updated: updated,
      skipped: skipped,
    };
  } catch (err) {
    return {
      success: false,
      message: "Gagal import: " + err.message,
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Riwayat Setting Tagihan dikelompokkan per Kelas + Pos + Periode
 * Join ke sheet Siswa untuk mendapatkan Kelas
 */
function getRiwayatSettingKelas(page, perPage) {
  page = page || 1;
  perPage = perPage || 10;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ===== 1. Sheet Tagihan =====
    let sheetTagihan =
      ss.getSheetByName("Tagihan_Siswa") ||
      ss.getSheetByName("Tagihan") ||
      ss.getSheetByName("tagihan") ||
      ss.getSheetByName("TAGIHAN");

    if (!sheetTagihan) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().toLowerCase().indexOf("tagihan") !== -1) {
          sheetTagihan = sheets[i];
          break;
        }
      }
    }

    if (!sheetTagihan) {
      return {
        data: [],
        total: 0,
        page: 1,
        error: "Sheet Tagihan tidak ditemukan",
      };
    }

    const dataTagihan = sheetTagihan.getDataRange().getValues();
    if (dataTagihan.length <= 1) {
      return { data: [], total: 0, page: 1, error: "Sheet Tagihan kosong" };
    }

    const headersT = dataTagihan[0].map(function (h) {
      return String(h || "")
        .toLowerCase()
        .trim();
    });

    function findCol(headers, names) {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      for (let i = 0; i < headers.length; i++) {
        for (let j = 0; j < names.length; j++) {
          if (headers[i].indexOf(names[j]) !== -1) return i;
        }
      }
      return -1;
    }

    const col = {
      idSiswa: findCol(headersT, ["id_siswa"]),
      idPos: findCol(headersT, ["id_pos"]),
      periode: findCol(headersT, ["periode"]),
      namaItem: findCol(headersT, ["nama_item_snapshot"]),
      nominal: findCol(headersT, ["nominal_akhir", "nominal_asal", "nominal"]),
      tanggal: findCol(headersT, ["tanggal_tagihan"]),
      status: findCol(headersT, ["status"]),
    };

    if (col.idSiswa === -1) {
      return {
        data: [],
        total: 0,
        page: 1,
        error: "Kolom id_siswa tidak ditemukan di sheet Tagihan",
      };
    }

    // ===== 2. Sheet Siswa → map id_siswa → kelas =====
    let sheetSiswa =
      ss.getSheetByName("Siswa") ||
      ss.getSheetByName("siswa") ||
      ss.getSheetByName("SISWA");

    if (!sheetSiswa) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const n = sheets[i].getName().toLowerCase();
        if (
          n === "siswa" ||
          (n.indexOf("siswa") !== -1 && n.indexOf("tagihan") === -1)
        ) {
          sheetSiswa = sheets[i];
          break;
        }
      }
    }

    const mapSiswaKelas = {};

    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      if (dataSiswa.length > 1) {
        const headersS = dataSiswa[0].map(function (h) {
          return String(h || "")
            .toLowerCase()
            .trim();
        });

        const colSiswaId = findCol(headersS, ["id_siswa"]);
        const colSiswaKelas = findCol(headersS, ["kelas"]);

        if (colSiswaId !== -1 && colSiswaKelas !== -1) {
          for (let i = 1; i < dataSiswa.length; i++) {
            const id = String(dataSiswa[i][colSiswaId] || "").trim();
            const kelas = String(dataSiswa[i][colSiswaKelas] || "").trim();
            if (id) {
              mapSiswaKelas[id] = kelas || "-";
            }
          }
        }
      }
    }

    // ===== 3. Agregasi =====
    const groupMap = {};

    for (let i = 1; i < dataTagihan.length; i++) {
      const row = dataTagihan[i];

      const idSiswa = String(row[col.idSiswa] || "").trim();
      if (!idSiswa) continue;

      const kelas = mapSiswaKelas[idSiswa] || "Tidak diketahui";

      // Nama Pos dari nama_item_snapshot
      let posNama = "-";
      if (col.namaItem !== -1 && row[col.namaItem]) {
        posNama = String(row[col.namaItem]).trim();
      }

      const periode =
        col.periode !== -1 ? String(row[col.periode] || "").trim() : "-";

      // Tanggal
      let tanggalStr = "-";
      let tanggalBatch = "";
      if (col.tanggal !== -1 && row[col.tanggal]) {
        const t = row[col.tanggal];
        if (
          Object.prototype.toString.call(t) === "[object Date]" &&
          !isNaN(t.getTime())
        ) {
          tanggalStr = Utilities.formatDate(
            t,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy",
          );
          tanggalBatch = Utilities.formatDate(
            t,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
        } else {
          tanggalStr = String(t);
          tanggalBatch = tanggalStr;
        }
      }

      const key = kelas + "|" + posNama + "|" + periode + "|" + tanggalBatch;

      if (!groupMap[key]) {
        groupMap[key] = {
          tanggal: tanggalStr,
          tanggalBatch: tanggalBatch,
          kelas: kelas,
          pos: posNama,
          periode: periode,
          jumlahSiswa: 0,
          nominal: 0,
          totalTagihan: 0,
          statusCount: { lunas: 0, sebagian: 0, belum: 0 },
          status: "Belum Lunas",
        };
      }

      const g = groupMap[key];
      g.jumlahSiswa += 1;

      const nominal = col.nominal !== -1 ? Number(row[col.nominal]) || 0 : 0;
      if (g.nominal === 0 && nominal > 0) g.nominal = nominal;
      g.totalTagihan += nominal;

      if (col.status !== -1) {
        const st = String(row[col.status] || "").toLowerCase();
        if (st.indexOf("lunas") !== -1) {
          g.statusCount.lunas += 1;
        } else if (
          st.indexOf("sebagian") !== -1 ||
          st.indexOf("cicil") !== -1
        ) {
          g.statusCount.sebagian += 1;
        } else {
          g.statusCount.belum += 1;
        }
      } else {
        g.statusCount.belum += 1;
      }
    }

    // Status final
    const list = [];
    for (const key in groupMap) {
      const g = groupMap[key];
      if (g.jumlahSiswa > 0 && g.statusCount.lunas === g.jumlahSiswa) {
        g.status = "Lunas";
      } else if (g.statusCount.lunas > 0 || g.statusCount.sebagian > 0) {
        g.status = "Sebagian";
      } else {
        g.status = "Belum Lunas";
      }
      list.push(g);
    }

    // Urutkan terbaru
    list.sort(function (a, b) {
      if (a.tanggalBatch && b.tanggalBatch) {
        return b.tanggalBatch.localeCompare(a.tanggalBatch);
      }
      return 0;
    });

    const total = list.length;
    const start = (page - 1) * perPage;
    const paged = list.slice(start, start + perPage);

    return {
      data: paged,
      total: total,
      page: page,
      _debug: {
        sheetTagihan: sheetTagihan.getName(),
        sheetSiswa: sheetSiswa ? sheetSiswa.getName() : "TIDAK DITEMUKAN",
        jumlahMapSiswa: Object.keys(mapSiswaKelas).length,
        totalGroup: total,
        headersTagihan: headersT,
      },
    };
  } catch (err) {
    return {
      data: [],
      total: 0,
      page: 1,
      error: err.message,
    };
  }
}

/**
 * Hapus satu tagihan berdasarkan ID.
 */
function hapusTagihanSiswa(idTagihan) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    idTagihan = normalizeText(idTagihan);
    if (!idTagihan) throw new Error("ID tagihan wajib diisi.");

    const sheet = getSheet("Tagihan_Siswa");
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) throw new Error("Tagihan tidak ditemukan.");

    const headers = values[0].map(function (header) {
      return normalizeText(header).toLowerCase();
    });
    const idCol = headers.indexOf("id_tagihan");
    const paidCol = headers.indexOf("terbayar");
    const statusCol = headers.indexOf("status");

    if (idCol === -1) throw new Error("Kolom id_tagihan tidak ditemukan.");

    for (let i = 1; i < values.length; i++) {
      if (normalizeText(values[i][idCol]) !== idTagihan) continue;

      const terbayar = paidCol === -1 ? 0 : toNumber(values[i][paidCol]);
      const status = statusCol === -1
        ? ""
        : normalizeText(values[i][statusCol]).toUpperCase();
      if (terbayar > 0 || status === "LUNAS" || status === "CICILAN") {
        throw new Error("Tagihan yang sudah memiliki pembayaran tidak dapat dihapus.");
      }

      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return {
        success: true,
        message: "Tagihan " + idTagihan + " berhasil dihapus.",
      };
    }

    throw new Error("Tagihan " + idTagihan + " tidak ditemukan.");
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hapus semua baris Tagihan yang termasuk group:
 * Kelas + Pos (nama_item_snapshot) + Periode + Tanggal
 */
function hapusSettingKelas(kelas, pos, periode, tanggalBatch) {
  try {
    kelas = String(kelas || "").trim();
    pos = String(pos || "").trim();
    periode = String(periode || "").trim();
    tanggalBatch = String(tanggalBatch || "").trim();

    if (!kelas || !pos) {
      return {
        success: false,
        message: "Parameter kelas/pos kosong",
        deleted: 0,
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ===== Sheet Tagihan =====
    let sheetTagihan =
      ss.getSheetByName("Tagihan_Siswa") ||
      ss.getSheetByName("Tagihan") ||
      ss.getSheetByName("tagihan") ||
      ss.getSheetByName("TAGIHAN");

    if (!sheetTagihan) {
      return {
        success: false,
        message: "Sheet Tagihan tidak ditemukan",
        deleted: 0,
      };
    }

    const dataTagihan = sheetTagihan.getDataRange().getValues();
    if (dataTagihan.length <= 1) {
      return { success: true, deleted: 0, message: "Tidak ada data" };
    }

    const headersT = dataTagihan[0].map(function (h) {
      return String(h || "")
        .toLowerCase()
        .trim();
    });

    function findCol(headers, names) {
      for (var i = 0; i < names.length; i++) {
        var idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      for (var i = 0; i < headers.length; i++) {
        for (var j = 0; j < names.length; j++) {
          if (headers[i].indexOf(names[j]) !== -1) return i;
        }
      }
      return -1;
    }

    const col = {
      idSiswa: findCol(headersT, ["id_siswa"]),
      periode: findCol(headersT, ["periode"]),
      namaItem: findCol(headersT, ["nama_item_snapshot"]),
      tanggal: findCol(headersT, ["tanggal_tagihan"]),
    };

    if (col.idSiswa === -1 || col.namaItem === -1) {
      return {
        success: false,
        message: "Kolom id_siswa / nama_item_snapshot tidak ditemukan",
        deleted: 0,
      };
    }

    // ===== Map id_siswa → kelas =====
    let sheetSiswa =
      ss.getSheetByName("Siswa") ||
      ss.getSheetByName("siswa") ||
      ss.getSheetByName("SISWA");

    const mapSiswaKelas = {};
    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      if (dataSiswa.length > 1) {
        const headersS = dataSiswa[0].map(function (h) {
          return String(h || "")
            .toLowerCase()
            .trim();
        });
        const colSiswaId = findCol(headersS, ["id_siswa"]);
        const colSiswaKelas = findCol(headersS, ["kelas"]);

        if (colSiswaId !== -1 && colSiswaKelas !== -1) {
          for (var i = 1; i < dataSiswa.length; i++) {
            var id = String(dataSiswa[i][colSiswaId] || "").trim();
            var kls = String(dataSiswa[i][colSiswaKelas] || "").trim();
            if (id) mapSiswaKelas[id] = kls || "-";
          }
        }
      }
    }

    // ===== Cari baris yang cocok (dari bawah ke atas agar aman dihapus) =====
    const rowsToDelete = []; // 1-based index di sheet

    for (var i = 1; i < dataTagihan.length; i++) {
      var row = dataTagihan[i];
      var idSiswa = String(row[col.idSiswa] || "").trim();
      if (!idSiswa) continue;

      var kelasRow = mapSiswaKelas[idSiswa] || "Tidak diketahui";
      if (kelasRow !== kelas) continue;

      var posRow = String(row[col.namaItem] || "").trim();
      if (posRow !== pos) continue;

      if (col.periode !== -1) {
        var periodeRow = String(row[col.periode] || "").trim();
        if (periode && periodeRow !== periode) continue;
      }

      // Cocokkan tanggal (yyyy-MM-dd)
      if (tanggalBatch && col.tanggal !== -1 && row[col.tanggal]) {
        var t = row[col.tanggal];
        var tBatch = "";
        if (
          Object.prototype.toString.call(t) === "[object Date]" &&
          !isNaN(t.getTime())
        ) {
          tBatch = Utilities.formatDate(
            t,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
        } else {
          tBatch = String(t);
        }
        if (tBatch !== tanggalBatch) continue;
      }

      rowsToDelete.push(i + 1); // sheet row number (1-based, +1 karena header di index 0)
    }

    // Hapus dari bawah supaya index tidak bergeser
    rowsToDelete.sort(function (a, b) {
      return b - a;
    });

    for (var j = 0; j < rowsToDelete.length; j++) {
      sheetTagihan.deleteRow(rowsToDelete[j]);
    }

    return {
      success: true,
      deleted: rowsToDelete.length,
      message: "Berhasil menghapus " + rowsToDelete.length + " baris",
    };
  } catch (err) {
    return {
      success: false,
      deleted: 0,
      message: err.message,
    };
  }
}

/**
 * Hapus maksimal `batchSize` baris yang cocok dengan group.
 * Baris yang sudah ada pembayaran (terbayar > 0) TIDAK dihapus.
 */
function hapusSettingKelasBatch(kelas, pos, periode, tanggalBatch, batchSize) {
  try {
    kelas = String(kelas || "").trim();
    pos = String(pos || "").trim();
    periode = String(periode || "").trim();
    tanggalBatch = String(tanggalBatch || "").trim();
    batchSize = Number(batchSize) || 20;
    if (batchSize < 1) batchSize = 20;

    if (!kelas || !pos) {
      return {
        success: false,
        deleted: 0,
        remaining: 0,
        skipped: 0,
        message: "Parameter kelas/pos kosong",
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetTagihan =
      ss.getSheetByName("Tagihan_Siswa") ||
      ss.getSheetByName("Tagihan") ||
      ss.getSheetByName("tagihan");

    if (!sheetTagihan) {
      return {
        success: false,
        deleted: 0,
        remaining: 0,
        skipped: 0,
        message: "Sheet Tagihan_Siswa tidak ditemukan",
      };
    }

    const dataTagihan = sheetTagihan.getDataRange().getValues();
    if (dataTagihan.length <= 1) {
      return {
        success: true,
        deleted: 0,
        remaining: 0,
        skipped: 0,
        totalMatched: 0,
        message: "Tidak ada data",
      };
    }

    const headersT = dataTagihan[0].map(function (h) {
      return String(h || "")
        .toLowerCase()
        .trim();
    });

    function findCol(headers, names) {
      for (var i = 0; i < names.length; i++) {
        var idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      for (var i = 0; i < headers.length; i++) {
        for (var j = 0; j < names.length; j++) {
          if (headers[i].indexOf(names[j]) !== -1) return i;
        }
      }
      return -1;
    }

    const col = {
      idSiswa: findCol(headersT, ["id_siswa"]),
      periode: findCol(headersT, ["periode"]),
      namaItem: findCol(headersT, ["nama_item_snapshot"]),
      tanggal: findCol(headersT, ["tanggal_tagihan"]),
      terbayar: findCol(headersT, ["terbayar"]),
      status: findCol(headersT, ["status"]),
    };

    if (col.idSiswa === -1 || col.namaItem === -1) {
      return {
        success: false,
        deleted: 0,
        remaining: 0,
        skipped: 0,
        message: "Kolom wajib tidak ditemukan",
      };
    }

    // Map siswa → kelas
    let sheetSiswa =
      ss.getSheetByName("Siswa") ||
      ss.getSheetByName("siswa") ||
      ss.getSheetByName("SISWA");

    const mapSiswaKelas = {};
    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      if (dataSiswa.length > 1) {
        const headersS = dataSiswa[0].map(function (h) {
          return String(h || "")
            .toLowerCase()
            .trim();
        });
        const colSiswaId = findCol(headersS, ["id_siswa"]);
        const colSiswaKelas = findCol(headersS, ["kelas"]);
        if (colSiswaId !== -1 && colSiswaKelas !== -1) {
          for (var i = 1; i < dataSiswa.length; i++) {
            var id = String(dataSiswa[i][colSiswaId] || "").trim();
            var kls = String(dataSiswa[i][colSiswaKelas] || "").trim();
            if (id) mapSiswaKelas[id] = kls || "-";
          }
        }
      }
    }

    /**
     * Cek apakah tagihan sudah ada pembayaran → tidak boleh dihapus
     */
    function isSudahDibayar(row) {
      // 1. Kolom terbayar > 0
      if (col.terbayar !== -1) {
        var bayar = Number(row[col.terbayar]) || 0;
        if (bayar > 0) return true;
      }
      // 2. Status lunas / sebagian / partial
      if (col.status !== -1) {
        var st = String(row[col.status] || "").toLowerCase();
        if (st.indexOf("lunas") !== -1) return true;
        if (st.indexOf("sebagian") !== -1) return true;
        if (st.indexOf("partial") !== -1) return true;
        if (st.indexOf("cicil") !== -1) return true;
      }
      return false;
    }

    // Kumpulkan baris yang boleh dihapus & yang di-skip
    const deletableRows = []; // 1-based sheet row
    var skippedPaid = 0;

    for (var i = 1; i < dataTagihan.length; i++) {
      var row = dataTagihan[i];
      var idSiswa = String(row[col.idSiswa] || "").trim();
      if (!idSiswa) continue;

      var kelasRow = mapSiswaKelas[idSiswa] || "Tidak diketahui";
      if (kelasRow !== kelas) continue;

      var posRow = String(row[col.namaItem] || "").trim();
      if (posRow !== pos) continue;

      if (col.periode !== -1) {
        var periodeRow = String(row[col.periode] || "").trim();
        if (periode && periodeRow !== periode) continue;
      }

      if (tanggalBatch && col.tanggal !== -1 && row[col.tanggal]) {
        var t = row[col.tanggal];
        var tBatch = "";
        if (
          Object.prototype.toString.call(t) === "[object Date]" &&
          !isNaN(t.getTime())
        ) {
          tBatch = Utilities.formatDate(
            t,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
        } else {
          tBatch = String(t);
        }
        if (tBatch !== tanggalBatch) continue;
      }

      // ===== PENGAMAN: sudah dibayar → skip =====
      if (isSudahDibayar(row)) {
        skippedPaid++;
        continue;
      }

      deletableRows.push(i + 1);
    }

    const totalDeletable = deletableRows.length;

    // Tidak ada yang bisa dihapus
    if (totalDeletable === 0) {
      return {
        success: true,
        deleted: 0,
        remaining: 0,
        skipped: skippedPaid,
        totalMatched: skippedPaid,
        message:
          skippedPaid > 0
            ? "Semua tagihan pada group ini sudah ada pembayaran (" +
              skippedPaid +
              " data dilindungi)."
            : "Tidak ada data yang cocok",
      };
    }

    // Hapus batch dari bawah
    deletableRows.sort(function (a, b) {
      return b - a;
    });
    const toDeleteNow = deletableRows.slice(0, batchSize);

    for (var j = 0; j < toDeleteNow.length; j++) {
      sheetTagihan.deleteRow(toDeleteNow[j]);
    }

    const deleted = toDeleteNow.length;
    const remaining = totalDeletable - deleted;

    return {
      success: true,
      deleted: deleted,
      remaining: remaining,
      skipped: skippedPaid,
      totalMatched: totalDeletable + skippedPaid,
      totalDeletable: totalDeletable,
      message: "Batch OK",
    };
  } catch (err) {
    return {
      success: false,
      deleted: 0,
      remaining: 0,
      skipped: 0,
      message: err.message,
    };
  }
}
