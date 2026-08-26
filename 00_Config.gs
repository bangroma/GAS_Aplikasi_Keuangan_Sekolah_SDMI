/**
 * KONFIGURASI & HELPER DASAR
 * Sistem Keuangan Sekolah
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
    ["POS-SRG", "SRG", "Seragam Siswa", "SEKALI", 2900000, true, true],
    ["POS-KGT", "KGT", "Kegiatan Tahunan", "KEGIATAN", 240000, true, true],
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