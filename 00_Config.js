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
    "status_siswa",
    "id_potongan_default",
  ],

  // P1: membership siswa per tahun pelajaran
  Siswa_Tahun_Pelajaran: [
    "id_keanggotaan",
    "id_siswa",
    "tahun_pelajaran",
    "kelas",
    "status_siswa",
    "id_potongan_default",
    "tanggal_masuk",
    "tanggal_keluar",
    "keterangan",
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
    "diskon_tambahan_total",
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
function migrateDatabaseSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = {
    success: true,
    migrasi: [],
    pesan: [],
  };

  const sheet = ss.getSheetByName("Siswa");

  if (!sheet) {
    hasil.pesan.push("Sheet Siswa belum ada. Tidak ada migrasi yang diperlukan.");
    return hasil;
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    hasil.pesan.push("Sheet Siswa kosong. Tidak ada migrasi yang diperlukan.");
    return hasil;
  }

  let headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function (h) {
      return normalizeText(h);
    });

  const statusIndex = headers.indexOf("status");
  const statusSiswaIndex = headers.indexOf("status_siswa");

  if (statusIndex === -1 && statusSiswaIndex === -1) {
    hasil.pesan.push("Kolom status/status_siswa tidak ditemukan.");
    return hasil;
  }

  if (statusIndex === -1 && statusSiswaIndex !== -1) {
    hasil.pesan.push("Siswa sudah menggunakan status_siswa.");
    return hasil;
  }

  if (statusIndex !== -1 && statusSiswaIndex === -1) {
    const newColumn = sheet.getLastColumn() + 1;

    sheet
      .getRange(1, newColumn)
      .setValue("status_siswa");

    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const values = sheet
        .getRange(2, statusIndex + 1, lastRow - 1, 1)
        .getValues();

      sheet
        .getRange(2, newColumn, lastRow - 1, 1)
        .setValues(values);
    }

    sheet.deleteColumn(statusIndex + 1);

    hasil.migrasi.push("Siswa.status → Siswa.status_siswa");
    return hasil;
  }

  if (statusIndex !== -1 && statusSiswaIndex !== -1) {
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const statusValues = sheet
        .getRange(2, statusIndex + 1, lastRow - 1, 1)
        .getValues();

      const statusSiswaValues = sheet
        .getRange(2, statusSiswaIndex + 1, lastRow - 1, 1)
        .getValues();

      const mergedValues = statusSiswaValues.map(function (row, index) {
        return [
          row[0] !== "" && row[0] !== null
            ? row[0]
            : statusValues[index][0],
        ];
      });

      sheet
        .getRange(2, statusSiswaIndex + 1, lastRow - 1, 1)
        .setValues(mergedValues);
    }

    const deleteIndex = statusIndex;

    sheet.deleteColumn(deleteIndex + 1);

    hasil.migrasi.push(
      "Siswa.status digabung ke Siswa.status_siswa"
    );

    return hasil;
  }

  return hasil;
}
// ============================================================
// ============================================================
// P1_MEMBERSHIP_SCHEMA_AND_MIGRATION
// Fondasi keanggotaan siswa per tahun pelajaran.
// Tidak mengubah record lama pada sheet Siswa.
// ============================================================

function generateIdKeanggotaan_() {
  return (
    "ANG-" +
    Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase()
  );
}

function previewMigrasiSiswaKeMembership(tahunPelajaran) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siswaSheet = ss.getSheetByName("Siswa");
  const membershipSheet = ss.getSheetByName("Siswa_Tahun_Pelajaran");

  if (!siswaSheet) {
    return {
      success: false,
      message: "Sheet Siswa tidak ditemukan.",
    };
  }

  if (!membershipSheet) {
    return {
      success: false,
      message:
        "Sheet Siswa_Tahun_Pelajaran belum ada. Jalankan setupDatabase().",
    };
  }

  const tahunTarget = normalizeText(tahunPelajaran);

  if (!tahunTarget) {
    return {
      success: false,
      message: "Tahun pelajaran wajib diisi.",
    };
  }

  const siswaList = getSheetDataAsObjects("Siswa");

  const membershipList = getSheetDataAsObjects(
    "Siswa_Tahun_Pelajaran"
  );

  const existingKeys = new Set();

  membershipList.forEach(function (row) {
    const idSiswa = normalizeText(row.id_siswa);
    const tahun = normalizeText(row.tahun_pelajaran);

    if (idSiswa && tahun) {
      existingKeys.add(idSiswa + "|" + tahun);
    }
  });

  const data = [];
  let baru = 0;
  let sudahAda = 0;
  let tidakValid = 0;

  siswaList.forEach(function (siswa) {
    const idSiswa = normalizeText(siswa.id_siswa);
    const tahun = tahunTarget;
    const kelas = normalizeText(siswa.kelas);

    if (!idSiswa || !kelas) {
      tidakValid++;

      data.push({
        status: "TIDAK_VALID",
        id_siswa: idSiswa,
        tahun_pelajaran: tahun,
        kelas: kelas,
        nama_lengkap: normalizeText(siswa.nama_lengkap),
        keterangan:
          "id_siswa, tahun_pelajaran, atau kelas kosong",
      });

      return;
    }

    const key = idSiswa + "|" + tahun;

    if (existingKeys.has(key)) {
      sudahAda++;

      data.push({
        status: "SUDAH_ADA",
        id_siswa: idSiswa,
        tahun_pelajaran: tahun,
        kelas: kelas,
        nama_lengkap: normalizeText(siswa.nama_lengkap),
        keterangan: "Membership sudah tersedia",
      });

      return;
    }

    baru++;

    data.push({
      status: "BARU",
      id_siswa: idSiswa,
      tahun_pelajaran: tahun,
      kelas: kelas,
      nama_lengkap: normalizeText(siswa.nama_lengkap),
      keterangan: "Akan dibuat membership baru",
    });
  });

  return {
    success: true,
    total_siswa: siswaList.length,
    total_membership: membershipList.length,
    baru: baru,
    sudah_ada: sudahAda,
    tidak_valid: tidakValid,
    data: data,
  };
}

function migrasiSiswaKeMembership(tahunPelajaran) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const siswaSheet = ss.getSheetByName("Siswa");
    const membershipSheet = ss.getSheetByName("Siswa_Tahun_Pelajaran");

    if (!siswaSheet) {
      return {
        success: false,
        message: "Sheet Siswa tidak ditemukan.",
      };
    }

    if (!membershipSheet) {
      return {
        success: false,
        message:
          "Sheet Siswa_Tahun_Pelajaran belum ada. Jalankan setupDatabase().",
      };
    }

    const tahunTarget = normalizeText(tahunPelajaran);

    if (!tahunTarget) {
      return {
        success: false,
        message: "Tahun pelajaran wajib diisi.",
      };
    }

    const siswaList = getSheetDataAsObjects("Siswa");

    const membershipList = getSheetDataAsObjects(
      "Siswa_Tahun_Pelajaran"
    );

    const existingKeys = new Set();

    membershipList.forEach(function (row) {
      const idSiswa = normalizeText(row.id_siswa);
      const tahun = normalizeText(row.tahun_pelajaran);

      if (idSiswa && tahun) {
        existingKeys.add(idSiswa + "|" + tahun);
      }
    });

    const headerMap = getHeaderMap(membershipSheet);
    const lastColumn = membershipSheet.getLastColumn();

    const requiredHeaders = [
      "id_keanggotaan",
      "id_siswa",
      "tahun_pelajaran",
      "kelas",
      "status_siswa",
      "id_potongan_default",
      "tanggal_masuk",
      "tanggal_keluar",
      "keterangan",
    ];

    requiredHeaders.forEach(function (header) {
      if (headerMap[header] === undefined) {
        throw new Error(
          "Kolom " +
            header +
            " tidak ditemukan di Siswa_Tahun_Pelajaran."
        );
      }
    });

    const rows = [];
    let inserted = 0;
    let skipped = 0;
    let invalid = 0;

    siswaList.forEach(function (siswa) {
      const idSiswa = normalizeText(siswa.id_siswa);
      const tahun = tahunTarget;
      const kelas = normalizeText(siswa.kelas);

      if (!idSiswa || !kelas) {
        invalid++;
        return;
      }

      const key = idSiswa + "|" + tahun;

      if (existingKeys.has(key)) {
        skipped++;
        return;
      }

      const row = new Array(lastColumn).fill("");

      row[headerMap.id_keanggotaan] = generateIdKeanggotaan_();
      row[headerMap.id_siswa] = idSiswa;
      row[headerMap.tahun_pelajaran] = tahun;
      row[headerMap.kelas] = kelas;
      row[headerMap.status_siswa] =
        normalizeText(siswa.status_siswa) || "AKTIF";
      row[headerMap.id_potongan_default] =
        normalizeText(siswa.id_potongan_default);
      row[headerMap.tanggal_masuk] = "";
      row[headerMap.tanggal_keluar] = "";
      row[headerMap.keterangan] = "Migrasi awal dari Siswa";

      rows.push(row);
      existingKeys.add(key);
      inserted++;
    });

    if (rows.length > 0) {
      membershipSheet
        .getRange(
          membershipSheet.getLastRow() + 1,
          1,
          rows.length,
          lastColumn
        )
        .setValues(rows);

      SpreadsheetApp.flush();
    }

    return {
      success: true,
      inserted: inserted,
      skipped: skipped,
      invalid: invalid,
      message:
        "Migrasi selesai. Dibuat " +
        inserted +
        " membership, dilewati " +
        skipped +
        ", tidak valid " +
        invalid +
        ".",
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

function auditSiswaMembership() {
  const siswaList = getSheetDataAsObjects("Siswa");
  const membershipList = getSheetDataAsObjects(
    "Siswa_Tahun_Pelajaran"
  );

  const siswaIds = new Set();

  siswaList.forEach(function (siswa) {
    const idSiswa = normalizeText(siswa.id_siswa);

    if (idSiswa) {
      siswaIds.add(idSiswa);
    }
  });

  const membershipKeys = new Set();
  const duplicateKeys = [];
  const orphanMemberships = [];
  const tanpaSiswa = [];
  const tanpaTahun = [];

  membershipList.forEach(function (row) {
    const idKeanggotaan = normalizeText(row.id_keanggotaan);
    const idSiswa = normalizeText(row.id_siswa);
    const tahun = normalizeText(row.tahun_pelajaran);
    const label = idKeanggotaan || "(tanpa id_keanggotaan)";

    if (!idSiswa) {
      tanpaSiswa.push(label);
    }

    if (!tahun) {
      tanpaTahun.push(label);
    }

    if (!idSiswa || !tahun) {
      return;
    }

    const key = idSiswa + "|" + tahun;

    if (membershipKeys.has(key)) {
      duplicateKeys.push(key);
    } else {
      membershipKeys.add(key);
    }

    if (!siswaIds.has(idSiswa)) {
      orphanMemberships.push(
        label +
          " → id_siswa " +
          idSiswa +
          " tidak ditemukan di Siswa"
      );
    }
  });

  return {
    success: true,
    total_siswa: siswaList.length,
    total_membership: membershipList.length,
    unique_membership: membershipKeys.size,
    duplicate_id_siswa_tahun: duplicateKeys.length,
    membership_tanpa_id_siswa: tanpaSiswa.length,
    membership_tanpa_tahun_pelajaran: tanpaTahun.length,
    orphan_membership: orphanMemberships.length,
    duplicate_keys: duplicateKeys,
    tanpa_siswa: tanpaSiswa,
    tanpa_tahun: tanpaTahun,
    orphan_memberships: orphanMemberships,
    sehat:
      duplicateKeys.length === 0 &&
      tanpaSiswa.length === 0 &&
      tanpaTahun.length === 0 &&
      orphanMemberships.length === 0,
  };
}

// ============================================================
// P1_MEMBERSHIP_BACKEND_CONTRACTS
// Backend pemilihan dan penyimpanan membership siswa per tahun.
// ============================================================

function getKandidatSiswaMembership(tahunPelajaran) {
  const tahunTarget = normalizeText(tahunPelajaran);
  if (!tahunTarget) return { success:false, message:"Tahun pelajaran wajib diisi." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Siswa")) {
    return { success:false, message:"Sheet Siswa tidak ditemukan." };
  }
  if (!ss.getSheetByName("Siswa_Tahun_Pelajaran")) {
    return {
      success:false,
      message:"Sheet Siswa_Tahun_Pelajaran belum ada. Jalankan setupDatabase()."
    };
  }

  const siswaList = getSheetDataAsObjects("Siswa");
  const memberships = getSheetDataAsObjects("Siswa_Tahun_Pelajaran");
  const existing = new Set();

  memberships.forEach(function(row) {
    const id = normalizeText(row.id_siswa);
    const tahun = normalizeText(row.tahun_pelajaran);
    if (id && tahun) existing.add(id + "|" + tahun);
  });

  const data = siswaList
    .map(function(siswa) {
      const id = normalizeText(siswa.id_siswa);
      if (!id) return null;

      const sudah = existing.has(id + "|" + tahunTarget);

      return {
        status: sudah ? "SUDAH_TERDAFTAR" : "BELUM_TERDAFTAR",
        id_siswa: id,
        nisn: normalizeText(siswa.nisn),
        nis: normalizeText(siswa.nis),
        nama_lengkap: normalizeText(siswa.nama_lengkap),
        jenis_kelamin: normalizeText(siswa.jenis_kelamin),
        kelas_master: normalizeText(siswa.kelas),
        tahun_pelajaran_master: normalizeText(siswa.tahun_pelajaran),
        status_siswa_master: normalizeText(siswa.status_siswa),
        id_potongan_default_master: normalizeText(siswa.id_potongan_default),
        tahun_pelajaran_target: tahunTarget
      };
    })
    .filter(Boolean);

  return {
    success:true,
    tahun_pelajaran:tahunTarget,
    total:data.length,
    belum_terdaftar:data.filter(function(x) {
      return x.status === "BELUM_TERDAFTAR";
    }).length,
    sudah_terdaftar:data.filter(function(x) {
      return x.status === "SUDAH_TERDAFTAR";
    }).length,
    data:data
  };
}

// ============================================================
// P1_MEMBERSHIP_VALIDATION_HARDENED
// Patch validasi membership siswa per tahun pelajaran.
// ============================================================

function validateSimpanMembershipSiswa_(
  tahunPelajaran,
  items,
  siswaMap,
  existing
) {
  const tahunTarget = normalizeText(tahunPelajaran);

  if (!tahunTarget) {
    throw new Error("Tahun pelajaran wajib diisi.");
  }

  if (!Array.isArray(items) || items.length === 0) {
    throw new Error("Minimal satu siswa harus dipilih.");
  }

  const validStatusSiswa = new Set([
    "AKTIF",
    "NONAKTIF",
    "LULUS",
    "PINDAH"
  ]);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const potonganSheet = ss.getSheetByName("Master_Potongan");
  const potonganMap = {};

  if (potonganSheet) {
    getSheetDataAsObjects("Master_Potongan").forEach(function(row) {
      const idPotongan = normalizeText(row.id_potongan);

      if (idPotongan) {
        potonganMap[idPotongan] = row;
      }
    });
  }

  const seen = new Set();
  const result = [];

  items.forEach(function(item, index) {
    if (!item || typeof item !== "object") {
      throw new Error(
        "Data siswa pada baris " + (index + 1) + " tidak valid."
      );
    }

    const idSiswa = normalizeText(item.id_siswa);

    if (!idSiswa) {
      throw new Error(
        "ID siswa pada baris " + (index + 1) + " wajib diisi."
      );
    }

    if (seen.has(idSiswa)) {
      throw new Error(
        "Siswa " + idSiswa + " dipilih lebih dari satu kali."
      );
    }

    seen.add(idSiswa);

    const siswa = siswaMap[idSiswa];

    if (!siswa) {
      throw new Error(
        "Siswa " + idSiswa + " tidak ditemukan pada sheet Siswa."
      );
    }

    const kelas = normalizeText(siswa.kelas);

    if (!kelas) {
      throw new Error(
        "Kelas siswa " + idSiswa +
        " belum tersedia pada sheet Siswa."
      );
    }

    const membershipKey = idSiswa + "|" + tahunTarget;

    if (existing.has(membershipKey)) {
      throw new Error(
        "Siswa " + idSiswa +
        " sudah memiliki membership tahun " +
        tahunTarget + "."
      );
    }

    const statusSiswa =
      normalizeText(siswa.status_siswa).toUpperCase();

    if (statusSiswa && !validStatusSiswa.has(statusSiswa)) {
      throw new Error(
        "Status siswa " + idSiswa +
        " tidak valid: " + statusSiswa +
        ". Gunakan AKTIF, NONAKTIF, LULUS, atau PINDAH."
      );
    }

    const idPotonganDefault =
      normalizeText(siswa.id_potongan_default);

    if (idPotonganDefault) {
      if (!potonganSheet) {
        throw new Error(
          "Master_Potongan tidak ditemukan untuk id_potongan: " +
          idPotonganDefault + "."
        );
      }

      if (!potonganMap[idPotonganDefault]) {
        throw new Error(
          "id_potongan_default " + idPotonganDefault +
          " untuk siswa " + idSiswa +
          " tidak ditemukan pada Master_Potongan."
        );
      }
    }

    result.push({
      id_keanggotaan: generateIdKeanggotaan_(),
      id_siswa: idSiswa,
      tahun_pelajaran: tahunTarget,
      kelas: kelas,
      status_siswa: statusSiswa,
      id_potongan_default: idPotonganDefault,
      tanggal_masuk: item.tanggal_masuk || "",
      tanggal_keluar: item.tanggal_keluar || "",
      keterangan: normalizeText(item.keterangan) ||
        "Membership siswa per tahun pelajaran"
    });
  });

  return result;
}


// ============================================================
// END P1_MEMBERSHIP_VALIDATION_HARDENED
// ============================================================

function prepareMembershipValidation_(tahunPelajaran, items) {
  const tahunTarget = normalizeText(tahunPelajaran);

  if (!tahunTarget) {
    throw new Error("Tahun pelajaran wajib diisi.");
  }

  if (!Array.isArray(items) || items.length === 0) {
    throw new Error("Minimal satu siswa harus dipilih.");
  }

  const siswaList = getSheetDataAsObjects("Siswa");
  const memberships = getSheetDataAsObjects("Siswa_Tahun_Pelajaran");

  const siswaMap = {};
  siswaList.forEach(function(siswa) {
    const id = normalizeText(siswa.id_siswa);
    if (id) siswaMap[id] = siswa;
  });

  const existing = new Set();
  memberships.forEach(function(row) {
    const id = normalizeText(row.id_siswa);
    const tahun = normalizeText(row.tahun_pelajaran);
    if (id && tahun) {
      existing.add(id + "|" + tahun);
    }
  });

  const validated = validateSimpanMembershipSiswa_(
    tahunTarget,
    items,
    siswaMap,
    existing
  );

  return {
    tahun_pelajaran: tahunTarget,
    items: validated
  };
}

function previewSimpanMembershipSiswa(tahunPelajaran, items) {
  try {
    const prepared = prepareMembershipValidation_(
      tahunPelajaran,
      items
    );

    return {
      success:true,
      mode:"PREVIEW",
      tahun_pelajaran:prepared.tahun_pelajaran,
      total:prepared.items.length,
      data:prepared.items
    };
  } catch (err) {
    return {
      success:false,
      message:err && err.message
        ? err.message
        : String(err)
    };
  }
}

function simpanMembershipSiswa(tahunPelajaran, items) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const prepared = prepareMembershipValidation_(
      tahunPelajaran,
      items
    );

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Siswa_Tahun_Pelajaran");

    if (!sheet) {
      throw new Error(
        "Sheet Siswa_Tahun_Pelajaran tidak ditemukan."
      );
    }

    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      .map(function(header) {
        return normalizeText(header);
      });
    const requiredHeaders = [
      "id_keanggotaan",
      "id_siswa",
      "tahun_pelajaran",
      "kelas",
      "status_siswa",
      "id_potongan_default",
      "tanggal_masuk",
      "tanggal_keluar",
      "keterangan"
    ];

    requiredHeaders.forEach(function(header) {
      if (headers.indexOf(header) === -1) {
        throw new Error(
          "Header wajib tidak ditemukan: " + header
        );
      }
    });

    const rows = prepared.items.map(function(item) {
      return headers.map(function(header) {
        return item[header] !== undefined
          ? item[header]
          : "";
      });
    });

    if (rows.length > 0) {
      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          rows.length,
          headers.length
        )
        .setValues(rows);
    }

    return {
      success:true,
      mode:"SIMPAN",
      tahun_pelajaran:prepared.tahun_pelajaran,
      total:rows.length,
      inserted:rows.length,
      message:
        rows.length +
        " membership siswa berhasil disimpan untuk tahun " +
        prepared.tahun_pelajaran + "."
    };

  } catch (err) {
    return {
      success:false,
      message:err && err.message
        ? err.message
        : String(err)
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Abaikan jika lock sudah dilepas.
    }
  }
}

// ============================================================
// END P1_MEMBERSHIP_BACKEND_CONTRACTS

// ============================================================
// P1 TEST - GET KANDIDAT MEMBERSHIP SUMMARY
// READ-ONLY
// ============================================================

// P1_TEST_GET_KANDIDAT_MEMBERSHIP_SUMMARY
function testP1GetKandidatMembershipSummary() {
  const tahunTarget = "2027/2028";

  const result = getKandidatSiswaMembership(tahunTarget);

  const summary = {
    success: result.success,
    tahun_pelajaran: result.tahun_pelajaran || tahunTarget,
    total: result.total || 0,
    belum_terdaftar: result.belum_terdaftar || 0,
    sudah_terdaftar: result.sudah_terdaftar || 0,
    message: result.message || ""
  };

  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

// ============================================================
// END P1 TEST - GET KANDIDAT MEMBERSHIP SUMMARY
// ============================================================

// ============================================================

// ============================================================

// ============================================================


// ============================================================


// ============================================================


// ============================================================
// P1 TEST 4 - DUPLICATE MEMBERSHIP PROTECTION
// ============================================================

// ============================================================
// P1 TEST - MIGRATION EXISTING MEMBERSHIP
// ============================================================

function testP1MigrationExistingMembership() {
  const tahunTarget = '2027/2028';

  const membershipList =
    getSheetDataAsObjects('Siswa_Tahun_Pelajaran') || [];

  const membershipTarget = membershipList.filter(function(item) {
    return normalizeText(item.tahun_pelajaran) === tahunTarget;
  });

  const expectedSkipped = membershipTarget.length;

  const result = migrasiSiswaKeMembership(tahunTarget);

  const actualInserted = Number(result && result.inserted || 0);
  const actualSkipped = Number(result && result.skipped || 0);
  const actualInvalid = Number(result && result.invalid || 0);

  const passed =
    !!result &&
    result.success === true &&
    actualInserted === 0 &&
    actualSkipped === expectedSkipped &&
    actualInvalid === 0;

  const output = {
    success: passed,
    test: 'P1 Migration Existing Membership',
    tahun_pelajaran: tahunTarget,
    expected: {
      inserted: 0,
      skipped: expectedSkipped,
      invalid: 0
    },
    actual: result
  };

  Logger.log(JSON.stringify(output, null, 2));
  return output;
}

// P1 TEST 4 - DUPLICATE MEMBERSHIP ASSERTION FIX
function testP1DuplicateMembership() {
  const tahunTarget = '2027/2028';

  const previewBefore = previewMigrasiSiswaKeMembership(tahunTarget);

  if (!previewBefore || previewBefore.success !== true) {
    throw new Error(
      'Test 4 FAIL: preview sebelum migrasi tidak berhasil.'
    );
  }

  const expectedExisting =
    Number(previewBefore.sudah_ada || 0);

  const migrationFirst =
    migrasiSiswaKeMembership(tahunTarget);

  const auditAfterFirst =
    auditSiswaMembership();

  const migrationSecond =
    migrasiSiswaKeMembership(tahunTarget);

  const auditAfterSecond =
    auditSiswaMembership();

  const checks = {
    preview_success:
      previewBefore.success === true,

    first_success:
      migrationFirst &&
      migrationFirst.success === true,

    second_success:
      migrationSecond &&
      migrationSecond.success === true,

    first_inserted_zero:
      Number(migrationFirst.inserted) === 0,

    second_inserted_zero:
      Number(migrationSecond.inserted) === 0,

    first_invalid_zero:
      Number(migrationFirst.invalid) === 0,

    second_invalid_zero:
      Number(migrationSecond.invalid) === 0,

    first_skipped_matches_existing:
      Number(migrationFirst.skipped) === expectedExisting,

    second_skipped_matches_existing:
      Number(migrationSecond.skipped) === expectedExisting,

    audit_first_no_duplicates:
      auditAfterFirst &&
      Number(auditAfterFirst.duplicate_id_siswa_tahun || 0) === 0,

    audit_second_no_duplicates:
      auditAfterSecond &&
      Number(auditAfterSecond.duplicate_id_siswa_tahun || 0) === 0,

    audit_first_no_orphans:
      auditAfterFirst &&
      Number(auditAfterFirst.orphan_membership_count || 0) === 0,

    audit_second_no_orphans:
      auditAfterSecond &&
      Number(auditAfterSecond.orphan_membership_count || 0) === 0,

    audit_first_no_missing_id:
      auditAfterFirst &&
      Number(auditAfterFirst.membership_tanpa_id_siswa || 0) === 0,

    audit_second_no_missing_id:
      auditAfterSecond &&
      Number(auditAfterSecond.membership_tanpa_id_siswa || 0) === 0,

    audit_first_no_missing_year:
      auditAfterFirst &&
      Number(auditAfterFirst.membership_tanpa_tahun_pelajaran || 0) === 0,

    audit_second_no_missing_year:
      auditAfterSecond &&
      Number(auditAfterSecond.membership_tanpa_tahun_pelajaran || 0) === 0
  };

  const failedChecks = Object.keys(checks)
    .filter(function(key) {
      return checks[key] !== true;
    });

  const success = failedChecks.length === 0;

  const result = {
    success: success,
    test: 'P1 Duplicate Membership',
    tahun_pelajaran: tahunTarget,
    expected_existing: expectedExisting,
    checks: checks,
    failed_checks: failedChecks,
    preview_before: previewBefore,
    migration_first: migrationFirst,
    migration_second: migrationSecond,
    audit_after_first: auditAfterFirst,
    audit_after_second: auditAfterSecond
  };

  Logger.log(JSON.stringify(result, null, 2));

  if (!success) {
    throw new Error(
      'Test 4 FAIL: ' + failedChecks.join(', ')
    );
  }

  return result;
}

// ============================================================
// END P1 TEST 4 - DUPLICATE MEMBERSHIP PROTECTION
// ============================================================
// ============================================================
// P1 TEST 4 SUMMARY - READ ONLY
// ============================================================

function testP1DuplicateMembershipSummary() {
  const tahunTarget = "2027/2028";

  const siswaList = getSheetDataAsObjects("Siswa");
  const membershipList = getSheetDataAsObjects("Siswa_Tahun_Pelajaran");

  const targetMembership = membershipList.filter(function(row) {
    return normalizeText(row.tahun_pelajaran) === tahunTarget;
  });

  const uniqueKeys = {};
  const duplicateKeys = {};
  const invalidRows = [];
  const masterIds = {};

  siswaList.forEach(function(siswa) {
    const idSiswa = normalizeText(siswa.id_siswa);
    if (idSiswa) masterIds[idSiswa] = true;
  });

  targetMembership.forEach(function(row, index) {
    const idSiswa = normalizeText(row.id_siswa);
    const tahun = normalizeText(row.tahun_pelajaran);
    const kelas = normalizeText(row.kelas);
    const key = idSiswa + "|" + tahun;

    if (!idSiswa || !tahun || !kelas) {
      invalidRows.push({
        row: index + 2,
        id_siswa: idSiswa,
        tahun_pelajaran: tahun,
        kelas: kelas
      });
    }

    if (uniqueKeys[key]) {
      duplicateKeys[key] = (duplicateKeys[key] || 1) + 1;
    } else {
      uniqueKeys[key] = true;
    }
  });

  const duplicateCount = Object.keys(duplicateKeys).length;

  const orphanCount = targetMembership.filter(function(row) {
    return !masterIds[normalizeText(row.id_siswa)];
  }).length;

  const audit = auditSiswaMembership();

  const result = {
    success: true,
    tahun_pelajaran: tahunTarget,
    total_siswa_master: siswaList.length,
    total_membership_target: targetMembership.length,
    unique_membership_keys_target: Object.keys(uniqueKeys).length,
    duplicate_key_count_target: duplicateCount,
    duplicate_keys_target: duplicateKeys,
    invalid_row_count_target: invalidRows.length,
    orphan_membership_count_target: orphanCount,
    expected_total_target: siswaList.length,
    count_matches_master: targetMembership.length === siswaList.length,
    audit_sehat: audit.sehat,
    audit_summary: {
      total_membership: audit.total_membership,
      unique_membership: audit.unique_membership,
      duplicate_id_siswa_tahun: audit.duplicate_id_siswa_tahun,
      membership_tanpa_id_siswa: audit.membership_tanpa_id_siswa,
      membership_tanpa_tahun_pelajaran: audit.membership_tanpa_tahun_pelajaran,
      orphan_membership: audit.orphan_membership
    }
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// ============================================================
// END P1 TEST 4 SUMMARY - READ ONLY
// ============================================================

// END P1_MEMBERSHIP_SCHEMA_AND_MIGRATION
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
    migrasi: [],
  };

  const migrasi = migrateDatabaseSchema();

  if (migrasi && migrasi.migrasi) {
    hasil.migrasi = migrasi.migrasi;
  }

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
