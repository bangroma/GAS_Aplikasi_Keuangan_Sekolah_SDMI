/**
 * MENU DATA SISWA
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. GET SISWA BY KELAS
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
      const status = normalizeText(s.status_siswa).toUpperCase();
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
        nama_lengkap: normalizeText(s.nama_lengkap),
        jenis_kelamin: normalizeText(s.jenis_kelamin), // <-- Tambahkan ini
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
// 2. GET LIST KELAS
// ============================================================

function getListKelas() {
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
// 3. GET SISWA LIST
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
      _rowIndex: idx + 2,
      id_siswa: normalizeText(s.id_siswa),
      nisn: normalizeText(s.nisn),
      nis: normalizeText(s.nis),
      nama_lengkap: normalizeText(s.nama_lengkap),
      jenis_kelamin: normalizeText(s.jenis_kelamin),
      kelas: normalizeText(s.kelas),
      tahun_pelajaran: normalizeText(s.tahun_pelajaran),
      status_siswa: normalizeText(s.status_siswa) || "AKTIF",
      id_potongan_default: idPotongan,
      nama_potongan: potongan
        ? normalizeText(potongan.nama_potongan)
        : "Tidak Ada",
    };
  });

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

// ============================================================
// 4. SAVE SISWA
// ============================================================

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

    // Handle status
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

// ============================================================
// 5. DELETE SISWA
// ============================================================

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

// ============================================================
// 6. EXPORT SISWA TO EXCEL
// ============================================================

function exportSiswaToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getSiswaList().data || [];

  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    Session.getScriptTimeZone() || "Asia/Jakarta",
    "yyyyMMdd_HHmmss",
  );

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

  const fileId = tempSS.getId();
  const url =
    "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

  return {
    success: true,
    downloadUrl: url,
    fileName: "Data_Siswa_" + timestamp + ".xlsx",
    message: "File Excel siap diunduh (" + data.length + " data).",
  };
}

// ============================================================
// 7. IMPORT SISWA
// ============================================================

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

      if (
        headerMap.status !== undefined &&
        headerMap.status_siswa === undefined
      ) {
        rowValues[headerMap.status] = item.status_siswa || "AKTIF";
      }

      if (mapExisting[idSiswa]) {
        const rowIndex = mapExisting[idSiswa];
        sheet.getRange(rowIndex, 1, 1, lastColumn).setValues([rowValues]);
        updated++;
      } else {
        const newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1, 1, lastColumn).setValues([rowValues]);
        mapExisting[idSiswa] = newRow;
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