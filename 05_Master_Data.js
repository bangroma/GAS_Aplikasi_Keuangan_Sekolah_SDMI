/**
 * MENU MASTER DATA
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. MASTER DATA CONFIG
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

// ============================================================
// 2. GET MASTER DATA
// ============================================================

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
    const obj = { _rowIndex: i + 1 };

    headers.forEach(function (header, idx) {
      let value = row[idx];

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

// ============================================================
// 3. SAVE MASTER DATA
// ============================================================

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

    const pkValue = normalizeText(payload[primaryKey]);
    if (!pkValue) {
      throw new Error("Primary key (" + primaryKey + ") wajib diisi.");
    }

    const existing = getSheetDataAsObjects(sheetName);
    const duplicate = existing.find(function (row, idx) {
      const existingPk = normalizeText(row[primaryKey]);
      const existingRowIndex = idx + 2;
      return existingPk === pkValue && existingRowIndex !== rowIndex;
    });

    if (duplicate) {
      throw new Error(primaryKey + ' "' + pkValue + '" sudah digunakan.');
    }

    const rowValues = new Array(lastColumn).fill("");

    headers.forEach(function (header) {
      if (header === "_rowIndex") return;

      let value = payload[header];

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

// ============================================================
// 4. DELETE MASTER DATA
// ============================================================

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