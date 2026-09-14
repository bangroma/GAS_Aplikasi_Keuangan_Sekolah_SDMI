from pathlib import Path
from datetime import datetime
import shutil

TARGET = Path("00_Config.js")
MARKER = "// P1_MEMBERSHIP_SCHEMA_AND_MIGRATION"

if not TARGET.exists():
    raise SystemExit("ERROR: 00_Config.js tidak ditemukan.")

original = TARGET.read_bytes()

if MARKER.encode("utf-8") in original:
    raise SystemExit("ABORT: marker P1 sudah ada. Installer tidak dijalankan ulang.")

if b"\r\n" in original:
    eol = "\r\n"
    text = original.decode("utf-8").replace("\r\n", "\n")
else:
    eol = "\n"
    text = original.decode("utf-8")

schema_anchor = "  Master_Pos: ["
if text.count(schema_anchor) != 1:
    raise SystemExit(
        "ERROR: anchor Master_Pos tidak ditemukan tepat satu kali."
    )

membership_schema = '''  // P1: membership siswa per tahun pelajaran
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

'''

text = text.replace(schema_anchor, membership_schema + schema_anchor, 1)

setup_anchor = "// 4. SETUP DATABASE"
if text.count(setup_anchor) != 1:
    raise SystemExit(
        "ERROR: anchor SETUP DATABASE tidak ditemukan tepat satu kali."
    )

functions = r'''// ============================================================
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

  const siswaList = getSheetDataAsObjects("Siswa").filter(function (siswa) {
    return normalizeText(siswa.tahun_pelajaran) === tahunTarget;
  });

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
    const tahun = normalizeText(siswa.tahun_pelajaran);
    const kelas = normalizeText(siswa.kelas);

    if (!idSiswa || !tahun || !kelas) {
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

    const siswaList = getSheetDataAsObjects("Siswa").filter(function (siswa) {
      return normalizeText(siswa.tahun_pelajaran) === tahunTarget;
    });

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
      const tahun = normalizeText(siswa.tahun_pelajaran);
      const kelas = normalizeText(siswa.kelas);

      if (!idSiswa || !tahun || !kelas) {
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
// END P1_MEMBERSHIP_SCHEMA_AND_MIGRATION
// ============================================================

'''

text = text.replace(setup_anchor, functions + setup_anchor, 1)

backup = TARGET.with_name(
    TARGET.name
    + ".P1_BACKUP_"
    + datetime.now().strftime("%Y%m%d_%H%M%S")
)

shutil.copy2(TARGET, backup)
TARGET.write_bytes(text.replace("\n", eol).encode("utf-8"))

print("OK: P1 berhasil dipasang.")
print("Target :", TARGET)
print("Backup :", backup)
print("Marker :", MARKER)
