/**
 * MENU SETTING TAGIHAN
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. GENERATE TAGIHAN SPP BULANAN
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
      posSPP.tarif_default,
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
      const status = normalizeText(s.status_siswa).toUpperCase();

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
// 2. GENERATE TAGIHAN MASSAL (FLEXIBLE)
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

    let nominalAsal =
      nominalInput > 0 ? nominalInput : toNumber(pos.tarif_default);

    if (nominalAsal <= 0) {
      throw new Error("Nominal tagihan harus lebih dari 0.");
    }

    const sheetTagihan = getSheet("Tagihan_Siswa");
    const tagihanMap = getHeaderMap(sheetTagihan);
    const lastColumn = sheetTagihan.getLastColumn();

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

      let nominalPotongan = 0;
      const idPotongan =
        idPotonganInput === "NONE"
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

      const row = new Array(lastColumn).fill("");

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

      if (tagihanMap.status_siswa !== undefined) {
        row[tagihanMap.status_siswa] =
          normalizeText(siswa.status_siswa) || "AKTIF";
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
// 3. RIWAYAT GENERATE TAGIHAN
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

      let periodeTampil = "";
      const rawPeriode = t.periode;

      if (rawPeriode instanceof Date && !isNaN(rawPeriode.getTime())) {
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
          normalizeText(siswa.nama_lengkap) || "-",
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
        tanggal_tagihan: tanggalTampil,
        tanggal_ms: tanggalMs,
        bisa_dicicil: toBoolean(t.bisa_dicicil),
      };
    });

    semua.sort(function (a, b) {
      if (b.tanggal_ms !== a.tanggal_ms) {
        return b.tanggal_ms - a.tanggal_ms;
      }
      return String(b.id_tagihan).localeCompare(String(a.id_tagihan));
    });

    const total = semua.length;
    const start = (page - 1) * limit;
    const data = semua.slice(start, start + limit);

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
// 4. RIWAYAT SETTING KELAS
// ============================================================

function getRiwayatSettingKelas(page, perPage) {
  page = page || 1;
  perPage = perPage || 10;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

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

    const groupMap = {};

    for (let i = 1; i < dataTagihan.length; i++) {
      const row = dataTagihan[i];

      const idSiswa = String(row[col.idSiswa] || "").trim();
      if (!idSiswa) continue;

      const kelas = mapSiswaKelas[idSiswa] || "Tidak diketahui";

      // Ambil id_pos dari data tagihan
      let idPos = "";
      if (col.idPos !== -1 && row[col.idPos]) {
        idPos = String(row[col.idPos]).trim();
      }

      let posNama = "-";
      if (col.namaItem !== -1 && row[col.namaItem]) {
        posNama = String(row[col.namaItem]).trim();
      }

      // Jika idPos kosong, coba dapatkan dari Master_Pos berdasarkan nama
      if (!idPos && posNama !== "-") {
        const posList = getSheetDataAsObjects("Master_Pos");
        const foundPos = posList.find(function(p) {
          return normalizeText(p.nama_pos) === normalizeText(posNama);
        });
        if (foundPos) {
          idPos = normalizeText(foundPos.id_pos);
        }
      }

      const periode =
        col.periode !== -1 ? String(row[col.periode] || "").trim() : "-";

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
          id_pos: idPos, // Tambahkan id_pos
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

// ============================================================
// 5. HAPUS TAGIHAN SISWA
// ============================================================

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
      const status =
        statusCol === -1
          ? ""
          : normalizeText(values[i][statusCol]).toUpperCase();
      if (terbayar > 0 || status === "LUNAS" || status === "CICILAN") {
        throw new Error(
          "Tagihan yang sudah memiliki pembayaran tidak dapat dihapus.",
        );
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

// ============================================================
// 6. HAPUS SETTING KELAS (DENGAN PENGAMAN PEMBAYARAN)
// ============================================================

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
      terbayar: findCol(headersT, ["terbayar"]),
      status: findCol(headersT, ["status"]),
      nominalAkhir: findCol(headersT, ["nominal_akhir"]),
      sisa: findCol(headersT, ["sisa_tunggakan"]),
    };

    if (col.idSiswa === -1 || col.namaItem === -1) {
      return {
        success: false,
        message: "Kolom id_siswa / nama_item_snapshot tidak ditemukan",
        deleted: 0,
      };
    }

    // Map siswa ke kelas
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

    const rowsToDelete = [];
    var skippedPaid = 0;
    var totalMatched = 0;

    // ============================================================
    // PENGAMAN: CEK APAKAH ADA YANG SUDAH DIBAYAR
    // ============================================================
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

      totalMatched++;

      // ============================================================
      // PENGAMAN: CEK APAKAH SUDAH ADA PEMBAYARAN
      // ============================================================
      var terbayar = 0;
      if (col.terbayar !== -1) {
        terbayar = Number(row[col.terbayar]) || 0;
      }

      var status = "";
      if (col.status !== -1) {
        status = String(row[col.status] || "").toLowerCase();
      }

      // Jika sudah dibayar (terbayar > 0 atau status lunas/cicilan)
      if (terbayar > 0 || 
          status.indexOf("lunas") !== -1 || 
          status.indexOf("cicil") !== -1 ||
          status.indexOf("sebagian") !== -1) {
        skippedPaid++;
        continue;
      }

      rowsToDelete.push(i + 1);
    }

    // Jika ada yang sudah dibayar, tolak penghapusan
    if (skippedPaid > 0) {
      return {
        success: false,
        message: "Tidak dapat menghapus setting tagihan ini karena terdapat " + skippedPaid + " tagihan yang sudah dibayar (LUNAS/CICILAN).",
        deleted: 0,
        skipped: skippedPaid,
        totalMatched: totalMatched
      };
    }

    // Jika tidak ada yang bisa dihapus
    if (rowsToDelete.length === 0) {
      return {
        success: true,
        deleted: 0,
        message: "Tidak ada tagihan yang bisa dihapus (mungkin sudah lunas atau tidak ditemukan)",
        skipped: skippedPaid,
        totalMatched: totalMatched
      };
    }

    // Hapus dari bawah ke atas
    rowsToDelete.sort(function (a, b) {
      return b - a;
    });

    for (var j = 0; j < rowsToDelete.length; j++) {
      sheetTagihan.deleteRow(rowsToDelete[j]);
    }

    return {
      success: true,
      deleted: rowsToDelete.length,
      message: "Berhasil menghapus " + rowsToDelete.length + " tagihan.",
      skipped: skippedPaid,
      totalMatched: totalMatched
    };
  } catch (err) {
    return {
      success: false,
      deleted: 0,
      message: err.message
    };
  }
}

// ============================================================
// 7. HAPUS SETTING KELAS BATCH
// ============================================================

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

    function isSudahDibayar(row) {
      if (col.terbayar !== -1) {
        var bayar = Number(row[col.terbayar]) || 0;
        if (bayar > 0) return true;
      }
      if (col.status !== -1) {
        var st = String(row[col.status] || "").toLowerCase();
        if (st.indexOf("lunas") !== -1) return true;
        if (st.indexOf("sebagian") !== -1) return true;
        if (st.indexOf("partial") !== -1) return true;
        if (st.indexOf("cicil") !== -1) return true;
      }
      return false;
    }

    const deletableRows = [];
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

      if (isSudahDibayar(row)) {
        skippedPaid++;
        continue;
      }

      deletableRows.push(i + 1);
    }

    const totalDeletable = deletableRows.length;

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


// ============================================================
// 8. FITUR PENCARIAN SETTING TAGIHAN
// ============================================================

/**
 * PENCARIAN TAGIHAN BERDASARKAN KELAS
 * @param {string} keyword - Kata kunci pencarian (kelas / pos / periode)
 * @param {string} filterPos - Filter berdasarkan pos pembayaran (opsional)
 * @param {number} page - Halaman
 * @param {number} limit - Jumlah data per halaman
 * @returns {object} Hasil pencarian dengan paginasi
 */
/**
 * PENCARIAN TAGIHAN BERDASARKAN KELAS
 * @param {string} keyword - Kata kunci pencarian (kelas / pos / periode)
 * @param {string} filterPos - Filter berdasarkan pos pembayaran (opsional)
 * @param {number} page - Halaman
 * @param {number} limit - Jumlah data per halaman
 * @returns {object} Hasil pencarian dengan paginasi
 */
function cariSettingTagihanByKelas(keyword, filterPos, page, limit) {
  try {
    keyword = normalizeText(keyword).toLowerCase();
    filterPos = normalizeText(filterPos);
    page = Math.max(1, parseInt(page, 10) || 1);
    limit = Math.min(Math.max(parseInt(limit, 10) || 10, 5), 50);

    // Ambil data riwayat setting kelas
    const result = getRiwayatSettingKelas(1, 9999);
    
    if (result.error) {
      return {
        success: false,
        message: result.error,
        data: [],
        total: 0,
        page: page,
        limit: limit,
        total_pages: 0
      };
    }

    let data = result.data || [];

    // Filter berdasarkan keyword
    if (keyword) {
      data = data.filter(function (item) {
        const searchable = [
          item.kelas,
          item.pos,
          item.periode,
          item.status,
          item.tanggal
        ].join(" ").toLowerCase();
        
        return searchable.indexOf(keyword) !== -1;
      });
    }

    // Filter berdasarkan pos pembayaran (gunakan id_pos)
    if (filterPos && filterPos !== "ALL" && filterPos !== "") {
      data = data.filter(function (item) {
        // Coba filter berdasarkan id_pos jika ada, atau nama pos
        if (item.id_pos) {
          return normalizeText(item.id_pos) === filterPos;
        }
        return normalizeText(item.pos).toLowerCase() === filterPos.toLowerCase();
      });
    }

    // Urutkan berdasarkan tanggal terbaru
    data.sort(function (a, b) {
      if (a.tanggalBatch && b.tanggalBatch) {
        return b.tanggalBatch.localeCompare(a.tanggalBatch);
      }
      return 0;
    });

    const total = data.length;
    const start = (page - 1) * limit;
    const pagedData = data.slice(start, start + limit);

    return {
      success: true,
      data: pagedData,
      total: total,
      page: page,
      limit: limit,
      total_pages: Math.ceil(total / limit) || 1
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
      data: [],
      total: 0,
      page: 1,
      limit: 10,
      total_pages: 0
    };
  }
}

/**
 * PENCARIAN TAGIHAN BERDASARKAN SISWA
 * @param {string} keyword - Kata kunci pencarian (nama siswa / nisn / nis)
 * @param {string} filterPos - Filter berdasarkan pos pembayaran (opsional)
 * @param {string} filterKelas - Filter berdasarkan kelas (opsional)
 * @param {string} filterStatus - Filter status tagihan (opsional)
 * @param {number} page - Halaman
 * @param {number} limit - Jumlah data per halaman
 * @returns {object} Hasil pencarian dengan paginasi
 */
function cariSettingTagihanBySiswa(keyword, filterPos, filterKelas, filterStatus, page, limit) {
  try {
    keyword = normalizeText(keyword).toLowerCase();
    filterPos = normalizeText(filterPos);
    filterKelas = normalizeText(filterKelas);
    filterStatus = normalizeText(filterStatus).toUpperCase();
    page = Math.max(1, parseInt(page, 10) || 1);
    limit = Math.min(Math.max(parseInt(limit, 10) || 10, 5), 50);

    // Ambil semua tagihan
    const tagihanList = getSheetDataAsObjects("Tagihan_Siswa");
    const siswaList = getSheetDataAsObjects("Siswa");
    const posList = getSheetDataAsObjects("Master_Pos");

    // Buat map untuk akses cepat
    const mapSiswa = {};
    siswaList.forEach(function (s) {
      mapSiswa[normalizeText(s.id_siswa)] = s;
    });

    const mapPos = {};
    posList.forEach(function (p) {
      mapPos[normalizeText(p.id_pos)] = p;
    });

    const timezone = Session.getScriptTimeZone() || "Asia/Jakarta";

    // Proses data tagihan
    let data = tagihanList.map(function (t) {
      const idSiswa = normalizeText(t.id_siswa);
      const siswa = mapSiswa[idSiswa] || {};
      const pos = mapPos[normalizeText(t.id_pos)] || {};

      // Tanggal tagihan
      let tanggalTampil = "-";
      const rawTgl = t.tanggal_tagihan;
      if (rawTgl instanceof Date && !isNaN(rawTgl.getTime())) {
        tanggalTampil = Utilities.formatDate(rawTgl, timezone, "dd/MM/yyyy");
      } else if (rawTgl) {
        const d = new Date(rawTgl);
        if (!isNaN(d.getTime())) {
          tanggalTampil = Utilities.formatDate(d, timezone, "dd/MM/yyyy");
        }
      }

      // Periode
      let periodeTampil = "";
      const rawPeriode = t.periode;
      if (rawPeriode instanceof Date && !isNaN(rawPeriode.getTime())) {
        periodeTampil = Utilities.formatDate(rawPeriode, timezone, "yyyy-MM");
      } else {
        periodeTampil = normalizeText(rawPeriode);
      }

      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const sisa = Math.max(0, toNumber(t.sisa_tunggakan) || nominalAkhir - terbayar);
      
      let status = normalizeText(t.status) || "BELUM_BAYAR";
      if (sisa <= 0) {
        status = "LUNAS";
      } else if (terbayar > 0) {
        status = "CICILAN";
      }

      const namaSiswa = normalizeText(siswa.nama_lengkap) || "-";
      const kelas = normalizeText(siswa.kelas) || "-";

      return {
        id_tagihan: normalizeText(t.id_tagihan),
        id_siswa: idSiswa,
        nama_siswa: namaSiswa,
        nisn: normalizeText(siswa.nisn) || "-",
        nis: normalizeText(siswa.nis) || "-",
        kelas: kelas,
        id_pos: normalizeText(t.id_pos),
        nama_pos: normalizeText(pos.nama_pos) || normalizeText(t.nama_item_snapshot) || "-",
        periode: periodeTampil || "-",
        nominal_akhir: nominalAkhir,
        terbayar: terbayar,
        sisa_tunggakan: sisa,
        status: status,
        tanggal_tagihan: tanggalTampil,
        bisa_dicicil: toBoolean(t.bisa_dicicil)
      };
    });

    // Filter berdasarkan keyword (nama siswa, nisn, nis, id tagihan)
    if (keyword) {
      data = data.filter(function (item) {
        const searchable = [
          item.nama_siswa,
          item.nisn,
          item.nis,
          item.id_tagihan,
          item.nama_pos,
          item.periode
        ].join(" ").toLowerCase();
        
        return searchable.indexOf(keyword) !== -1;
      });
    }

    // Filter berdasarkan pos pembayaran
    if (filterPos && filterPos !== "ALL" && filterPos !== "") {
      data = data.filter(function (item) {
        return normalizeText(item.id_pos) === filterPos;
      });
    }

    // Filter berdasarkan kelas
    if (filterKelas && filterKelas !== "ALL" && filterKelas !== "") {
      data = data.filter(function (item) {
        return normalizeText(item.kelas) === filterKelas;
      });
    }

    // Filter berdasarkan status
    if (filterStatus && filterStatus !== "ALL" && filterStatus !== "") {
      data = data.filter(function (item) {
        return normalizeText(item.status) === filterStatus;
      });
    }

    // Urutkan berdasarkan nama siswa
    data.sort(function (a, b) {
      return a.nama_siswa.localeCompare(b.nama_siswa);
    });

    const total = data.length;
    const start = (page - 1) * limit;
    const pagedData = data.slice(start, start + limit);

    return {
      success: true,
      data: pagedData,
      total: total,
      page: page,
      limit: limit,
      total_pages: Math.ceil(total / limit) || 1
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
      data: [],
      total: 0,
      page: 1,
      limit: 10,
      total_pages: 0
    };
  }
}

/**
 * GET DAFTAR SISWA UNTUK FILTER DI SETTING TAGIHAN
 * @param {string} keyword - Kata kunci pencarian
 * @param {string} filterKelas - Filter kelas (opsional)
 * @returns {array} Daftar siswa
 */
function getSiswaForTagihanFilter(keyword, filterKelas) {
  try {
    keyword = normalizeText(keyword).toLowerCase();
    filterKelas = normalizeText(filterKelas);

    let siswaList = getSheetDataAsObjects("Siswa");
    const potonganList = getSheetDataAsObjects("Master_Potongan");

    // Filter siswa aktif
    siswaList = siswaList.filter(function (s) {
      const status = normalizeText(s.status_siswa).toUpperCase();
      return status === "AKTIF";
    });

    // Filter keyword
    if (keyword) {
      siswaList = siswaList.filter(function (s) {
        const searchable = [
          normalizeText(s.nama_lengkap),
          normalizeText(s.nisn),
          normalizeText(s.nis),
          normalizeText(s.kelas)
        ].join(" ").toLowerCase();
        
        return searchable.indexOf(keyword) !== -1;
      });
    }

    // Filter kelas
    if (filterKelas && filterKelas !== "ALL" && filterKelas !== "") {
      siswaList = siswaList.filter(function (s) {
        return normalizeText(s.kelas) === filterKelas;
      });
    }

    // Map potongan
    const mapPotongan = {};
    potonganList.forEach(function (p) {
      mapPotongan[normalizeText(p.id_potongan)] = p;
    });

    // Format output
    return siswaList.map(function (s) {
      const idPotongan = normalizeText(s.id_potongan_default);
      const potongan = mapPotongan[idPotongan];

      return {
        id_siswa: normalizeText(s.id_siswa),
        nisn: normalizeText(s.nisn),
        nis: normalizeText(s.nis),
        nama_lengkap: normalizeText(s.nama_lengkap),
        kelas: normalizeText(s.kelas),
        tahun_pelajaran: normalizeText(s.tahun_pelajaran),
        id_potongan_default: idPotongan,
        nama_potongan: potongan ? normalizeText(potongan.nama_potongan) : "Tidak Ada"
      };
    }).sort(function (a, b) {
      return a.nama_lengkap.localeCompare(b.nama_lengkap);
    });
  } catch (err) {
    return [];
  }
}

/**
 * GET STATISTIK CEPAT UNTUK SETTING TAGIHAN
 * @returns {object} Statistik tagihan
 */
function getTagihanStatistics() {
  try {
    const tagihanList = getSheetDataAsObjects("Tagihan_Siswa");
    const siswaList = getSheetDataAsObjects("Siswa");

    // Filter siswa aktif
    const siswaAktifIds = new Set();
    siswaList.forEach(function (s) {
      const status = normalizeText(s.status_siswa).toUpperCase();
      if (status === "AKTIF") {
        siswaAktifIds.add(normalizeText(s.id_siswa));
      }
    });

    let totalTagihan = 0;
    let totalBelumBayar = 0;
    let totalCicilan = 0;
    let totalLunas = 0;
    let totalSisaTunggakan = 0;
    let siswaDenganTunggakan = new Set();

    tagihanList.forEach(function (t) {
      const idSiswa = normalizeText(t.id_siswa);
      
      // Hanya hitung untuk siswa aktif
      if (!siswaAktifIds.has(idSiswa)) return;

      totalTagihan++;
      
      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const sisa = Math.max(0, nominalAkhir - terbayar);

      let status = normalizeText(t.status) || "BELUM_BAYAR";
      if (sisa <= 0) {
        status = "LUNAS";
        totalLunas++;
      } else if (terbayar > 0) {
        status = "CICILAN";
        totalCicilan++;
        totalSisaTunggakan += sisa;
        siswaDenganTunggakan.add(idSiswa);
      } else {
        status = "BELUM_BAYAR";
        totalBelumBayar++;
        totalSisaTunggakan += sisa;
        siswaDenganTunggakan.add(idSiswa);
      }
    });

    return {
      success: true,
      total_tagihan: totalTagihan,
      total_belum_bayar: totalBelumBayar,
      total_cicilan: totalCicilan,
      total_lunas: totalLunas,
      total_sisa_tunggakan: totalSisaTunggakan,
      total_siswa_menunggak: siswaDenganTunggakan.size,
      total_siswa_aktif: siswaAktifIds.size
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}


/**
 * EXPORT HASIL PENCARIAN TAGIHAN KE EXCEL (DOWNLOAD LANGSUNG)
 * @param {string} keyword - Kata kunci pencarian
 * @param {string} filterPos - Filter pos
 * @param {string} filterKelas - Filter kelas
 * @param {string} filterStatus - Filter status
 * @param {string} searchType - Tipe pencarian ('kelas' atau 'siswa')
 * @returns {object} Hasil export dengan URL download
 */
function exportTagihanSearchResult(keyword, filterPos, filterKelas, filterStatus, searchType) {
  try {
    var data = [];
    var fileName = "";

    if (searchType === "kelas") {
      var result = cariSettingTagihanByKelas(keyword, filterPos, 1, 9999);
      if (result.success) {
        data = result.data;
        fileName = "Setting_Tagihan_Kelas_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
      }
    } else {
      var result = cariSettingTagihanBySiswa(keyword, filterPos, filterKelas, filterStatus, 1, 9999);
      if (result.success) {
        data = result.data;
        fileName = "Setting_Tagihan_Siswa_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
      }
    }

    if (!data || data.length === 0) {
      throw new Error("Tidak ada data untuk diekspor.");
    }

    // Buat spreadsheet temporer
    var tempSS = SpreadsheetApp.create(fileName);
    var sheet = tempSS.getActiveSheet();
    sheet.setName("Data");

    // Tentukan header berdasarkan tipe
    var headers;
    if (searchType === "kelas") {
      headers = ["Kelas", "Pos", "Periode", "Tanggal", "Jumlah Siswa", "Total Tagihan", "Status"];
    } else {
      headers = ["ID Tagihan", "Nama Siswa", "NISN", "Kelas", "Pos", "Periode", "Total Tagihan", "Sudah Bayar", "Sisa", "Status", "Tanggal Tagihan"];
    }

    // Write header
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#d1fae5");
    headerRange.setHorizontalAlignment("center");

    // Write data
    var rows = [];
    data.forEach(function (item) {
      if (searchType === "kelas") {
        rows.push([
          item.kelas || "-",
          item.pos || "-",
          item.periode || "-",
          item.tanggal || "-",
          item.jumlahSiswa || 0,
          toNumber(item.totalTagihan) || 0,
          item.status || "-"
        ]);
      } else {
        rows.push([
          item.id_tagihan || "-",
          item.nama_siswa || "-",
          item.nisn || "-",
          item.kelas || "-",
          item.nama_pos || "-",
          item.periode || "-",
          toNumber(item.nominal_akhir) || 0,
          toNumber(item.terbayar) || 0,
          toNumber(item.sisa_tunggakan) || 0,
          item.status || "-",
          item.tanggal_tagihan || "-"
        ]);
      }
    });

    if (rows.length > 0) {
      var dataRange = sheet.getRange(2, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Format kolom nominal
      if (searchType === "kelas") {
        if (headers.length >= 6) {
          sheet.getRange(2, 6, rows.length, 1).setNumberFormat('"Rp" #,##0');
        }
      } else {
        if (headers.length >= 9) {
          sheet.getRange(2, 7, rows.length, 3).setNumberFormat('"Rp" #,##0');
        }
      }
    }

    // Info footer
    var infoRow = rows.length + 3;
    var infoRange = sheet.getRange(infoRow, 1, 1, headers.length);
    infoRange.merge();
    infoRange.setValue("Total Data: " + data.length + " | Tanggal Export: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));
    infoRange.setFontSize(9);
    infoRange.setFontStyle("italic");

    sheet.autoResizeColumns(1, headers.length);

    // Flush untuk memastikan semua data tersimpan
    SpreadsheetApp.flush();

    // Dapatkan ID file dan buat URL download
    var fileId = tempSS.getId();
    var downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

    // Hapus file temporer setelah beberapa saat (opsional, bisa dihapus nanti)
    // Untuk sekarang, kita kembalikan URL download

    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: fileName + ".xlsx",
      message: "File Excel siap diunduh (" + data.length + " data).",
      totalData: data.length
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}

/**
 * GET DAFTAR POS UNTUK FILTER SETTING TAGIHAN
 * @returns {array} Daftar pos aktif
 */
function getPosForTagihanFilter() {
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


