/**
 * CETAK NOTA / KUITANSI
 * Sistem Keuangan Sekolah
 */

// ============================================================
// 1. SETUP CETAK NOTA
// ============================================================

function setupCetakNota() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Cetak_Nota";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = [
    [
      "Waktu Cetak",
      "No. Kuitansi",
      "NISN / NIS",
      "Nama Siswa",
      "Kelas",
      "Rincian Item",
      "Total Bayar",
      "Metode Pembayaran",
      "Kasir",
      "Catatan",
      "Link File PDF",
    ],
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers[0].length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#059669");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 35);

  sheet.setFrozenRows(1);

  sheet.getRange("G2:G").setNumberFormat("Rp#,##0");

  sheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");

  for (let i = 1; i <= headers[0].length; i++) {
    sheet.autoResizeColumn(i);
  }

  Logger.log("Sheet 'Cetak Nota' berhasil dibuat dan diformat.");
  return "Berhasil menyiapkan Sheet 'Cetak_Nota'!";
}

// ============================================================
// 2. SIMPAN LOG CETAK NOTA
// ============================================================

function simpanLogCetakNota(data) {
  setupCetakNota();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("Cetak Nota");
  if (!sheet) {
    sheet = ss.getSheetByName("Cetak_Nota");
  }

  const noKuitansiTarget = String(data.noKuitansi).trim();

  // CEK DUPLIKASI
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    for (let i = 0; i < dataRange.length; i++) {
      const existingNoKuitansi = String(dataRange[i][1]).trim();
      const existingPdfUrl = dataRange[i][10];

      if (existingNoKuitansi === noKuitansiTarget && noKuitansiTarget !== "" && noKuitansiTarget !== "-") {
        return {
          success: true,
          pdfUrl: existingPdfUrl,
          isDuplicate: true,
        };
      }
    }
  }

  // FOLDER TARGET
  const folderId = "1BSKtVP4G6UbpSJqoRVjl3gzRbtIcJwFV";
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }

  let tanggalFormatted = data.tanggal || "-";

  let rincianRows = "";
  let rincianTeksArr = [];

  if (data.items && Array.isArray(data.items)) {
    data.items.forEach(function (item) {
      const nomFormatted = "Rp " + Number(item.nominal_dibayar).toLocaleString("id-ID");
      const diskonText = item.diskon_tambahan > 0
        ? `<div style="font-size: 7.5px; color: #059669; font-weight: normal;">Diskon: -Rp ${Number(item.diskon_tambahan).toLocaleString("id-ID")}</div>`
        : "";

      rincianRows += `
        <tr>
          <td style="font-weight: bold; vertical-align: top; padding: 2px 2px 2px 0; color: #0f172a;">
            ${item.nama_item}
            ${diskonText}
          </td>
          <td style="font-weight: bold; text-align: right; vertical-align: top; padding: 2px 0 2px 2px; white-space: nowrap; color: #0f172a;">
            ${nomFormatted}
          </td>
        </tr>
      `;

      rincianTeksArr.push(`${item.nama_item} (${nomFormatted})`);
    });
  }

  const rincianTeks = rincianTeksArr.join(", ");

  // TEMPLATE PDF
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page {
          size: 58mm auto;
          margin: 0mm;
        }
        * { box-sizing: border-box; }
        html, body {
          width: 58mm;
          margin: 0 auto;
          padding: 1mm;
          background-color: #ffffff;
          font-family: 'Courier New', Courier, monospace;
          font-size: 8.5px;
          line-height: 1.25;
          color: #1e293b;
        }
        .receipt-box {
          width: 100%;
          border: 1px dashed #94a3b8;
          border-radius: 6px;
          padding: 6px 5px;
          background: #ffffff;
        }
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .bold { font-weight: bold; color: #0f172a; }
        .muted { color: #64748b; font-weight: normal; }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 0;
        }
        td {
          vertical-align: top;
          padding: 1px 0;
          font-size: 8.5px;
        }
        .dashed-line {
          border-bottom: 1px dashed #94a3b8;
          margin: 5px 0;
        }
        .solid-title {
          border-bottom: 1px solid #cbd5e1;
          padding-bottom: 2px;
          margin-bottom: 3px;
        }
      </style>
    </head>
    <body>
      <div class="receipt-box">
        <div class="text-center">
          <div class="bold" style="font-size: 9.5px; text-transform: uppercase;">MI MUHAMMADIYAH KERTONATAN</div>
          <div class="muted" style="font-size: 8px; margin-top: 1px;">Bukti Pembayaran Keuangan</div>
          <div class="muted" style="font-size: 8px;">Tahun Pelajaran 2026/2027</div>
        </div>

        <div class="dashed-line"></div>

        <div>
          <table>
            <tr>
              <td class="muted" style="width: 40%;">No. Kuitansi:</td>
              <td class="bold text-right" style="width: 60%;">${data.noKuitansi}</td>
            </tr>
            <tr>
              <td class="muted">Tanggal:</td>
              <td class="text-right" style="word-break: break-all;">${tanggalFormatted}</td>
            </tr>
            <tr>
              <td class="muted">Kasir:</td>
              <td class="text-right">${data.kasir}</td>
            </tr>
          </table>
        </div>

        <div class="dashed-line"></div>

        <div>
          <table>
            <tr>
              <td class="muted" style="width: 30%;">Siswa:</td>
              <td class="bold text-right" style="width: 70%;">${data.namaSiswa}</td>
            </tr>
            <tr>
              <td class="muted">NISN /<br>Kelas:</td>
              <td class="text-right">${data.nisn} / Kelas<br>${data.kelas}</td>
            </tr>
          </table>
        </div>

        <div class="dashed-line"></div>

        <div>
          <div class="muted solid-title" style="font-size: 8px; text-transform: uppercase;">
            DETAIL PEMBAYARAN
          </div>
          <table>
            ${rincianRows}
          </table>
        </div>

        <div class="dashed-line"></div>

        <div>
          <table>
            <tr>
              <td class="bold" style="font-size: 9.5px; padding-bottom: 2px;">TOTAL BAYAR:</td>
              <td class="bold text-right" style="font-size: 9.5px; padding-bottom: 2px;">Rp ${Number(data.total).toLocaleString('id-ID')}</td>
            </tr>
            <tr>
              <td class="muted">Metode Bayar:</td>
              <td class="bold text-right">${data.metode}</td>
            </tr>
            ${data.catatan ? `
            <tr>
              <td colspan="2" class="muted" style="font-style: italic; font-size: 7.5px; padding-top: 3px;">
                Catatan: ${data.catatan}
              </td>
            </tr>` : ''}
          </table>
        </div>

        <div class="dashed-line"></div>

        <div class="text-center" style="padding-top: 1px;">
          <div class="bold" style="font-size: 8.5px; margin-bottom: 1px;">*** TERIMA KASIH ***</div>
          <div class="muted" style="font-size: 7.5px; line-height: 1.1;">Simpan kuitansi ini sebagai bukti pembayaran sah.</div>
        </div>
      </div>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent);
  const pdfBlob = htmlOutput.getAs('application/pdf').setName(`Nota_${data.noKuitansi}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = pdfFile.getUrl();

  // CATAT KE SHEET
  sheet.appendRow([
    new Date(),
    data.noKuitansi,
    data.nisn || "-",
    data.namaSiswa,
    data.kelas || "-",
    rincianTeks || "-",
    Number(data.total),
    data.metode,
    data.kasir,
    data.catatan || "-",
    pdfUrl,
  ]);

  return {
    success: true,
    pdfUrl: pdfUrl,
    isDuplicate: false,
  };
}
/* ============================================================
   3. AMBIL NOTA UNTUK CETAK ULANG
============================================================ */

function ambilNotaCetakUlang(noKuitansi) {
  const target = String(noKuitansi || "").trim();

  if (!target || target === "-") {
    throw new Error("Nomor kuitansi tidak valid.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("Cetak_Nota") ||
    ss.getSheetByName("Cetak Nota");

  if (!sheet) {
    throw new Error("Sheet 'Cetak_Nota' tidak ditemukan.");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("Belum ada arsip nota pada Sheet 'Cetak_Nota'.");
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  for (let i = 0; i < values.length; i++) {
    const existingNoKuitansi = String(values[i][1] || "").trim();

    if (existingNoKuitansi === target) {
      const pdfUrl = String(values[i][10] || "").trim();

      if (!pdfUrl) {
        throw new Error(
          "Nota ditemukan, tetapi Link File PDF belum tersedia."
        );
      }

      return {
        success: true,
        no_kuitansi: target,
        pdfUrl: pdfUrl,
      };
    }
  }

  throw new Error(
    "Nota dengan nomor kuitansi " + target + " tidak ditemukan di Sheet 'Cetak_Nota'."
  );
}
