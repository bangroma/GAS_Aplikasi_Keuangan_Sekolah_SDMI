from pathlib import Path
import shutil
from datetime import datetime

FILE = Path("07_Laporan.js")
FUNCTION = "getSiswaByKelasForLaporan"

def find_function(data, name):
    marker = "function " + name + "("
    start = data.find(marker)
    if start < 0:
        raise RuntimeError(f"Function {name} tidak ditemukan.")

    brace = data.find("{", start)
    if brace < 0:
        raise RuntimeError(f"Opening brace {name} tidak ditemukan.")

    depth = 0
    in_string = None
    escape = False
    in_line_comment = False
    in_block_comment = False

    i = brace
    while i < len(data):
        ch = data[i]
        nxt = data[i + 1] if i + 1 < len(data) else ""

        if in_line_comment:
            if ch == "\n":
                in_line_comment = False
        elif in_block_comment:
            if ch == "*" and nxt == "/":
                in_block_comment = False
                i += 1
        elif in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == in_string:
                in_string = None
        else:
            if ch in ("'", '"', "`"):
                in_string = ch
            elif ch == "/" and nxt == "/":
                in_line_comment = True
                i += 1
            elif ch == "/" and nxt == "*":
                in_block_comment = True
                i += 1
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return start, i + 1
        i += 1

    raise RuntimeError(f"Closing brace {name} tidak ditemukan.")

def main():
    if not FILE.exists():
        raise RuntimeError(f"{FILE} tidak ditemukan.")

    raw = FILE.read_bytes()

    if b"\r\n" in raw:
        raise RuntimeError("07_Laporan.js ternyata CRLF. Patch dibatalkan.")
    if b"\r" in raw:
        raise RuntimeError("07_Laporan.js mengandung CR. Patch dibatalkan.")

    data = raw.decode("utf-8")
    start, end = find_function(data, FUNCTION)

    old = data[start:end]

    if "diskon_tambahan_total" in old:
        raise RuntimeError(
            "Function sudah mengandung diskon_tambahan_total. "
            "Patch dibatalkan agar tidak double patch."
        )

    new = r'''function getSiswaByKelasForLaporan(kelas, idPos) {
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
}'''

    backup = FILE.with_name(
        f"{FILE.stem}_BACKUP_{datetime.now().strftime('%Y%m%d_%H%M%S')}{FILE.suffix}"
    )
    shutil.copy2(FILE, backup)

    patched = data[:start] + new + data[end:]
    patched_bytes = patched.encode("utf-8")

    if b"\r\n" in patched_bytes or b"\r" in patched_bytes:
        shutil.copy2(backup, FILE)
        raise RuntimeError("Line ending berubah. File dikembalikan.")

    if patched.count("function getSiswaByKelasForLaporan(") != 1:
        shutil.copy2(backup, FILE)
        raise RuntimeError("Jumlah function target tidak valid. File dikembalikan.")

    if patched.count("diskon_tambahan_total") < 2:
        shutil.copy2(backup, FILE)
        raise RuntimeError("Field diskon tidak ditemukan lengkap. File dikembalikan.")

    FILE.write_bytes(patched_bytes)

    print("P0 BACKEND LAPORAN FIX: BERHASIL")
    print("BACKUP:", backup.name)
    print("Function diperbarui:", FUNCTION)
    print("Logika cash/diskon/efektif/sisa/progress/status sudah disinkronkan.")
    print("07_Laporan.js tetap LF.")
    print("JANGAN clasp push dulu.")

if __name__ == "__main__":
    main()
