from pathlib import Path
import shutil
from datetime import datetime

FILE = Path("07_Laporan.js")

def backup_file(path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.stem}_BACKUP_{stamp}{path.suffix}")
    shutil.copy2(path, backup)
    return backup

def find_function(data, name):
    marker = ("function " + name + "(").encode("utf-8")
    start = data.find(marker)
    if start < 0:
        raise RuntimeError(f"Function {name} tidak ditemukan.")

    brace = data.find(b"{", start)
    if brace < 0:
        raise RuntimeError(f"Opening brace function {name} tidak ditemukan.")

    depth = 0
    quote = None
    escape = False
    i = brace

    while i < len(data):
        c = data[i]

        if quote:
            if escape:
                escape = False
            elif c == 92:
                escape = True
            elif c == quote:
                quote = None
        else:
            if c in (34, 39, 96):
                quote = c
            elif c == 123:
                depth += 1
            elif c == 125:
                depth -= 1
                if depth == 0:
                    return start, i + 1

        i += 1

    raise RuntimeError(f"Closing brace function {name} tidak ditemukan.")

def main():
    original = FILE.read_bytes()

    old_crlf = original.count(b"\r\n")
    old_lf = original.count(b"\n")

    start, end = find_function(original, "getDetailTunggakanPerKelas")

    new_function = r'''function getDetailTunggakanPerKelas(kelas, idPos) {
  kelas = normalizeText(kelas);
  idPos = normalizeText(idPos);

  if (!kelas) throw new Error("Kelas wajib diisi.");

  const siswaList = getSheetDataAsObjects("Siswa");
  let tagihanList = getSheetDataAsObjects("Tagihan_Siswa");

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

      totalTagihan += nominalAkhir;
      totalDibayar += terbayar;
      totalEfektif += efektif;
      totalSisa += sisa;

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

    if (totalSisa <= 0) return;

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
      total_diskon: Math.max(0, totalEfektif - totalDibayar),
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

    replacement = new_function.encode("utf-8")

    updated = original[:start] + replacement + original[end:]

    new_crlf = updated.count(b"\r\n")
    new_lf = updated.count(b"\n")
    new_bare_lf = new_lf - new_crlf
    new_bare_cr = updated.count(b"\r") - new_crlf

    if new_crlf != old_crlf or new_lf != old_lf + replacement.count(b"\n") - original[start:end].count(b"\n"):
        raise RuntimeError(
            "Validasi line ending gagal."
        )

    if new_bare_cr != 0:
        raise RuntimeError(
            f"Ditemukan bare CR baru: {new_bare_cr}"
        )

    backup = backup_file(FILE)
    FILE.write_bytes(updated)

    print("P0 BACKEND TUNGGAKAN FIX: BERHASIL")
    print(f"BACKUP: {backup}")
    print("Function getDetailTunggakanPerKelas() berhasil diperbarui.")
    print()
    print("Logika:")
    print("- Terbayar = uang cash")
    print("- Diskon tambahan = terpisah")
    print("- Pelunasan efektif = cash + diskon")
    print("- Sisa = nominal akhir - pelunasan efektif")
    print("- Progress = pelunasan efektif / nominal akhir")
    print("- Status = berdasarkan sisa")
    print()
    print(f"CRLF: {old_crlf} -> {new_crlf}")
    print(f"LF:   {old_lf} -> {new_lf}")
    print(f"Bare LF: {new_bare_lf}")
    print(f"Bare CR: {new_bare_cr}")

if __name__ == "__main__":
    main()
