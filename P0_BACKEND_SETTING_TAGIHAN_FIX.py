from pathlib import Path
import shutil
from datetime import datetime

FILE = Path("06_Setting_Tagihan.js")

def backup_file(path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.stem}_BACKUP_{stamp}{path.suffix}")
    shutil.copy2(path, backup)
    return backup

def replace_exact(data, old, new, label):
    count = data.count(old)
    if count != 1:
        raise RuntimeError(
            f"{label}: pola ditemukan {count} kali, harus tepat 1 kali."
        )
    return data.replace(old, new, 1)

def main():
    if not FILE.exists():
        raise RuntimeError(f"{FILE} tidak ditemukan.")

    original = FILE.read_text(encoding="utf-8")

    # Pastikan file masih dalam kondisi yang kita audit.
    if "diskon_tambahan_total" in original:
        print("WARNING: diskon_tambahan_total sudah ada di file.")

    data = original

    old1 = """      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const sisa = Math.max(
        0,
        toNumber(t.sisa_tunggakan) || nominalAkhir - terbayar,
      );
"""

    new1 = """      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const diskon = toNumber(t.diskon_tambahan_total);
      const pelunasanEfektif = Math.min(
        nominalAkhir,
        Math.max(0, terbayar + diskon)
      );
      const sisa = Math.max(0, nominalAkhir - pelunasanEfektif);
"""

    data = replace_exact(
        data,
        old1,
        new1,
        "Blok Setting Tagihan #1"
    )

    old2 = """      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const sisa = Math.max(0, toNumber(t.sisa_tunggakan) || nominalAkhir - terbayar);
      
      let status = normalizeText(t.status) || "BELUM_BAYAR";
      if (sisa <= 0) {
        status = "LUNAS";
      } else if (terbayar > 0) {
        status = "CICILAN";
      }
"""

    new2 = """      const nominalAkhir = toNumber(t.nominal_akhir);
      const terbayar = toNumber(t.terbayar);
      const diskon = toNumber(t.diskon_tambahan_total);
      const pelunasanEfektif = Math.min(
        nominalAkhir,
        Math.max(0, terbayar + diskon)
      );
      const sisa = Math.max(0, nominalAkhir - pelunasanEfektif);

      let status = normalizeText(t.status) || "BELUM_BAYAR";
      if (sisa <= 0) {
        status = "LUNAS";
      } else if (pelunasanEfektif > 0) {
        status = "CICILAN";
      }
"""

    data = replace_exact(
        data,
        old2,
        new2,
        "Blok Setting Tagihan #2"
    )

    # Safety check: harus benar-benar berubah.
    if data == original:
        raise RuntimeError("Tidak ada perubahan.")

    # Safety check: tidak boleh ada pola lama.
    if old1 in data:
        raise RuntimeError("Pola lama #1 masih tersisa.")

    if old2 in data:
        raise RuntimeError("Pola lama #2 masih tersisa.")

    # Backup tambahan otomatis.
    backup = backup_file(FILE)

    # Pertahankan line ending asli.
    raw = FILE.read_bytes()
    had_crlf = b"\r\n" in raw

    if had_crlf:
        output = data.replace("\r\n", "\n").replace("\n", "\r\n")
    else:
        output = data.replace("\r\n", "\n")

    FILE.write_bytes(output.encode("utf-8"))

    # Validasi hasil baca kembali.
    result = FILE.read_text(encoding="utf-8")

    if "const diskon = toNumber(t.diskon_tambahan_total);" not in result:
        raise RuntimeError("Patch tidak ditemukan setelah penulisan.")

    print("P0 BACKEND SETTING TAGIHAN FIX: BERHASIL")
    print(f"BACKUP: {backup.name}")
    print("2 blok perhitungan Setting Tagihan sudah disinkronkan.")
    print("Rumus: pelunasanEfektif = terbayar + diskon.")
    print("Rumus: sisa = nominalAkhir - pelunasanEfektif.")
    print("Status CICILAN menggunakan pelunasan efektif.")
    print("JANGAN clasp push dulu.")

if __name__ == "__main__":
    try:
        main()
    except Exception as err:
        print("P0 BACKEND SETTING TAGIHAN FIX: GAGAL")
        print("ERROR:", err)
        print("06_Setting_Tagihan.js TIDAK diubah.")
        raise SystemExit(1)
