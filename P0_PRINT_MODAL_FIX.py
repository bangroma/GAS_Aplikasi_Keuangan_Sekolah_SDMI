from pathlib import Path
import shutil
from datetime import datetime

FILE = Path("Scripts.html")

def backup_file(path):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.stem}_BACKUP_{stamp}{path.suffix}")
    shutil.copy2(path, backup)
    return backup

def find_function_bytes(data, name):
    marker = f"function {name}(".encode("utf-8")
    start = data.find(marker)

    if start < 0:
        raise RuntimeError(f"Fungsi {name} tidak ditemukan.")

    brace = data.find(b"{", start)

    if brace < 0:
        raise RuntimeError(f"Pembuka fungsi {name} tidak ditemukan.")

    depth = 0
    quote = None
    escape = False
    i = brace

    while i < len(data):
        c = data[i]

        if quote is not None:
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

    raise RuntimeError(f"Penutup fungsi {name} tidak ditemukan.")

def replace_once(text, old, new, label):
    count = text.count(old)

    if count != 1:
        raise RuntimeError(
            f"{label}: ditemukan {count} kali, harus tepat 1 kali."
        )

    return text.replace(old, new, 1)

# ============================================================
# BACA RAW BYTE — JANGAN NORMALISASI FILE
# ============================================================

data = FILE.read_bytes()

if b"\r\n" not in data:
    print("P0 PRINT MODAL FIX v11: GAGAL")
    print("ERROR: Scripts.html tidak menggunakan CRLF.")
    print("Scripts.html TIDAK diubah.")
    raise SystemExit(1)

# Pastikan tidak ada LF tunggal pada file.
lf_only = data.replace(b"\r\n", b"")
if b"\n" in lf_only:
    print("P0 PRINT MODAL FIX v11: GAGAL")
    print("ERROR: ditemukan LF tunggal pada Scripts.html.")
    print("Scripts.html TIDAK diubah.")
    raise SystemExit(1)

start, end = find_function_bytes(data, "printModalTunggakan")

func_bytes = data[start:end]

# Normalisasi HANYA fungsi target di memori.
func = func_bytes.decode("utf-8").replace("\r\n", "\n")

original = func

# ============================================================
# 1. TOTAL PRINT
# ============================================================

old = """    let totalTagihan = 0;
    let totalDibayar = 0;
    let totalSisa = 0;

    currentModalTunggakanData.forEach(function (siswa) {
      totalTagihan += Number(siswa.total_tagihan) || 0;
      totalDibayar += Number(siswa.total_dibayar) || 0;
      totalSisa += Number(siswa.total_sisa) || 0;
    });"""

new = """    let totalTagihan = 0;
    let totalDibayar = 0;
    let totalSisa = 0;

    currentModalTunggakanData.forEach(function (siswa) {
      const tagihan = Array.isArray(siswa.tagihan)
        ? siswa.tagihan
        : [];

      tagihan.forEach(function (t) {
        const nominal = Number(t.nominal_akhir) || 0;
        const terbayar = Number(t.terbayar) || 0;
        const diskon = Number(t.diskon_tambahan_total) || 0;

        const efektif = Math.min(
          nominal,
          Math.max(0, terbayar + diskon)
        );

        totalTagihan += nominal;
        totalDibayar += terbayar;
        totalSisa += Math.max(0, nominal - efektif);
      });
    });"""

func = replace_once(
    func,
    old,
    new,
    "Agregasi total print modal"
)

# ============================================================
# 2. RINGKASAN SISWA
# ============================================================

old = """      const totalSiswa = Number(siswa.total_tagihan) || 0;
      const dibayarSiswa = Number(siswa.total_dibayar) || 0;
      const sisaSiswa = Number(siswa.total_sisa) || 0;

      let progressSiswa = Number(siswa.progress) || 0;

      progressSiswa = Math.max(0, Math.min(100, progressSiswa));

      const statusSiswa = dibayarSiswa > 0 ? "CICILAN" : "BELUM BAYAR";"""

new = """      const tagihanSiswa = Array.isArray(siswa.tagihan)
        ? siswa.tagihan
        : [];

      let totalSiswa = 0;
      let dibayarSiswa = 0;
      let efektifSiswa = 0;
      let sisaSiswa = 0;

      tagihanSiswa.forEach(function (t) {
        const nominal = Number(t.nominal_akhir) || 0;
        const terbayar = Number(t.terbayar) || 0;
        const diskon = Number(t.diskon_tambahan_total) || 0;

        const efektif = Math.min(
          nominal,
          Math.max(0, terbayar + diskon)
        );

        totalSiswa += nominal;
        dibayarSiswa += terbayar;
        efektifSiswa += efektif;
        sisaSiswa += Math.max(0, nominal - efektif);
      });

      const progressSiswa = totalSiswa > 0
        ? Math.max(
            0,
            Math.min(100, (efektifSiswa / totalSiswa) * 100)
          )
        : 100;

      const statusSiswa =
        sisaSiswa <= 0
          ? "LUNAS"
          : efektifSiswa > 0
            ? "CICILAN"
            : "BELUM BAYAR";"""

func = replace_once(
    func,
    old,
    new,
    "Ringkasan siswa print modal"
)

# ============================================================
# 3. DETAIL TAGIHAN
# ============================================================

old = """          const nominal = Number(t.nominal_akhir) || 0;

          const terbayar = Number(t.terbayar) || 0;

          const sisa = Number(t.sisa_tunggakan) || 0;

          let progress = Number(t.progress) || 0;

          /*
           * Jika backend belum mengirim progress,
           * hitung dari nominal dan pembayaran.
           */
          if (!progress && nominal > 0) {
            progress = (terbayar / nominal) * 100;
          }

          progress = Math.max(0, Math.min(100, progress));

          const statusTagihan =
            sisa <= 0 ? "LUNAS" : terbayar > 0 ? "CICILAN" : "BELUM BAYAR";"""

new = """          const nominal = Number(t.nominal_akhir) || 0;
          const terbayar = Number(t.terbayar) || 0;
          const diskon = Number(t.diskon_tambahan_total) || 0;

          const efektif = Math.min(
            nominal,
            Math.max(0, terbayar + diskon)
          );

          const sisa = Math.max(
            0,
            nominal - efektif
          );

          const progress = nominal > 0
            ? Math.max(
                0,
                Math.min(100, (efektif / nominal) * 100)
              )
            : 100;

          const statusTagihan =
            sisa <= 0
              ? "LUNAS"
              : efektif > 0
                ? "CICILAN"
                : "BELUM BAYAR";"""

func = replace_once(
    func,
    old,
    new,
    "Detail tagihan print modal"
)

# ============================================================
# VALIDASI PATCH
# ============================================================

if func == original:
    print("P0 PRINT MODAL FIX v11: GAGAL")
    print("ERROR: Tidak ada perubahan.")
    print("Scripts.html TIDAK diubah.")
    raise SystemExit(1)

# Konversi kembali HANYA fungsi target menjadi CRLF.
patched_func = func.replace("\n", "\r\n").encode("utf-8")

new_data = data[:start] + patched_func + data[end:]

# Pastikan seluruh file tetap CRLF.
lf_only_after = new_data.replace(b"\r\n", b"")

if b"\n" in lf_only_after:
    print("P0 PRINT MODAL FIX v11: GAGAL")
    print("ERROR: hasil patch mengandung LF tunggal.")
    print("Scripts.html TIDAK diubah.")
    raise SystemExit(1)

# Pastikan fungsi target memang berubah.
if data[start:end] == patched_func:
    print("P0 PRINT MODAL FIX v11: GAGAL")
    print("ERROR: fungsi target tidak berubah.")
    print("Scripts.html TIDAK diubah.")
    raise SystemExit(1)

# ============================================================
# BACKUP + WRITE
# ============================================================

backup = backup_file(FILE)
FILE.write_bytes(new_data)

print()
print("P0 PRINT MODAL FIX v11: BERHASIL")
print(f"BACKUP: {backup.name}")
print("Scripts.html berhasil diperbarui.")
print()
print("Logika print sekarang:")
print("- Terbayar = uang cash")
print("- Diskon tambahan = terpisah")
print("- Pelunasan efektif = cash + diskon")
print("- Sisa = nominal akhir - cash - diskon")
print("- Progress = pelunasan efektif / nominal akhir")
print("- Status = berdasarkan sisa")
print()
print("CRLF dipertahankan.")
print()
print("JANGAN clasp push dulu.")
