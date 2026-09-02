from pathlib import Path
import shutil
from datetime import datetime
import re
import sys

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

    raise RuntimeError(f"Penutup fungsi {name} tidak ditemukan.")

def replace_once(data, old, new, label):
    count = data.count(old)

    if count != 1:
        raise RuntimeError(
            f"{label}: ditemukan {count} kali, harus tepat 1 kali."
        )

    return data.replace(old, new, 1)

def patch_render_function(data):
    start, end = find_function_bytes(
        data,
        "renderModalTunggakanKelas"
    )

    fn = data[start:end]

    old_total = (
        b'    let totalDibayar = 0;\r\n'
        b'    let totalSisa = 0;\r\n'
        b'\r\n'
        b'    data.forEach(function (siswa, siswaIndex) {\r\n'
        b'      totalDibayar += Number(siswa.total_dibayar) || 0;\r\n'
        b'      totalSisa += Number(siswa.total_sisa) || 0;\r\n'
        b'    });'
    )

    new_total = (
        b'    let totalDibayar = 0;\r\n'
        b'    let totalSisa = 0;\r\n'
        b'\r\n'
        b'    data.forEach(function (siswa) {\r\n'
        b'      const tagihan = Array.isArray(siswa.tagihan)\r\n'
        b'        ? siswa.tagihan\r\n'
        b'        : [];\r\n'
        b'\r\n'
        b'      tagihan.forEach(function (t) {\r\n'
        b'        const nominal = Number(t.nominal_akhir) || 0;\r\n'
        b'        const terbayar = Number(t.terbayar) || 0;\r\n'
        b'        const diskon = Number(t.diskon_tambahan_total) || 0;\r\n'
        b'\r\n'
        b'        const efektif = Math.min(\r\n'
        b'          nominal,\r\n'
        b'          Math.max(0, terbayar + diskon)\r\n'
        b'        );\r\n'
        b'\r\n'
        b'        totalDibayar += terbayar;\r\n'
        b'        totalSisa += Math.max(0, nominal - efektif);\r\n'
        b'      });\r\n'
        b'    });'
    )

    fn = replace_once(
        fn,
        old_total,
        new_total,
        "Agregasi total modal"
    )

    old_progress = (
        b'      const progress = Math.max(0, Math.min(100, Number(siswa.progress) || 0));\r\n'
        b'\r\n'
        b'      const statusClass =\r\n'
        b'        siswa.status === "CICILAN"\r\n'
        b'          ? "bg-amber-100 text-amber-700"\r\n'
        b'          : "bg-rose-100 text-rose-700";'
    )

    new_progress = (
        b'      const tagihan = Array.isArray(siswa.tagihan)\r\n'
        b'        ? siswa.tagihan\r\n'
        b'        : [];\r\n'
        b'\r\n'
        b'      let totalTagihanSiswa = 0;\r\n'
        b'      let totalCashSiswa = 0;\r\n'
        b'      let totalEfektifSiswa = 0;\r\n'
        b'      let totalSisaSiswa = 0;\r\n'
        b'\r\n'
        b'      tagihan.forEach(function (t) {\r\n'
        b'        const nominal = Number(t.nominal_akhir) || 0;\r\n'
        b'        const terbayar = Number(t.terbayar) || 0;\r\n'
        b'        const diskon = Number(t.diskon_tambahan_total) || 0;\r\n'
        b'\r\n'
        b'        const efektif = Math.min(\r\n'
        b'          nominal,\r\n'
        b'          Math.max(0, terbayar + diskon)\r\n'
        b'        );\r\n'
        b'\r\n'
        b'        totalTagihanSiswa += nominal;\r\n'
        b'        totalCashSiswa += terbayar;\r\n'
        b'        totalEfektifSiswa += efektif;\r\n'
        b'        totalSisaSiswa += Math.max(0, nominal - efektif);\r\n'
        b'      });\r\n'
        b'\r\n'
        b'      const progress = totalTagihanSiswa > 0\r\n'
        b'        ? Math.max(0, Math.min(100, (totalEfektifSiswa / totalTagihanSiswa) * 100))\r\n'
        b'        : 100;\r\n'
        b'\r\n'
        b'      const status = totalSisaSiswa <= 0\r\n'
        b'        ? "LUNAS"\r\n'
        b'        : totalEfektifSiswa > 0\r\n'
        b'          ? "CICILAN"\r\n'
        b'          : "BELUM BAYAR";\r\n'
        b'\r\n'
        b'      const statusClass =\r\n'
        b'        status === "LUNAS"\r\n'
        b'          ? "bg-emerald-100 text-emerald-700"\r\n'
        b'          : status === "CICILAN"\r\n'
        b'            ? "bg-amber-100 text-amber-700"\r\n'
        b'            : "bg-rose-100 text-rose-700";'
    )

    fn = replace_once(
        fn,
        old_progress,
        new_progress,
        "Perhitungan progress siswa"
    )

    old_status = b'            ${escapeHtmlLaporan(siswa.status)}'
    new_status = b'            ${escapeHtmlLaporan(status)}'

    fn = replace_once(
        fn,
        old_status,
        new_status,
        "Status siswa"
    )

    old_total_tagihan = b'              ${formatRupiah(siswa.total_tagihan)}'
    new_total_tagihan = b'              ${formatRupiah(totalTagihanSiswa)}'

    fn = replace_once(
        fn,
        old_total_tagihan,
        new_total_tagihan,
        "Total tagihan siswa"
    )

    old_dibayar = b'              ${formatRupiah(siswa.total_dibayar)}'
    new_dibayar = b'              ${formatRupiah(totalCashSiswa)}'

    fn = replace_once(
        fn,
        old_dibayar,
        new_dibayar,
        "Total dibayar siswa"
    )

    old_sisa = b'              ${formatRupiah(siswa.total_sisa)}'
    new_sisa = b'              ${formatRupiah(totalSisaSiswa)}'

    fn = replace_once(
        fn,
        old_sisa,
        new_sisa,
        "Total sisa siswa"
    )

    return data[:start] + fn + data[end:]

def patch_detail_function(data):
    start, end = find_function_bytes(
        data,
        "renderDetailTagihanModal"
    )

    fn = data[start:end]

    old = (
        b'        const progress = Math.max(0, Math.min(100, Number(t.progress) || 0));'
    )

    new = (
        b'        const nominal = Number(t.nominal_akhir) || 0;\r\n'
        b'        const terbayar = Number(t.terbayar) || 0;\r\n'
        b'        const diskon = Number(t.diskon_tambahan_total) || 0;\r\n'
        b'\r\n'
        b'        const efektif = Math.min(\r\n'
        b'          nominal,\r\n'
        b'          Math.max(0, terbayar + diskon)\r\n'
        b'        );\r\n'
        b'\r\n'
        b'        const sisa = Math.max(0, nominal - efektif);\r\n'
        b'\r\n'
        b'        const progress = nominal > 0\r\n'
        b'          ? Math.max(0, Math.min(100, (efektif / nominal) * 100))\r\n'
        b'          : 100;'
    )

    fn = replace_once(
        fn,
        old,
        new,
        "Progress detail tagihan"
    )

    old_sisa = b'              ${formatRupiah(t.sisa_tunggakan)}'
    new_sisa = b'              ${formatRupiah(sisa)}'

    fn = replace_once(
        fn,
        old_sisa,
        new_sisa,
        "Sisa detail tagihan"
    )

    old_terbayar = b'              ${formatRupiah(t.terbayar)}'
    new_terbayar = b'              ${formatRupiah(terbayar)}'

    fn = replace_once(
        fn,
        old_terbayar,
        new_terbayar,
        "Terbayar detail tagihan"
    )

    old_nominal = b'              ${formatRupiah(t.nominal_akhir)}'
    new_nominal = b'              ${formatRupiah(nominal)}'

    fn = replace_once(
        fn,
        old_nominal,
        new_nominal,
        "Nominal detail tagihan"
    )

    return data[:start] + fn + data[end:]

def main():
    if not FILE.exists():
        print("P0 MODAL FIX v8: GAGAL")
        print(f"ERROR: {FILE} tidak ditemukan.")
        sys.exit(1)

    original = FILE.read_bytes()

    if b"\r\n" not in original:
        print("P0 MODAL FIX v8: GAGAL")
        print("ERROR: Scripts.html tidak menggunakan CRLF.")
        print("File TIDAK diubah.")
        sys.exit(1)

    try:
        data = original

        data = patch_render_function(data)
        data = patch_detail_function(data)

        if data == original:
            raise RuntimeError("Tidak ada perubahan yang dihasilkan.")

        # Pastikan hasil tetap CRLF.
        if b"\r\n" not in data:
            raise RuntimeError("Line ending CRLF hilang setelah patch.")

        backup = backup_file(FILE)

        FILE.write_bytes(data)

        print()
        print("P0 MODAL FIX v8: BERHASIL")
        print(f"BACKUP: {backup.name}")
        print("Scripts.html berhasil diperbarui.")
        print()
        print("Perhitungan modal:")
        print("- Terbayar = uang cash")
        print("- Diskon tambahan = terpisah")
        print("- Pelunasan efektif = cash + diskon")
        print("- Sisa = nominal akhir - cash - diskon")
        print("- Progress = pelunasan efektif / nominal akhir")
        print("- Status = berdasarkan sisa")
        print()
        print("Line ending CRLF dipertahankan.")
        print()
        print("JANGAN clasp push dulu.")
        print("Lanjut:")
        print("  git diff --ignore-space-at-eol --stat -- Scripts.html")
        print("  git diff --ignore-space-at-eol -- Scripts.html | sed -n '1,260p'")

    except Exception as e:
        print()
        print("P0 MODAL FIX v8: GAGAL")
        print(f"ERROR: {e}")
        print()
        print("Scripts.html TIDAK diubah.")
        sys.exit(1)

if __name__ == "__main__":
    main()
