from pathlib import Path
import shutil
import sys
import re
from datetime import datetime

ROOT = Path(__file__).resolve().parent
TARGET_FILES = ["00_Config.gs", "02_Kasir_Bayar.gs"]
BACKUP_PREFIX = "P0_BACKUP_"


def read_file(name):
    path = ROOT / name
    if not path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {name}")
    return path.read_text(encoding="utf-8")


def backup_files():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = ROOT / f"{BACKUP_PREFIX}{stamp}"
    backup_dir.mkdir(parents=True, exist_ok=False)

    for name in TARGET_FILES:
        shutil.copy2(ROOT / name, backup_dir / name)

    print(f"BACKUP: {backup_dir.name}")
    return backup_dir


def replace_once(text, old, new, label):
    count = text.count(old)

    if count != 1:
        raise RuntimeError(
            f"{label}: ditemukan {count}x, harus tepat 1x."
        )

    return text.replace(old, new, 1)


def replace_regex(text, pattern, replacement, label, flags=0):
    text, count = re.subn(
        pattern,
        replacement,
        text,
        count=1,
        flags=flags,
    )

    if count != 1:
        raise RuntimeError(
            f"{label}: pola tidak ditemukan tepat 1x."
        )

    return text


def patch_config(text):
    if '"diskon_tambahan_total",' in text:
        print("00_Config.gs: diskon_tambahan_total sudah ada")
        return text

    pattern = (
        r'("nominal_potongan",\s*'
        r'"nominal_akhir",\s*)'
        r'("status",)'
    )

    replacement = (
        r'\1'
        '"diskon_tambahan_total",\n'
        r'    \2'
    )

    return replace_regex(
        text,
        pattern,
        replacement,
        "00_Config.gs / Tagihan_Siswa",
        flags=re.MULTILINE,
    )


def patch_get_tagihan(text):
    if "const diskonTambahanTotal = toNumber(t.diskon_tambahan_total);" in text:
        print("getTagihanSiswa: diskon tambahan sudah ada")
        return text

    pattern = (
        r'(\s*const terbayar = toNumber\(t\.terbayar\);\s*)'
        r'const sisaTunggakan = Math\.max\(0, nominalAkhir - terbayar\);'
    )

    replacement = '''\\1
      const diskonTambahanTotal = toNumber(t.diskon_tambahan_total);
      const pelunasanEfektif = terbayar + diskonTambahanTotal;
      const sisaTunggakan = Math.max(
        0,
        nominalAkhir - pelunasanEfektif,
      );'''

    text = replace_regex(
        text,
        pattern,
        replacement,
        "getTagihanSiswa / sisa",
        flags=re.MULTILINE,
    )

    text = replace_once(
        text,
        '} else if (terbayar > 0) {',
        '} else if (pelunasanEfektif > 0) {',
        "getTagihanSiswa / status",
    )

    pattern = (
        r'(nominal_akhir:\s*nominalAkhir,\s*)'
        r'(terbayar:\s*terbayar,)'
    )

    replacement = (
        r'\1'
        '        diskon_tambahan_total: diskonTambahanTotal,\n'
        r'        \2'
    )

    text = replace_regex(
        text,
        pattern,
        replacement,
        "getTagihanSiswa / return",
        flags=re.MULTILINE,
    )

    return text


def patch_simpan_transaksi(text):
    # 1. requiredTagihan
    if '"diskon_tambahan_total"' not in text:
        pattern = (
            r'("nominal_akhir",\s*)'
            r'("terbayar",)'
        )

        replacement = (
            r'\1'
            '"diskon_tambahan_total",\n      '
            r'\2'
        )

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / requiredTagihan",
            flags=re.MULTILINE,
        )
        print("requiredTagihan: OK")
    else:
        print("requiredTagihan: diskon_tambahan_total sudah ada")

    # 2. kolom diskon tambahan
    if "diskonTambahanTotal: tagihanHeaders.indexOf" not in text:
        pattern = (
            r'(nominalAkhir:\s*tagihanHeaders\.indexOf\("nominal_akhir"\),\s*)'
            r'(terbayar:\s*tagihanHeaders\.indexOf\("terbayar"\),)'
        )

        replacement = (
            r'\1'
            '      diskonTambahanTotal: tagihanHeaders.indexOf("diskon_tambahan_total"),\n\n'
            r'      \2'
        )

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / col",
            flags=re.MULTILINE,
        )
        print("col diskonTambahanTotal: OK")
    else:
        print("col diskonTambahanTotal: sudah ada")

    # 3. saldo awal
    if "const currentDiskonTotal = toNumber(row[col.diskonTambahanTotal]);" not in text:
        pattern = (
            r'(const nominalAkhir = toNumber\(row\[col\.nominalAkhir\]\);\s*)'
            r'(const currentTerbayar = toNumber\(row\[col\.terbayar\]\);\s*)'
            r'(const currentSisa = calculateSisa\(nominalAkhir, currentTerbayar\);)'
        )

        replacement = '''\\1
      \\2
      const currentDiskonTotal = toNumber(
        row[col.diskonTambahanTotal],
      );

      const currentSisa = calculateSisa(
        nominalAkhir,
        currentTerbayar + currentDiskonTotal,
      );

      const diskonTambahan = Math.max(
        0,
        toNumber(item.diskon_tambahan),
      );

      const pelunasanEfektif =
        item.nominal_dibayar + diskonTambahan;'''

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / saldo awal",
            flags=re.MULTILINE,
        )
        print("saldo awal: OK")
    else:
        print("saldo awal: sudah ada")

    # 4. validasi pembayaran + diskon
    if "if (diskonTambahan > currentSisa)" not in text:
        pattern = (
            r'(\s*if \(item\.nominal_dibayar > currentSisa\) \{.*?'
            r'\n\s*\}\s*)'
            r'(\s*const pos = posList\.find)'
        )

        replacement = '''\\1

      if (diskonTambahan > currentSisa) {
        throw new Error(
          "Diskon tambahan tagihan " +
            item.id_tagihan +
            " melebihi sisa tagihan. Sisa: " +
            formatRupiah(currentSisa),
        );
      }

      if (pelunasanEfektif > currentSisa) {
        throw new Error(
          "Pembayaran + diskon tambahan tagihan " +
            item.id_tagihan +
            " melebihi sisa tagihan. Sisa: " +
            formatRupiah(currentSisa),
        );
      }

      \\2'''

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / validasi diskon",
            flags=re.MULTILINE | re.DOTALL,
        )
        print("validasi diskon: OK")
    else:
        print("validasi diskon: sudah ada")

    # 5. cicilan berdasarkan pelunasan efektif
    if "const isPartial = pelunasanEfektif < currentSisa;" not in text:
        text = replace_once(
            text,
            "const isPartial = item.nominal_dibayar < currentSisa;",
            "const isPartial = pelunasanEfektif < currentSisa;",
            "simpanTransaksi / cicilan",
        )
        print("validasi cicilan: OK")
    else:
        print("validasi cicilan: sudah ada")

    # 6. saldo baru
    if "const newDiskonTambahanTotal =" not in text:
        pattern = (
            r'(\s*const newTerbayar = currentTerbayar \+ item\.nominal_dibayar;\s*)'
            r'(\s*const newSisa = calculateSisa\(nominalAkhir, newTerbayar\);\s*)'
            r'(\s*const newStatus = newSisa <= 0 \? "LUNAS" : "CICILAN";)'
        )

        replacement = '''\\1

      const newDiskonTambahanTotal =
        currentDiskonTotal + diskonTambahan;

      const newSisa = calculateSisa(
        nominalAkhir,
        newTerbayar + newDiskonTambahanTotal,
      );

      const newStatus =
        newSisa <= 0 ? "LUNAS" : "CICILAN";'''

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / saldo baru",
            flags=re.MULTILINE,
        )
        print("saldo baru: OK")
    else:
        print("saldo baru: sudah ada")

    # 7. snapshot rollback
    if "diskonTambahanTotal: row[col.diskonTambahanTotal]" not in text:
        pattern = (
            r'(originalTagihan\.push\(\{\s*'
            r'rowIndex:\s*found\.rowIndex,\s*)'
            r'(terbayar:\s*row\[col\.terbayar\],)'
        )

        replacement = (
            r'\1'
            '        diskonTambahanTotal: row[col.diskonTambahanTotal],\n\n'
            r'        \2'
        )

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / rollback snapshot",
            flags=re.MULTILINE,
        )
        print("rollback snapshot: OK")
    else:
        print("rollback snapshot: sudah ada")

    # 8. updatePlan
    if "newDiskonTambahanTotal: newDiskonTambahanTotal" not in text:
        pattern = (
            r'(updatePlan\.push\(\{\s*'
            r'rowIndex:\s*found\.rowIndex,\s*)'
            r'(newTerbayar:\s*newTerbayar,)'
        )

        replacement = (
            r'\1'
            '        newDiskonTambahanTotal: newDiskonTambahanTotal,\n\n'
            r'        \2'
        )

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / updatePlan",
            flags=re.MULTILINE,
        )
        print("updatePlan: OK")
    else:
        print("updatePlan: sudah ada")

    # 9. tulis diskon ke Tagihan_Siswa
    if "update.newDiskonTambahanTotal" not in text:
        pattern = (
            r'(updatePlan\.forEach\(function \(update\) \{\s*)'
            r'(tagihanSheet\s*'
            r'\.getRange\(update\.rowIndex, col\.terbayar \+ 1\)\s*'
            r'\.setValue\(update\.newTerbayar\);)'
        )

        replacement = '''\\1
      tagihanSheet
        .getRange(update.rowIndex, col.diskonTambahanTotal + 1)
        .setValue(update.newDiskonTambahanTotal);

      \\2'''

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / update tagihan",
            flags=re.MULTILINE,
        )
        print("update tagihan: OK")
    else:
        print("update tagihan: sudah ada")

    # 10. rollback diskon
    if "tagihanMap.diskon_tambahan_total" not in text:
        pattern = (
            r'(originalTagihan\.forEach\(function \(original\) \{\s*)'
            r'(if \(tagihanMap\.terbayar !== undefined\) \{)'
        )

        replacement = '''\\1
        if (tagihanMap.diskon_tambahan_total !== undefined) {
          tagihanSheet
            .getRange(
              original.rowIndex,
              tagihanMap.diskon_tambahan_total + 1,
            )
            .setValue(original.diskonTambahanTotal);
        }

        \\2'''

        text = replace_regex(
            text,
            pattern,
            replacement,
            "simpanTransaksi / rollback diskon",
            flags=re.MULTILINE,
        )
        print("rollback diskon: OK")
    else:
        print("rollback diskon: sudah ada")

    return text


def verify(config, kasir):
    checks = [
        ('"diskon_tambahan_total",', config, "Schema diskon_tambahan_total"),
        (
            "const diskonTambahanTotal = toNumber(t.diskon_tambahan_total);",
            kasir,
            "getTagihanSiswa diskon",
        ),
        ("pelunasanEfektif", kasir, "Pelunasan efektif"),
        ("diskonTambahanTotal: tagihanHeaders.indexOf", kasir, "Kolom diskon"),
        ("currentDiskonTotal", kasir, "Saldo diskon saat ini"),
        ("newDiskonTambahanTotal", kasir, "Saldo diskon baru"),
        ("update.newDiskonTambahanTotal", kasir, "Update diskon"),
        (
            "original.diskonTambahanTotal",
            kasir,
            "Rollback diskon",
        ),
        (
            "diskon_tambahan_total: diskonTambahanTotal",
            kasir,
            "Return diskon",
        ),
    ]

    failed = []

    for needle, text, label in checks:
        if needle not in text:
            failed.append(label)

    if failed:
        raise RuntimeError(
            "VERIFIKASI GAGAL:\n- " +
            "\n- ".join(failed)
        )

    print("VERIFY: PASS")


def main():
    print("=" * 60)
    print("P0 FIX v3 - DISKON TAMBAHAN TAGIHAN")
    print("=" * 60)

    config = read_file("00_Config.gs")
    kasir = read_file("02_Kasir_Bayar.gs")

    backup_dir = backup_files()

    try:
        print("Patch 00_Config.gs...")
        new_config = patch_config(config)

        print("Patch 02_Kasir_Bayar.gs...")
        new_kasir = patch_get_tagihan(kasir)
        new_kasir = patch_simpan_transaksi(new_kasir)

        print("Verifikasi...")
        verify(new_config, new_kasir)

        (ROOT / "00_Config.gs").write_text(
            new_config,
            encoding="utf-8",
        )

        (ROOT / "02_Kasir_Bayar.gs").write_text(
            new_kasir,
            encoding="utf-8",
        )

        print()
        print("P0 FIX v3: BERHASIL")
        print(f"Backup: {backup_dir.name}")
        print()
        print("File berubah:")
        print("  - 00_Config.gs")
        print("  - 02_Kasir_Bayar.gs")
        print()
        print("status_siswa TIDAK disentuh.")
        print("diskon_tambahan_total sudah terintegrasi.")
        print("Terbayar tetap hanya uang aktual.")

    except Exception as e:
        print()
        print("P0 FIX v3: GAGAL")
        print(f"Error: {e}")
        print(f"Backup aman di: {backup_dir.name}")
        print("Tidak ada file yang ditulis jika verifikasi gagal.")
        sys.exit(1)


if __name__ == "__main__":
    main()
