#!/usr/bin/env python3
"""
Sematkan logo Twenty Swim (PNG) sebagai base64 data URI ke semua file HTML di folder output/.

Cara pakai (dari folder repo):
    python3 output/tools/sematkan-logo.py logo-twentyswim.png

Setiap file HTML memiliki blok penanda:
    <!-- LOGO:START --> ... <!-- LOGO:END -->
Isi di antara penanda diganti dengan <img src="data:image/png;base64,..."> tanpa filter CSS.
Skrip ini aman dijalankan berulang kali.
"""
import base64
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
PATTERN = re.compile(r"(<!-- LOGO:START -->)(.*?)(<!-- LOGO:END -->)", re.S)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    logo_path = pathlib.Path(sys.argv[1])
    if not logo_path.exists():
        print(f"File logo tidak ditemukan: {logo_path}")
        return 1
    data = base64.b64encode(logo_path.read_bytes()).decode("ascii")
    ext = logo_path.suffix.lower().lstrip(".")
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "svg": "image/svg+xml", "webp": "image/webp"}.get(ext, "image/png")
    img = f'<img class="logo" alt="Twenty Swim Swimming School" src="data:{mime};base64,{data}">'

    changed = 0
    for html in sorted(ROOT.rglob("*.html")):
        text = html.read_text(encoding="utf-8")
        new_text, n = PATTERN.subn(lambda m: f"{m.group(1)}{img}{m.group(3)}", text)
        if n:
            html.write_text(new_text, encoding="utf-8")
            changed += 1
            print(f"logo disematkan ({n} slot): {html.relative_to(ROOT)}")
    print(f"Selesai. {changed} file diperbarui.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
