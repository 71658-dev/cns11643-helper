"""
Copyright (C) 2026 71658-dev
SPDX-License-Identifier: GPL-3.0-or-later

產生第三方元件授權聲明 THIRD-PARTY-NOTICES.txt（打包時由 全字庫小幫手.spec 自動呼叫）

    python tools/collect_licenses.py [輸出資料夾，預設 dist/全字庫小幫手]

打包後的 exe 內含 Python 與下列套件，依各自授權條款，散布時必須附上著作權聲明與授權全文：
Python（含 OpenSSL、libffi、zlib、Tcl/Tk 等，見 Python 的 LICENSE.txt）、requests、urllib3、
idna、charset-normalizer、certifi、Pillow（含 FreeType、libjpeg、libpng 等）、pywin32、
PyInstaller 開機程式。授權全文一律從本機安裝的套件讀取，避免手抄出錯。
"""

import importlib.metadata as md
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# (套件名稱, 用途, 原始碼網址)
PACKAGES = [
    ("requests", "連線全字庫官網", "https://github.com/psf/requests"),
    ("urllib3", "requests 的連線元件", "https://github.com/urllib3/urllib3"),
    ("idna", "requests 的網域名稱處理", "https://github.com/kjd/idna"),
    ("charset-normalizer", "requests 的文字編碼偵測", "https://github.com/jawah/charset_normalizer"),
    ("certifi", "HTTPS 根憑證清單", "https://github.com/certifi/python-certifi"),
    ("pillow", "字形圖片處理", "https://github.com/python-pillow/Pillow"),
    ("pywin32", "Windows 剪貼簿與手寫辨識", "https://github.com/mhammond/pywin32"),
    ("pyinstaller", "打包工具（exe 內含其開機程式 bootloader）", "https://github.com/pyinstaller/pyinstaller"),
]

SEP = "=" * 78


def _license_files(dist_name: str) -> list:
    """套件 dist-info 裡的授權檔（LICENSE、NOTICE、COPYING）"""
    out = []
    for f in md.files(dist_name) or []:
        name = str(f).replace("\\", "/")
        base = name.rsplit("/", 1)[-1].upper()
        if ".dist-info/" in name and any(k in base for k in ("LICENSE", "LICENCE", "NOTICE", "COPYING")):
            out.append(f)
    return sorted(out, key=str)


def build_notices() -> str:
    parts = [
        "全字庫小幫手 — 第三方元件授權聲明（THIRD-PARTY NOTICES）",
        "",
        "本程式以 GPL-3.0-or-later 授權（見 LICENSE）。打包後的程式內含下列第三方元件，",
        "各元件仍依其原授權條款授權，著作權屬於各自的作者。以下依序附上授權全文。",
        "",
        f"  Python {sys.version.split()[0]}（PSF License；含 OpenSSL、libffi、zlib、bzip2、xz、Tcl/Tk 等）",
        "      https://www.python.org/",
    ]
    for dist, purpose, url in PACKAGES:
        meta = md.metadata(dist)
        lic = meta.get("License-Expression") or meta.get("License") or ""
        lic = lic.splitlines()[0][:80] if lic else ""
        parts.append(f"  {meta['Name']} {md.version(dist)}（{lic}）— {purpose}")
        parts.append(f"      {url}")
    parts.append("")

    py_license = os.path.join(sys.base_prefix, "LICENSE.txt")
    with open(py_license, encoding="utf-8", errors="replace") as f:
        parts += [SEP, f"Python {sys.version.split()[0]} — LICENSE.txt", SEP, f.read().rstrip(), ""]

    for dist, _, _ in PACKAGES:
        files = _license_files(dist)
        if not files:
            raise RuntimeError(f"{dist} 找不到授權檔，請確認套件安裝完整")
        for f in files:
            text = f.locate().read_bytes().decode("utf-8", errors="replace")
            parts += [SEP, f"{md.metadata(dist)['Name']} {md.version(dist)} — {str(f).replace(chr(92), '/')}",
                      SEP, text.rstrip(), ""]
    return "\n".join(parts) + "\n"


def write_notices(out_dir: str):
    """寫出 THIRD-PARTY-NOTICES.txt，並複製本專案的 LICENSE（GPL-3.0 全文）"""
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "THIRD-PARTY-NOTICES.txt"), "w", encoding="utf-8", newline="\r\n") as f:
        f.write(build_notices())
    shutil.copyfile(os.path.join(ROOT, "LICENSE"), os.path.join(out_dir, "LICENSE.txt"))
    print(f"已寫出授權聲明：{out_dir}")


if __name__ == "__main__":
    write_notices(os.path.abspath(sys.argv[1]) if len(sys.argv) > 1
                  else os.path.join(ROOT, "dist", "全字庫小幫手"))
