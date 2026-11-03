"""
Copyright (C) 2026 71658-dev
SPDX-License-Identifier: GPL-3.0-or-later

建置離線字庫資料夾 cnsdata/（只在打包前執行）

    python tools/build_cnsdata.py [輸出資料夾，預設為專案根目錄下的 cnsdata]
    python tools/build_cnsdata.py --manifest-only   只依現有 cnsdata 重算 cnsdata_manifest.py

從全字庫官網開放資料下載：
  - release.txt          版本資訊
  - Fonts_Kai.zip        全字庫正楷體（BMP、Ext-B、Plus 三個 TTF）
  - MapingTables.zip     CNS ↔ Unicode 對照表
  - Properties.zip       字屬性（取注音、部首、總筆畫、倉頡、筆順、部件、拼音）
並產生：
  cnsdata/fonts/*.ttf
  cnsdata/cns2unicode.tsv      字面-碼位 <TAB> Unicode 十六進位
  cnsdata/cns_phonetic.tsv     字面-碼位 <TAB> 注音（一字多音時多行）
  cnsdata/cns_radical_stroke.tsv  字面-碼位 <TAB> 部首序號 <TAB> 總筆畫
  cnsdata/cns_radicals.tsv     部首序號 <TAB> 部首字
  cnsdata/cns_cangjie.tsv、cns_strokes_sequence.tsv、cns_component.tsv、
          cns_component_ref.tsv、cns_pinyin.tsv   倉頡、筆順、部件、拼音（格式見 build_extra_props）
  cnsdata/VERSION.json         版本號與各檔 SHA-256
  cnsdata/LICENSE-OFL-1.1.txt  字型授權全文（OFL-1.1 要求散布時附上）
  cnsdata/來源與授權.txt
  cnsdata_manifest.py          （專案根目錄）字庫核對碼，打包進 exe 供啟動時比對
連線沿用主程式的安全連線設定（完整驗證憑證、只連全字庫官網）。
"""

import hashlib
import importlib.util
import io
import json
import os
import re
import sys
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_FILE = os.path.join(ROOT, "全字庫小幫手.py")
BASE = "https://www.cns11643.gov.tw/opendata/"
FONT_FILES = ("TW-Kai-98_1.ttf", "TW-Kai-Ext-B-98_1.ttf", "TW-Kai-Plus-98_1.ttf")


def load_app():
    spec = importlib.util.spec_from_file_location("cns_app", APP_FILE)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def zip_name(info: zipfile.ZipInfo) -> str:
    """官網 zip 內的中文檔名為 Big5 編碼（未設 UTF-8 旗標）"""
    if info.flag_bits & 0x800:
        return info.filename
    try:
        return info.filename.encode("cp437").decode("big5")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return info.filename


def download(session, headers, name: str) -> bytes:
    url = BASE + name
    t = time.time()
    print(f"下載 {url} …", flush=True)
    resp = session.get(url, headers=headers, timeout=(10, 900), allow_redirects=False)
    if resp.status_code != 200:
        raise RuntimeError(f"{name}：HTTP {resp.status_code}")
    print(f"  完成 {len(resp.content) / 1e6:.1f} MB，{time.time() - t:.0f} 秒", flush=True)
    return resp.content


def build_radical_stroke(props_zip: zipfile.ZipFile):
    """
    由 CNS_radical.txt（代碼→部首序號 1–214）、CNS_stroke.txt（代碼→總筆畫）、
    CNS_radical_word.txt（部首序號→部首字）產生：
      cns_radical_stroke.tsv  字面-碼位 <TAB> 部首序號 <TAB> 總筆畫
      cns_radicals.tsv        部首序號 <TAB> 部首字（如「乙（乚）」）
    """
    def read(name):
        info = next((i for i in props_zip.infolist() if os.path.basename(zip_name(i)).lower() == name.lower()), None)
        if info is None:
            raise RuntimeError(f"Properties.zip 缺少 {name}")
        return props_zip.read(info).decode("utf-8-sig", errors="replace").splitlines()

    def code_of(s):
        m = re.fullmatch(r"(\d{1,2})-([0-9A-Fa-f]{4})", s.strip())
        return f"{int(m.group(1))}-{m.group(2).upper()}" if m else None

    rad, stk = {}, {}
    for line in read("CNS_radical.txt"):
        p = line.split("\t")
        if len(p) >= 2 and code_of(p[0]) and p[1].strip().isdigit() and 1 <= int(p[1]) <= 214:
            rad[code_of(p[0])] = int(p[1])
    for line in read("CNS_stroke.txt"):
        p = line.split("\t")
        if len(p) >= 2 and code_of(p[0]) and p[1].strip().isdigit():
            stk[code_of(p[0])] = int(p[1])
    names = {}
    for line in read("CNS_radical_word.txt"):
        p = line.split("\t")
        if len(p) >= 2 and p[0].strip().isdigit() and p[1].strip():
            names[int(p[0])] = p[1].strip()
    codes = sorted(set(rad) & set(stk), key=lambda c: (int(c.split("-")[0]), c))
    if len(codes) < 50000 or len(names) != 214:
        raise RuntimeError(f"部首/筆畫資料異常：{len(codes)} 筆、部首 {len(names)} 個")
    rs = "".join(f"{c}\t{rad[c]}\t{stk[c]}\n" for c in codes).encode("utf-8")
    nm = "".join(f"{n}\t{names[n]}\n" for n in sorted(names)).encode("utf-8")
    return rs, nm, len(codes)


def _read_props(props_zip: zipfile.ZipFile, name: str) -> list:
    info = next((i for i in props_zip.infolist() if os.path.basename(zip_name(i)).lower() == name.lower()), None)
    if info is None:
        raise RuntimeError(f"Properties.zip 缺少 {name}")
    return props_zip.read(info).decode("utf-8-sig", errors="replace").splitlines()


def _norm_code(s: str):
    m = re.fullmatch(r"(\d{1,2})-([0-9A-Fa-f]{4})", s.strip())
    return f"{int(m.group(1))}-{m.group(2).upper()}" if m else None


def build_extra_props(props_zip: zipfile.ZipFile) -> dict:
    """
    倉頡、筆順、部件、拼音（同一個 Properties.zip），回傳 {檔名: 內容}：
      cns_cangjie.tsv            字面-碼位 <TAB> 倉頡碼（A–Z）
      cns_strokes_sequence.tsv   字面-碼位 <TAB> 筆順（1橫 2豎 3撇 4點 5折）
      cns_component.tsv          字面-碼位 <TAB> 部件代號,部件代號,…
      cns_component_ref.tsv      部件代號 <TAB> 部件 CNS 碼 <TAB> Unicode <TAB> 部件集合 <TAB> 筆畫
                                 （同一「部件集合」是同一部件的不同寫法，例如「火」與火字旁）
      cns_pinyin.tsv             注音 <TAB> 漢語 注音二式 耶魯 威妥瑪 通用（數字調）<TAB> 同五種（調號）
    """
    files, counts = {}, {}

    def simple(src, dst, pattern, minimum):
        rows = []
        for line in _read_props(props_zip, src):
            p = line.split("\t")
            code = _norm_code(p[0]) if p else None
            if code and len(p) >= 2 and re.fullmatch(pattern, p[1].strip()):
                rows.append(f"{code}\t{p[1].strip().upper()}\n")
        if len(rows) < minimum:
            raise RuntimeError(f"{src} 筆數異常：{len(rows)}")
        files[dst] = "".join(rows).encode("utf-8")
        counts[dst] = len(rows)

    simple("CNS_cangjie.txt", "cns_cangjie.tsv", r"[A-Za-z]{1,10}", 50000)
    simple("CNS_strokes_sequence.txt", "cns_strokes_sequence.tsv", r"[1-5]{1,99}", 50000)
    simple("CNS_component.txt", "cns_component.tsv", r"\d+(,\d+)*", 50000)

    word = {}
    for line in _read_props(props_zip, "CNS_component_word.txt"):
        p = line.split("\t")
        if len(p) >= 3 and p[0].strip().isdigit():
            word[p[0].strip()] = (p[1].strip(), p[2].strip())
    ref_rows = []
    for line in _read_props(props_zip, "CNS_component_ref.txt"):
        p = line.split("\t")
        if len(p) >= 3 and p[0].strip().isdigit() and p[0].strip() in word and _norm_code(p[1]) \
                and re.fullmatch(r"[0-9A-Fa-f]{4,6}", p[2].strip()):
            group, strokes = word[p[0].strip()]
            ref_rows.append(f"{p[0].strip()}\t{_norm_code(p[1])}\t{p[2].strip().upper()}\t{group}\t{strokes}\n")
    if len(ref_rows) < 400:
        raise RuntimeError(f"部件代號筆數異常：{len(ref_rows)}")
    files["cns_component_ref.tsv"] = "".join(ref_rows).encode("utf-8")
    counts["cns_component_ref.tsv"] = len(ref_rows)

    marks = {}
    for line in _read_props(props_zip, "CNS_pinyin_2.txt"):
        p = line.split("\t")
        if len(p) >= 6:
            marks[p[0].strip()] = [x.strip() for x in p[1:6]]
    py_rows = []
    for line in _read_props(props_zip, "CNS_pinyin_1.txt"):
        p = line.split("\t")
        if len(p) >= 6 and p[0].strip() in marks:
            py_rows.append("\t".join([p[0].strip()] + [x.strip() for x in p[1:6]] + marks[p[0].strip()]) + "\n")
    if len(py_rows) < 1000:
        raise RuntimeError(f"拼音筆數異常：{len(py_rows)}")
    files["cns_pinyin.tsv"] = "".join(py_rows).encode("utf-8")
    counts["cns_pinyin.tsv"] = len(py_rows)
    return files, counts


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


MANIFEST_FILE = os.path.join(ROOT, "cnsdata_manifest.py")


def write_app_manifest(data_dir: str):
    """
    依 cnsdata 現有檔案（fonts/ 的字型、*.tsv）寫出 cnsdata_manifest.py。
    主程式會 import 它，打包時一起包進 exe；啟動時比對，字庫被換掉就停用（見 OfflineCNSLibrary._verify）。
    """
    files = {}
    fonts_dir = os.path.join(data_dir, "fonts")
    names = [f"fonts/{n}" for n in os.listdir(fonts_dir) if n.lower().endswith((".ttf", ".otf", ".ttc"))]
    names += [n for n in os.listdir(data_dir) if n.lower().endswith(".tsv")]
    for rel in sorted(names):
        with open(os.path.join(data_dir, *rel.split("/")), "rb") as f:
            files[rel] = sha256(f.read())
    body = "".join(f"    {rel!r}: {digest!r},\n" for rel, digest in files.items())
    with open(MANIFEST_FILE, "w", encoding="utf-8", newline="\n") as f:
        f.write('"""離線字庫核對碼（由 tools/build_cnsdata.py 自動產生，請勿手動修改）"""\n\n'
                f"CNSDATA_SHA256 = {{\n{body}}}\n")
    print(f"已寫出 {MANIFEST_FILE}（{len(files)} 個檔）")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out_dir = os.path.abspath(args[0]) if args else os.path.join(ROOT, "cnsdata")
    if "--manifest-only" in sys.argv:
        write_app_manifest(out_dir)
        return
    app = load_app()
    session = app._make_secure_session()
    headers = app.CNSRareCharApp.HTTP_HEADERS

    release = download(session, headers, "release.txt").decode("utf-8-sig", errors="replace")
    m = re.search(r"Fonts_Kai\.zip\s*\n\s*版本：\s*(\d{8})", release)
    version = m.group(1) if m else time.strftime("%Y%m%d")
    maps_m = re.search(r"MapingTables\.zip\s*\n\s*版本：\s*(\d{8})", release)
    maps_version = maps_m.group(1) if maps_m else version

    fonts_zip = zipfile.ZipFile(io.BytesIO(download(session, headers, "Fonts_Kai.zip")))
    maps_zip = zipfile.ZipFile(io.BytesIO(download(session, headers, "MapingTables.zip")))

    os.makedirs(os.path.join(out_dir, "fonts"), exist_ok=True)
    manifest = {"version": version, "mapping_version": maps_version,
                "source": "數位發展部，CNS11643中文標準交換碼全字庫網站，https://www.cns11643.gov.tw",
                "files": {}}

    # 1. 字型
    found = set()
    for info in fonts_zip.infolist():
        name = zip_name(info)
        base = os.path.basename(name)
        if base in FONT_FILES:
            data = fonts_zip.read(info)
            with open(os.path.join(out_dir, "fonts", base), "wb") as f:
                f.write(data)
            manifest["files"][f"fonts/{base}"] = sha256(data)
            found.add(base)
        elif base.endswith(".txt"):
            data = fonts_zip.read(info)
            with open(os.path.join(out_dir, base), "wb") as f:
                f.write(data)
    missing = set(FONT_FILES) - found
    if missing:
        raise RuntimeError(f"Fonts_Kai.zip 缺少：{missing}")

    # 2. 對照表（CNS → Unicode）
    rows = {}
    for info in maps_zip.infolist():
        name = zip_name(info)
        if not re.search(r"CNS2UNICODE_Unicode .+\.txt$", name, re.I):
            continue
        text = maps_zip.read(info).decode("utf-8-sig", errors="replace")
        for line in text.splitlines():
            parts = line.split()
            if len(parts) < 2:
                continue
            cm = re.fullmatch(r"(\d{1,2})-([0-9A-Fa-f]{4})", parts[0])
            um = re.fullmatch(r"[0-9A-Fa-f]{4,6}", parts[1])
            if not (cm and um):
                continue
            plane = int(cm.group(1))
            if not 1 <= plane <= 24:
                continue
            rows[f"{plane}-{cm.group(2).upper()}"] = parts[1].upper()
    if len(rows) < 90000:
        raise RuntimeError(f"對照表筆數異常：{len(rows)}")
    tsv = "".join(f"{code}\t{cp}\n" for code, cp in
                  sorted(rows.items(), key=lambda kv: (int(kv[0].split("-")[0]), kv[0])))
    tsv_bytes = tsv.encode("utf-8")
    with open(os.path.join(out_dir, "cns2unicode.tsv"), "wb") as f:
        f.write(tsv_bytes)
    manifest["files"]["cns2unicode.tsv"] = sha256(tsv_bytes)
    manifest["mapping_count"] = len(rows)

    # 3. 注音（Properties.zip / CNS_phonetic.txt；一字多音時一行一個讀音，保留官網順序）
    props_m = re.search(r"Properties\.zip\s*\n\s*版本：\s*(\d{8})", release)
    manifest["phonetic_version"] = props_m.group(1) if props_m else version
    props_zip = zipfile.ZipFile(io.BytesIO(download(session, headers, "Properties.zip")))
    phon = next((i for i in props_zip.infolist()
                 if os.path.basename(zip_name(i)).lower() == "cns_phonetic.txt"), None)
    if phon is None:
        raise RuntimeError("Properties.zip 缺少 CNS_phonetic.txt")
    zy_rows = []
    for line in props_zip.read(phon).decode("utf-8-sig", errors="replace").splitlines():
        code, _, zy = line.strip().partition("\t")
        cm = re.fullmatch(r"(\d{1,2})-([0-9A-Fa-f]{4})", code)
        if cm and re.fullmatch(r"[ㄅ-ㄩˊˇˋ˙]{1,6}", zy.strip()):
            zy_rows.append(f"{int(cm.group(1))}-{cm.group(2).upper()}\t{zy.strip()}\n")
    if len(zy_rows) < 50000:
        raise RuntimeError(f"注音筆數異常：{len(zy_rows)}")
    zy_bytes = "".join(zy_rows).encode("utf-8")
    with open(os.path.join(out_dir, "cns_phonetic.tsv"), "wb") as f:
        f.write(zy_bytes)
    manifest["files"]["cns_phonetic.tsv"] = sha256(zy_bytes)
    manifest["phonetic_count"] = len(zy_rows)

    # 3b. 部首與總筆畫（同一個 Properties.zip）
    rs_bytes, names_bytes, rs_count = build_radical_stroke(props_zip)
    with open(os.path.join(out_dir, "cns_radical_stroke.tsv"), "wb") as f:
        f.write(rs_bytes)
    with open(os.path.join(out_dir, "cns_radicals.tsv"), "wb") as f:
        f.write(names_bytes)
    manifest["files"]["cns_radical_stroke.tsv"] = sha256(rs_bytes)
    manifest["files"]["cns_radicals.tsv"] = sha256(names_bytes)
    manifest["radical_stroke_count"] = rs_count

    # 3c. 倉頡、筆順、部件、拼音（同一個 Properties.zip）
    extra_files, extra_counts = build_extra_props(props_zip)
    for name, data in extra_files.items():
        with open(os.path.join(out_dir, name), "wb") as f:
            f.write(data)
        manifest["files"][name] = sha256(data)
    manifest["extra_counts"] = extra_counts

    # 4. 授權與來源
    with open(os.path.join(out_dir, "LICENSE-OFL-1.1.txt"), "w", encoding="utf-8") as f:
        f.write(OFL_TEXT)
    with open(os.path.join(out_dir, "來源與授權.txt"), "w", encoding="utf-8") as f:
        f.write(
            "本資料夾內容取自：數位發展部，CNS11643中文標準交換碼全字庫網站，https://www.cns11643.gov.tw\n"
            f"字型版本：{version}；對照表版本：{maps_version}\n\n"
            "授權：\n"
            "  - 全字庫正楷體字型（fonts/*.ttf）：(c) 數位發展部。字型內載明可依需求擇一適用"
            "政府資料開放授權條款-第1版或 SIL Open Font License 1.1；本資料夾依 OFL-1.1 散布，"
            "全文見 LICENSE-OFL-1.1.txt\n"
            "  - 中文碼對照表（cns2unicode.tsv，由官網 MapingTables.zip 轉換）：政府資料開放授權條款-第1版\n"
            "  - 注音、部首、筆畫、倉頡、筆順、部件、拼音（cns_*.tsv，由官網 Properties.zip 轉換）："
            "政府資料開放授權條款-第1版\n"
            "    https://data.gov.tw/license\n"
            "\n"
            "授權說明頁：https://www.cns11643.gov.tw/pageView.jsp?ID=59\n"
            "上述 tsv 檔為本專案由官網開放資料轉換格式而成（改作），內容以官網為準；"
            "本專案為非官方工具，與數位發展部無關。\n"
        )
    with open(os.path.join(out_dir, "VERSION.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    write_app_manifest(out_dir)

    print(f"完成：{out_dir}（字型版本 {version}，對照表 {len(rows)} 筆）")


OFL_TEXT = """Copyright (c) 數位發展部 (Ministry of Digital Affairs), CNS11643 中文標準交換碼全字庫

This Font Software is licensed under the SIL Open Font License, Version 1.1.
This license is copied below, and is also available with a FAQ at:
https://openfontlicense.org


-----------------------------------------------------------
SIL OPEN FONT LICENSE Version 1.1 - 26 February 2007
-----------------------------------------------------------

PREAMBLE
The goals of the Open Font License (OFL) are to stimulate worldwide
development of collaborative font projects, to support the font creation
efforts of academic and linguistic communities, and to provide a free and
open framework in which fonts may be shared and improved in partnership
with others.

The OFL allows the licensed fonts to be used, studied, modified and
redistributed freely as long as they are not sold by themselves. The
fonts, including any derivative works, can be bundled, embedded,
redistributed and/or sold with any software provided that any reserved
names are not used by derivative works. The fonts and derivatives,
however, cannot be released under any other type of license. The
requirement for fonts to remain under this license does not apply
to any document created using the fonts or their derivatives.

DEFINITIONS
"Font Software" refers to the set of files released by the Copyright
Holder(s) under this license and clearly marked as such. This may
include source files, build scripts and documentation.

"Reserved Font Name" refers to any names specified as such after the
copyright statement(s).

"Original Version" refers to the collection of Font Software components as
distributed by the Copyright Holder(s).

"Modified Version" refers to any derivative made by adding to, deleting,
or substituting -- in part or in whole -- any of the components of the
Original Version, by changing formats or by porting the Font Software to a
new environment.

"Author" refers to any designer, engineer, programmer, technical
writer or other person who contributed to the Font Software.

PERMISSION & CONDITIONS
Permission is hereby granted, free of charge, to any person obtaining
a copy of the Font Software, to use, study, copy, merge, embed, modify,
redistribute, and sell modified and unmodified copies of the Font
Software, subject to the following conditions:

1) Neither the Font Software nor any of its individual components,
in Original or Modified Versions, may be sold by itself.

2) Original or Modified Versions of the Font Software may be bundled,
redistributed and/or sold with any software, provided that each copy
contains the above copyright notice and this license. These can be
included either as stand-alone text files, human-readable headers or
in the appropriate machine-readable metadata fields within text or
binary files as long as those fields can be easily viewed by the user.

3) No Modified Version of the Font Software may use the Reserved Font
Name(s) unless explicit written permission is granted by the corresponding
Copyright Holder. This restriction only applies to the primary font name as
presented to the users.

4) The name(s) of the Copyright Holder(s) and the Author(s) of the Font
Software shall not be used to promote, endorse or advertise any
Modified Version, except to acknowledge the contribution(s) of the
Copyright Holder(s) and the Author(s) or with their explicit written
permission.

5) The Font Software, modified or unmodified, in part or in whole,
must be distributed entirely under this license, and must not be
distributed under any other license. The requirement for fonts to
remain under this license does not apply to any document created
using the Font Software.

TERMINATION
This license becomes null and void if any of the above conditions are
not met.

DISCLAIMER
THE FONT SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
OF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE
COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
INCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL
DAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM
OTHER DEALINGS IN THE FONT SOFTWARE.
"""


if __name__ == "__main__":
    main()
