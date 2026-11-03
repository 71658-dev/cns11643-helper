"""
================================================================================
全字庫小幫手（CNS11643 Helper）
Copyright (C) 2026 71658-dev
SPDX-License-Identifier: GPL-3.0-or-later

本程式為自由軟體：您可以依自由軟體基金會發布的 GNU 通用公眾授權條款第 3 版，
或（依您的選擇）任何更新的版本，重新散布或修改本程式。
本程式散布的目的是希望有用，但不提供任何擔保，亦無可售性或特定用途適用性的默示擔保。
詳見 GNU 通用公眾授權條款（LICENSE）。原始碼：https://github.com/71658-dev/cns11643-helper

用途：公文或申請書套印時，申請人姓名有電腦打不出、印不出的罕用字，
      不必安裝造字檔（EUDC），在 Word 直接貼上就能正常顯示與列印。

主要功能：
  1. 找字（查詢方式分頁，除 Unicode 外皆對應全字庫官網的查詢方式，全部離線查整個字庫）
     - CNS 代碼：第 1–19 字面，如 1-4442、12-6170、C-6170；也可查範圍（3-4200~3-42FF）或前幾碼（3-42）
     - 注音：可不加聲調（各聲調都列）、加 ˉ 只要一聲；也可輸入一個同音字（如「宣」）
     - 拼音：漢語、注音二式、耶魯、威妥瑪、通用；數字調（xuan2）或調號（xuán）
     - 部首/筆畫：部首可選清單或輸入「火」「86」，可再加注音一起篩選（如 火部＋ㄒㄩㄢ）
     - 部件：輸入組成的字（火宣）或從部件表點選；模糊＝包含、精確＝剛好由這些部件組成（可用 ＊ ？）
       同一部件的不同寫法（例如「火」與火字旁）視為相同
     - 倉頡：英文字母或字根（FJMM／火十一一），可用 ＊ ？
     - 筆順：橫豎撇點折（1–5），可用 ＊ ？
     - Unicode：碼位（U+714A）或直接貼上字
     - 複合：以上條件同時篩選（與官網「複合查詢」相同）
     - 「到官網查詢」開啟全字庫官網對應的查詢頁
  2. 取得字形，依序嘗試：
     (1) 離線字庫 cnsdata/（全字庫正楷體字型直接繪製，不需網路）
     (2) 本機快取 cache/
     (3) 連線全字庫官網下載（宋體、離線字庫沒有的字），下載後存入快取；
         官網沒有楷書時改用宋體並提醒；連不上時只給預覽用備援圖，禁止複製
  3. 複製到 Word：寫入剪貼簿的 RTF 格式（圖片尺寸以 twip、垂直位置以半點指定），
     在 Word 將「?」反白後按 Ctrl+V，就成為大小剛好、與前後文字對齊的行內字。
     可選字體（標楷體/宋體）、Word 字級（12–36 pt）與垂直對齊微調。
     另附 CF_DIB 圖片格式，供小畫家等軟體貼上。
  4. 手繪造字：全字庫查不到的字可直接手寫；以 Windows 內建的中文(繁體)手寫辨識
     列出候選字（本機辨識，筆跡不外傳），或直接把手繪字複製到 Word。
  5. 使用紀錄（history.json）：點一下載入、點兩下直接複製，可單筆刪除或全部清除。
  6. 查詢資料：cnsdata/cns_*.tsv（注音、部首、筆畫、倉頡、筆順、部件、拼音），
     取自官網 Properties.zip，由 tools/build_cnsdata.py 產生。部件字形多在私用區，
     程式啟動時把全字庫字型「只給本程式」載入（不安裝到系統）來顯示。
  7. 連線安全：完整驗證 HTTPS 憑證、不跟隨轉址，只連全字庫官網。
  8. 字庫核對：打包時把 cnsdata 各檔的 SHA-256 寫進程式（cnsdata_manifest.py），
     啟動時比對；字庫被修改或版本不符就停用離線字庫，改用線上模式。

資料存放：程式資料夾可寫入時，cache/ 與 history.json 放在程式旁；
          否則（例如安裝在 Program Files）改放 %LOCALAPPDATA%\全字庫小幫手\。

【開發與打包】
1. 安裝套件：
   pip install requests certifi pillow pywin32 pyinstaller
2. 直接執行：
   python 全字庫小幫手.py
3. 建立或更新離線字庫（約 110 MB，只在打包前需要）：
   python tools/build_cnsdata.py
   同時產生 cnsdata_manifest.py（字庫核對碼）。只想依現有 cnsdata 重算核對碼：
   python tools/build_cnsdata.py --manifest-only
   ※ 更新 cnsdata 後一定要重新打包 exe，舊程式會把新字庫當成「不符」而停用。
4. 打包（輸出 dist/全字庫小幫手/ 資料夾）：
   pyinstaller --noconfirm 全字庫小幫手.spec
   打包後把 cnsdata/ 整個複製到 dist/全字庫小幫手/ 裡。
5. 發送：程式與字庫可分成兩個 zip（最外層都是「全字庫小幫手」資料夾，解壓到同一處即可）。
   - 算出兩個 zip 的 SHA-256（certutil -hashfile 檔名 SHA256），另外從公文或內部網站
     公告；同仁收到後自行比對，避免收到被替換的檔案。
   - 程式要放在使用者自己的資料夾，或只有管理員能寫入的位置（如 Program Files）。
     不要放在大家都能寫入的共用資料夾直接執行，否則別人可以替換 _internal 裡的檔案。
================================================================================
"""

from __future__ import annotations

import os
import sys
import re
import queue
import io
import json
import base64
import hashlib
import ctypes
import ctypes.wintypes
import threading
import time
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk

# 嘗試載入第三方輔助套件（處理圖片渲染與 Windows 剪貼簿）
import importlib.util
try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
    # requests 載入約需 0.2 秒，啟動時只檢查是否存在，第一次下載時才在背景執行緒載入
    HAS_PIL_REQ = importlib.util.find_spec("requests") is not None
except ImportError:
    HAS_PIL_REQ = False

requests = None  # 延遲載入，見 _ensure_requests()


def _ensure_requests():
    """第一次連線前才載入 requests（在背景執行緒呼叫，不拖慢視窗開啟）"""
    global requests
    if requests is None:
        import requests as _requests
        requests = _requests
    return requests


def _make_secure_session():
    """
    建立「完整驗證 HTTPS 憑證」的連線階段。
    - 驗證憑證鏈與主機名稱（信任 Windows 憑證存放區 + certifi 根憑證，
      因此機關內部的 TLS 檢測 Proxy 憑證若已由 IT 部署到 Windows，也能正常使用）。
    - 只關閉 Python 3.13+ 預設的 VERIFY_X509_STRICT：全字庫官網的 TWCA 憑證鏈
      缺少 Subject Key Identifier 欄位，嚴格模式會誤擋；其餘驗證全部保留。
    - 不跟隨轉址（見 _http_get），避免被導向其他主機。
    """
    import ssl
    import certifi
    from requests.adapters import HTTPAdapter
    req = _ensure_requests()

    def build_context():
        ctx = ssl.create_default_context()            # Windows 憑證存放區
        ctx.load_verify_locations(cafile=certifi.where())
        if hasattr(ssl, "VERIFY_X509_STRICT"):
            ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        return ctx

    class GovTLSAdapter(HTTPAdapter):
        def init_poolmanager(self, *args, **kwargs):
            kwargs["ssl_context"] = build_context()
            return super().init_poolmanager(*args, **kwargs)

        def proxy_manager_for(self, *args, **kwargs):
            kwargs["ssl_context"] = build_context()
            return super().proxy_manager_for(*args, **kwargs)

    # 官網偶爾回應很慢（實測約 7% 請求的 TLS 握手超過 3 秒），自動重試：
    # 連線/讀取逾時、連線中斷、429/5xx 各最多重試 3 次，間隔 0.5、1、2 秒。
    # 只用於 GET（冪等），不跟隨轉址，憑證錯誤不重試（直接回報）。
    from urllib3.util.retry import Retry
    retry = Retry(
        total=3, connect=3, read=3, status=3, other=0,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        redirect=0, raise_on_redirect=False, raise_on_status=False,
        respect_retry_after_header=True,
    )
    session = req.Session()
    session.mount("https://", GovTLSAdapter(max_retries=retry))
    return session


_shared_session = None
_shared_session_lock = threading.Lock()


def _get_shared_session():
    """重複使用同一個連線階段（keep-alive），避免每次查詢都重新做 TLS 握手"""
    global _shared_session
    with _shared_session_lock:
        if _shared_session is None:
            _shared_session = _make_secure_session()
        return _shared_session


def _reset_shared_session():
    """連線出錯後丟棄舊連線，下次重新建立"""
    global _shared_session
    with _shared_session_lock:
        if _shared_session is not None:
            try:
                _shared_session.close()
            except Exception:
                pass
        _shared_session = None

try:
    import win32clipboard
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ==============================================================================
# 內建常見人名罕用字速查表 (供承辦人無代碼時快速檢索)
# ==============================================================================
# 代碼、部首、總筆畫皆已於 2026-11-02 逐筆核對全字庫官網字形資訊頁
# 注音取自官網 Properties.zip/CNS_phonetic.txt（20260805）；多音字把人名常用讀音排第一
COMMON_RARE_CHARS = [
    {"cns": "3-3476", "char": "堃", "radical": "土", "strokes": "11", "desc": "方+方+土", "zhuyin": ["ㄎㄨㄣ"]},
    {"cns": "3-424E", "char": "煊", "radical": "火", "strokes": "13", "desc": "火+宣", "zhuyin": ["ㄒㄩㄢ", "ㄋㄨㄢˇ"]},
    {"cns": "3-2D79", "char": "珏", "radical": "玉", "strokes": "9",  "desc": "王+玉", "zhuyin": ["ㄐㄩㄝˊ"]},
    {"cns": "3-3A38", "char": "喆", "radical": "口", "strokes": "12", "desc": "吉+吉", "zhuyin": ["ㄓㄜˊ"]},
    {"cns": "3-3A62", "char": "媖", "radical": "女", "strokes": "12", "desc": "女+英", "zhuyin": ["ㄧㄥ"]},
    {"cns": "3-4822", "char": "瑨", "radical": "玉", "strokes": "14", "desc": "王+晉", "zhuyin": ["ㄐㄧㄣˋ"]},
    {"cns": "2-4524", "char": "瑔", "radical": "玉", "strokes": "13", "desc": "王+泉", "zhuyin": ["ㄑㄩㄢˊ"]},
    {"cns": "1-742B", "char": "濬", "radical": "水", "strokes": "17", "desc": "氵+睿", "zhuyin": ["ㄐㄩㄣˋ"]},
    {"cns": "1-572A", "char": "祐", "radical": "示", "strokes": "10", "desc": "礻+右", "zhuyin": ["ㄧㄡˋ"]},
    {"cns": "2-6D44", "char": "亹", "radical": "亠", "strokes": "21", "desc": "亠+舋", "zhuyin": ["ㄨㄟˇ", "ㄇㄣˊ"]},
    {"cns": "6-222C", "char": "㐀", "radical": "一", "strokes": "5",  "desc": "卝+一", "zhuyin": ["ㄑㄧㄡ"]},
]

# 查詢方式分頁（「到官網查詢」開啟對應的全字庫官網查詢頁）
QUERY_MODES = [
    {"v": "cns", "label": "CNS 代碼",
     "hint": "支援第 1–19 字面，如 1-4442、12-6170；也可查一段範圍，如 3-4200~3-42FF，或只打前幾碼如 3-42",
     "url": "https://www.cns11643.gov.tw/search.jsp?ID=6"},
    {"v": "zhuyin", "label": "注音",
     "hint": "輸入注音，可不加聲調（列出各聲調），如 ㄒㄩㄢ；加 ˉ 表示只要一聲，如 ㄒㄩㄢˉ；也可輸入一個同音字",
     "ph": "ㄒㄩㄢ", "url": "https://www.cns11643.gov.tw/search.jsp?ID=3"},
    {"v": "pinyin", "label": "拼音",
     "hint": "輸入拼音，如 xuan、xuan1、xuān、hsüan2；不加聲調會列出各聲調，數字 1 表示只要一聲",
     "ph": "xuan", "url": "https://www.cns11643.gov.tw/search.jsp?ID=5"},
    {"v": "radical", "label": "部首/筆畫",
     "hint": "依部首、總筆畫查整個字庫，可再加注音縮小範圍，如 火部＋ㄒㄩㄢ",
     "url": "https://www.cns11643.gov.tw/searchR.jsp"},
    {"v": "part", "label": "部件",
     "hint": "輸入組成的字或部件，如 火宣、王玉，或按「部件表」點選。模糊＝包含這些部件；精確＝剛好由這些部件組成，可用 ＊（任意多個）？（任一個）",
     "ph": "火宣", "url": "https://www.cns11643.gov.tw/search.jsp?ID=7"},
    {"v": "cangjie", "label": "倉頡",
     "hint": "輸入倉頡碼或字根，如 FJMM 或 火十一一；可用 ＊ ？。A日 B月 C金 D木 E水 F火 G土 H竹 I戈 J十 K大 L中 M一 "
             "N弓 O人 P心 Q手 R口 S尸 T廿 U山 V女 W田 X難 Y卜",
     "ph": "FJMM", "url": "https://www.cns11643.gov.tw/search.jsp?ID=4"},
    {"v": "strokeseq", "label": "筆順",
     "hint": "依筆順按「橫 豎 撇 點 折」，或輸入數字 1–5；＊ 表示後面任意多筆、？ 表示任一筆。例：火字旁開頭＝點撇撇點＊",
     "ph": "4334＊", "url": "https://www.cns11643.gov.tw/search.jsp?ID=8"},
    {"v": "unicode", "label": "Unicode", "hint": "輸入 Unicode 碼位，如 U+714A 或 714A，也可直接貼上字",
     "ph": "U+714A", "url": "https://www.cns11643.gov.tw/search.jsp?ID=12"},
    {"v": "compound", "label": "複合",
     "hint": "可同時填多個條件，只列出全部符合的字（與官網「複合查詢」相同），例如 部首「火」＋注音「ㄒㄩㄢ」",
     "url": "https://www.cns11643.gov.tw/search.jsp?ID=99"},
]
Q_ALL_RADICALS, Q_ALL_STROKES = "全部部首", "全部筆畫"
Q_MAX_ROWS = 3              # 查整個字庫時最多列 3 排字卡，避免把預覽擠到畫面外

# 拼音方式：(顯示名稱, cns_pinyin.tsv 的欄位序號；None＝全部)
PINYIN_SYSTEMS = [("全部拼音方式", None), ("漢語拼音", 1), ("注音二式", 2), ("耶魯拼音", 3),
                  ("威妥瑪拼音", 4), ("通用拼音", 5)]
# 倉頡字根 → 英文字母
CJ_ROOTS = {ch: chr(ord("A") + i) for i, ch in enumerate("日月金木水火土竹戈十大中一弓人心手口尸廿山女田難卜")}
CJ_ROOTS["重"] = "Z"
# 筆順：按鈕（名稱, 插入的筆形）；輸入框裡的筆形與數字都可以
STROKE_KEYS = [("橫", "㇐"), ("豎", "㇑"), ("撇", "㇒"), ("點", "㇔"), ("折", "㇕"), ("＊", "＊"), ("？", "？")]
STROKE_DIGITS = {"㇐": "1", "一": "1", "橫": "1", "㇑": "2", "丨": "2", "豎": "2", "㇒": "3", "㇓": "3", "丿": "3",
                 "撇": "3", "㇔": "4", "丶": "4", "點": "4", "㇕": "5", "乙": "5", "折": "5", "＊": "*", "？": "?"}


class QueryInputError(Exception):
    """查詢條件看不懂或資料不足：訊息直接顯示在查詢結果下方"""

# 注音排序：依 ㄅㄆㄇ… 順序逐符號比較，同音再依聲調（一聲、ˊ、ˇ、ˋ、˙）
ZY_ORDER = "ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏㄐㄑㄒㄓㄔㄕㄖㄗㄘㄙㄧㄨㄩㄚㄛㄜㄝㄞㄟㄠㄡㄢㄣㄤㄥㄦ"
ZY_TONES = {"ˊ": 2, "ˇ": 3, "ˋ": 4, "˙": 5}


def zhuyin_sort_key(zy: str):
    if not zy:
        return ((99,), 9)
    tone = next((ZY_TONES[c] for c in zy if c in ZY_TONES), 1)
    return (tuple(ZY_ORDER.find(c) for c in zy if c not in ZY_TONES), tone)

# ==============================================================================
# 介面設計 Token（shadcn/ui slate 單色系）
# ==============================================================================
C_PAGE = "#F1F5F9"        # 視窗底色 slate-100
C_CARD = "#FFFFFF"
C_FG = "#0F172A"          # 主要文字 / 主按鈕 slate-900
C_FG_HOVER = "#1E293B"
C_MUTED = "#64748B"       # 次要文字 slate-500
C_SUBTLE = "#475569"      # slate-600
C_BORDER = "#E2E8F0"      # 細框 slate-200
C_INPUT = "#CBD5E1"       # 輸入框框線 slate-300
C_HOVER = "#F1F5F9"
C_SOFT = "#F8FAFC"        # slate-50
C_RED = "#EF4444"
C_RED_HOVER = "#DC2626"
C_RED_TEXT = "#B91C1C"
APP_VERSION = "1.0.0"
APP_AUTHOR = "71658-dev"
PROJECT_URL = "https://github.com/71658-dev/cns11643-helper"
FONT_UI = "Microsoft JhengHei UI"
FONT_MONO = "Consolas"
FONT_KAI = "標楷體"
FONT_SUNG = "新細明體"
STYLE_KAI = "標楷體 (全字庫標準楷書)"
STYLE_SUNG = "宋體 (全字庫標準明體)"

# 垂直對齊微調：數字輸入 -50 ~ 50（%，以字級為基準）；負值＝往下移，正值＝往上移。
# 2026-11-02 以 Word 16（新文件預設範本）實測「與標楷體原生字對齊」的建議值（%）：
#   12pt -22、14pt -28、16pt -27、18pt -23、20pt -27、22pt -26、24pt -22、28pt -25、36pt -22
#   不微調（0%）時，20pt 的字會比左右文字高約 5.5pt。
OFFSET_MIN, OFFSET_MAX = -50, 50

# 手繪造字：畫布座標系為 480×480（與設計稿一致），輸出 512×512 字圖
PAD_UNITS = 480
PEN_OPTIONS = [("細", 10), ("中", 18), ("粗", 28)]
PEN_DEFAULT = 18

BADGE_TONES = {
    "ok": ("#F1F5F9", "#0F172A"),
    "warn": ("#FEF3C7", "#92400E"),
    "bad": ("#FEE2E2", "#991B1B"),
}


def _blend_white(hex_color: str, ratio: float = 0.5) -> str:
    """與白色混色，用來模擬 50% 透明的停用狀態"""
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    mix = lambda c: int(c + (255 - c) * ratio)
    return f"#{mix(r):02X}{mix(g):02X}{mix(b):02X}"


class FlatButton(tk.Frame):
    """shadcn 風格平面按鈕（primary / outline / ghost / destructive）"""

    VARIANTS = {
        "primary": (C_FG, "#FFFFFF", C_FG_HOVER, C_FG),
        "outline": (C_CARD, C_FG, C_HOVER, C_BORDER),
        "ghost": (C_CARD, C_FG, C_HOVER, C_CARD),
        "destructive": (C_RED, "#FFFFFF", C_RED_HOVER, C_RED),
        "secondary": (C_HOVER, C_FG, C_BORDER, C_HOVER),
    }

    def __init__(self, parent, text, command=None, variant="primary", small=False,
                 padx=16, pady=8, font_size=10):
        bg, fg, hover, border = self.VARIANTS[variant]
        self._parent_bg = parent.cget("bg")
        super().__init__(parent, bg=border if variant != "ghost" else self._parent_bg, padx=1, pady=1)
        self._colors = (bg, fg, hover)
        self._command = command
        self._enabled = True
        if small:
            padx, pady, font_size = 12, 5, 9
        self.label = tk.Label(
            self, text=text, bg=bg, fg=fg, cursor="hand2",
            font=(FONT_UI, font_size, "bold"), padx=padx, pady=pady
        )
        self.label.pack(fill=tk.BOTH, expand=True)
        self.label.bind("<Enter>", lambda e: self._paint(hover=True))
        self.label.bind("<Leave>", lambda e: self._paint(hover=False))
        self.label.bind("<Button-1>", self._on_click)

    def _paint(self, hover=False):
        bg, fg, hover_bg = self._colors
        if not self._enabled:
            self.label.config(bg=_blend_white(bg) if bg != C_CARD else bg,
                              fg=_blend_white(fg), cursor="arrow")
        else:
            self.label.config(bg=hover_bg if hover else bg, fg=fg, cursor="hand2")

    def _on_click(self, event=None):
        if self._enabled and self._command:
            self._command()

    def set_text(self, text):
        self.label.config(text=text)

    def set_enabled(self, enabled: bool):
        self._enabled = bool(enabled)
        self._paint(hover=False)

    def set_variant(self, variant: str):
        """切換外觀（例如筆畫粗細按鈕的選取狀態）"""
        bg, fg, hover, border = self.VARIANTS[variant]
        self._colors = (bg, fg, hover)
        self.config(bg=border if variant != "ghost" else self._parent_bg)
        self._paint(hover=False)



class TransientServerError(Exception):
    """官網回傳 429/5xx 等暫時性錯誤（已自動重試仍失敗）"""

    def __init__(self, status: int):
        super().__init__(f"HTTP {status}")
        self.status = status


class OfflineCNSLibrary:
    """
    離線字庫：全字庫正楷體 TTF + CNS↔Unicode 對照表（cnsdata/，由 tools/build_cnsdata.py 產生）。
    以 512pt 在 (0,0) 繪製，實測與官網 /char/kai/512/ 字圖等同（筆畫框差 1px、形狀差 <0.4%）。
    """

    FONT_BMP, FONT_EXTB, FONT_PLUS = "TW-Kai-98_1.ttf", "TW-Kai-Ext-B-98_1.ttf", "TW-Kai-Plus-98_1.ttf"
    SIZE = 512

    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        self.version = ""
        self.code2cp = {}
        self.cp2codes = {}
        self.code2zy = {}
        self.code2rs = {}       # 代碼 → (部首序號, 總筆畫)
        self.rad2codes = {}
        self.stk2codes = {}
        self.rad_names = {}     # 部首序號 → 部首字，如 86 → 「火（灬）」
        # 倉頡、筆順、部件、拼音（核心載入後再於背景載入，見 _load_extras）
        self.code2cj, self.cj2codes = {}, {}
        self.code2seq, self.seq2codes = {}, {}
        self.code2comp = {}     # 代碼 → (部件代號, …)，依書寫順序
        self.comp_info = {}     # 部件代號 → (部件 CNS 碼, Unicode, 部件集合, 筆畫)
        self.cp2comp = {}       # 部件字形的 Unicode → 部件代號
        self.group2codes = {}   # 部件集合 → 含有該部件的代碼
        self.pinyin_rows = []   # [注音, 五種數字調拼音…, 五種調號拼音…]
        self.extras_loaded = threading.Event()
        self._fonts = {}
        self._notdef = {}
        self._loaded = threading.Event()
        self.error = ""
        self.available = all(
            os.path.isfile(os.path.join(data_dir, p))
            for p in ("cns2unicode.tsv", "fonts/" + self.FONT_BMP, "fonts/" + self.FONT_EXTB, "fonts/" + self.FONT_PLUS)
        )
        # 字庫可能與程式分開發送：先核對內容，不符就不讀（字型檔不交給 Windows／Pillow 解析）
        if self.available and not self._verify():
            self.available = False
            self.error = "離線字庫檔案與程式不符，已停用離線字庫（可能被修改或版本不同）"
        if self.available:
            threading.Thread(target=self._load_all, daemon=True).start()   # 不延遲視窗開啟
        else:
            self._loaded.set()
            self.extras_loaded.set()

    def _verify(self) -> bool:
        """
        比對 cnsdata 與打包時寫進程式的 SHA-256（cnsdata_manifest.py，由 tools/build_cnsdata.py 產生）。
        核對碼在 exe 裡，不在字庫資料夾裡，所以換掉字庫的人無法一併修改。
        程式會讀的檔（fonts/ 的字型、*.tsv）都必須列在清單上且內容相同；清單上的檔可以缺（舊版字庫）。
        開發中尚未產生 cnsdata_manifest.py 時不比對。
        """
        try:
            from cnsdata_manifest import CNSDATA_SHA256
        except ImportError:
            return True
        try:
            fonts_dir = os.path.join(self.data_dir, "fonts")
            present = {f"fonts/{n}" for n in os.listdir(fonts_dir)
                       if n.lower().endswith((".ttf", ".otf", ".ttc"))}
            present |= {n for n in os.listdir(self.data_dir) if n.lower().endswith(".tsv")}
            if present - set(CNSDATA_SHA256):
                return False
            for rel in present:
                h = hashlib.sha256()
                with open(os.path.join(self.data_dir, *rel.split("/")), "rb") as f:
                    for chunk in iter(lambda: f.read(1 << 20), b""):
                        h.update(chunk)
                if h.hexdigest() != CNSDATA_SHA256[rel]:
                    return False
            return True
        except OSError:
            return False

    def _load_all(self):
        self._load()
        self._load_extras()

    def _load_extras(self):
        """倉頡、筆順、部件、拼音（選用檔；缺檔或讀取失敗時，這幾種查詢改顯示「需要離線字庫」）"""
        try:
            if not self.available:
                return

            def rows(name):
                path = os.path.join(self.data_dir, name)
                if not os.path.isfile(path):
                    return []
                with open(path, encoding="utf-8") as f:
                    return [line.rstrip("\n").split("\t") for line in f]

            code2cj, cj2codes = {}, {}
            for r in rows("cns_cangjie.tsv"):
                if len(r) == 2:
                    code2cj[r[0]] = r[1]
                    cj2codes.setdefault(r[1], []).append(r[0])
            code2seq, seq2codes = {}, {}
            for r in rows("cns_strokes_sequence.tsv"):
                if len(r) == 2:
                    code2seq[r[0]] = r[1]
                    seq2codes.setdefault(r[1], []).append(r[0])
            comp_info, cp2comp = {}, {}
            for r in rows("cns_component_ref.tsv"):
                if len(r) == 5 and r[0].isdigit() and r[3].isdigit() and r[4].isdigit():
                    cid, cp = int(r[0]), int(r[2], 16)
                    comp_info[cid] = (r[1], cp, int(r[3]), int(r[4]))
                    cp2comp.setdefault(cp, cid)
            code2comp, group2codes = {}, {}
            for r in rows("cns_component.tsv"):
                if len(r) == 2:
                    ids = tuple(int(x) for x in r[1].split(",") if x.isdigit())
                    code2comp[r[0]] = ids
                    for cid in set(ids):
                        group2codes.setdefault(comp_info.get(cid, (0, 0, -cid, 0))[2], set()).add(r[0])
            pinyin_rows = [r for r in rows("cns_pinyin.tsv") if len(r) == 11]
            self.code2cj, self.cj2codes = code2cj, cj2codes
            self.code2seq, self.seq2codes = code2seq, seq2codes
            self.comp_info, self.cp2comp = comp_info, cp2comp
            self.code2comp, self.group2codes = code2comp, group2codes
            self.pinyin_rows = pinyin_rows
        except Exception:
            pass
        finally:
            self.extras_loaded.set()

    def comp_group(self, cid: int) -> int:
        """部件代號 → 部件集合（同一部件的不同寫法屬同一集合）"""
        info = self.comp_info.get(cid)
        return info[2] if info else -cid

    def _load(self):
        try:
            try:
                with open(os.path.join(self.data_dir, "VERSION.json"), encoding="utf-8") as f:
                    self.version = str(json.load(f).get("version", ""))[:16]
            except Exception:
                self.version = ""
            code2cp, cp2codes = {}, {}
            with open(os.path.join(self.data_dir, "cns2unicode.tsv"), encoding="utf-8") as f:
                for line in f:
                    code, _, hexcp = line.rstrip("\n").partition("\t")
                    try:
                        cp = int(hexcp, 16)
                    except ValueError:
                        continue
                    code2cp[code] = cp
                    cp2codes.setdefault(cp, []).append(code)
            self.code2cp, self.cp2codes = code2cp, cp2codes
            # 注音（選用檔；舊版 cnsdata 沒有也不影響字形功能）
            code2zy = {}
            zy_path = os.path.join(self.data_dir, "cns_phonetic.tsv")
            if os.path.isfile(zy_path):
                try:
                    with open(zy_path, encoding="utf-8") as f:
                        for line in f:
                            code, _, zy = line.rstrip("\n").partition("\t")
                            if zy:
                                code2zy.setdefault(code, []).append(zy)
                except Exception:
                    code2zy = {}
            self.code2zy = code2zy
            # 部首與總筆畫（選用檔）
            code2rs, rad2codes, stk2codes, rad_names = {}, {}, {}, {}
            try:
                with open(os.path.join(self.data_dir, "cns_radicals.tsv"), encoding="utf-8") as f:
                    for line in f:
                        num, _, name = line.rstrip("\n").partition("\t")
                        if num.isdigit() and name:
                            rad_names[int(num)] = name
                with open(os.path.join(self.data_dir, "cns_radical_stroke.tsv"), encoding="utf-8") as f:
                    for line in f:
                        p = line.rstrip("\n").split("\t")
                        if len(p) == 3 and p[1].isdigit() and p[2].isdigit():
                            r, s = int(p[1]), int(p[2])
                            code2rs[p[0]] = (r, s)
                            rad2codes.setdefault(r, []).append(p[0])
                            stk2codes.setdefault(s, []).append(p[0])
            except OSError:
                code2rs, rad2codes, stk2codes, rad_names = {}, {}, {}, {}
            if len(rad_names) == 214 and code2rs:
                self.code2rs, self.rad2codes, self.stk2codes, self.rad_names = code2rs, rad2codes, stk2codes, rad_names
        except Exception as e:
            self.available = False
            self.error = f"離線字庫讀取失敗：{e}"
        finally:
            self._loaded.set()

    def ready(self, timeout: float = 3.0) -> bool:
        self._loaded.wait(timeout)
        return self.available and bool(self.code2cp)

    def has(self, code: str) -> bool:
        return self.ready() and code in self.code2cp

    def char_for(self, code: str):
        cp = self.code2cp.get(code) if self.ready() else None
        return chr(cp) if cp is not None else None

    def codes_for(self, char: str) -> list:
        if not self.ready() or len(char) != 1:
            return []
        return list(self.cp2codes.get(ord(char), []))

    def thumb(self, code: str, size: int):
        """直接以小字級繪製縮圖（比 512 繪製再縮小快很多）；缺字回傳 None"""
        if not self.has(code):
            return None
        cp = self.code2cp[code]
        name = self._font_name(cp)
        try:
            key = (name, size)
            font = self._fonts.get(key)
            if font is None:
                font = ImageFont.truetype(os.path.join(self.data_dir, "fonts", name), size)
                self._fonts[key] = font

            def draw(ch):
                img = Image.new("L", (size, size), 255)
                ImageDraw.Draw(img).text((0, 0), ch, font=font, fill=0)
                return img

            img = draw(chr(cp))
            if key not in self._notdef:
                self._notdef[key] = draw("͸").tobytes()
            if img.tobytes() == self._notdef[key] or img.getextrema()[0] > 200:
                return None
            return img
        except Exception:
            return None

    def has_radicals(self) -> bool:
        """已載入部首/筆畫資料（不等待）"""
        return self._loaded.is_set() and bool(self.code2rs)

    def zhuyin_for(self, code: str) -> list:
        """官網注音（一字多音時依官網順序）；未安裝或無資料時回傳 []"""
        return list(self.code2zy.get(code, [])) if self.ready(0) else []

    def _font_name(self, cp: int) -> str:
        if cp >= 0xF0000:
            return self.FONT_PLUS
        if cp >= 0x20000:
            return self.FONT_EXTB
        return self.FONT_BMP

    def _draw(self, name: str, ch: str) -> Image.Image:
        font = self._fonts.get(name)
        if font is None:
            font = ImageFont.truetype(os.path.join(self.data_dir, "fonts", name), self.SIZE)
            self._fonts[name] = font
        img = Image.new("L", (self.SIZE, self.SIZE), 255)
        ImageDraw.Draw(img).text((0, 0), ch, font=font, fill=0)
        return img

    def render(self, code: str):
        """回傳 512×512 白底黑字（L 模式）；對照表沒有或字型缺字時回傳 None"""
        if not self.has(code):
            return None
        cp = self.code2cp[code]
        name = self._font_name(cp)
        try:
            img = self._draw(name, chr(cp))
            # 字型缺字時會畫出 .notdef（方框或空白），與「確定不存在的碼位」繪圖結果相同
            if name not in self._notdef:
                self._notdef[name] = self._draw(name, "͸").tobytes()
            if img.tobytes() == self._notdef[name] or img.getextrema()[0] > 200:
                return None
            return img
        except Exception:
            return None


class HandwritingRecognizer:
    """
    Windows 內建「Microsoft 中文(繁體)手寫辨識器」（Tablet PC COM API，msinkaut）。
    完全離線：手寫筆跡不會離開這台電腦。
    其他電腦需安裝「中文(台灣)」語言的「手寫」功能，否則 available 為 False。
    """

    HIMETRIC_PER_UNIT = 10      # 畫板 480 座標 → HIMETRIC（0.01mm），整格約 48mm
    MAX_CANDIDATES = 6

    def __init__(self):
        self._recognizer = None
        self._tried = False
        self.error = ""

    @property
    def available(self) -> bool:
        self._ensure()
        return self._recognizer is not None

    def _ensure(self):
        if self._tried:
            return
        self._tried = True
        try:
            import win32com.client
            recs = win32com.client.Dispatch("msinkaut.InkRecognizers")
            for i in range(recs.Count):
                rec = recs.Item(i)
                name = str(rec.Name)
                if "Gesture" in name:
                    continue
                if "中文(繁體)" in name or "Traditional" in name or "繁" in name:
                    self._recognizer = rec
                    break
            if self._recognizer is None:
                self.error = "此電腦未安裝中文(台灣)手寫辨識"
        except Exception as e:
            self.error = f"無法啟用手寫辨識：{e}"

    @staticmethod
    def _is_cjk_char(s: str) -> bool:
        """只保留單一漢字（排除 + f . □ 等符號、英數）"""
        if len(s) != 1:
            return False
        cp = ord(s)
        return (0x3400 <= cp <= 0x4DBF or 0x4E00 <= cp <= 0x9FFF or 0xF900 <= cp <= 0xFAFF
                or 0x20000 <= cp <= 0x323AF)

    def recognize(self, strokes) -> list:
        """strokes：[{"pts": [(x, y), ...]}, ...]（480 座標）→ 候選漢字（依可能性排序）"""
        self._ensure()
        if self._recognizer is None or not strokes:
            return []
        import win32com.client
        import pythoncom
        from win32com.client import VARIANT
        k = self.HIMETRIC_PER_UNIT
        ink = win32com.client.Dispatch("msinkaut.InkObject")
        for stroke in strokes:
            pts = stroke["pts"]
            if len(pts) == 1:                   # 單點：補一個極短線段，辨識器才收得到
                pts = [pts[0], (pts[0][0] + 1, pts[0][1] + 1)]
            flat = [int(round(v * k)) for p in pts for v in p]
            ink.CreateStroke(VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, flat),
                             VARIANT(pythoncom.VT_EMPTY, None))
        ctx = self._recognizer.CreateRecognizerContext()
        ctx.Strokes = ink.Strokes
        ctx.EndInkInput()
        result, _status = ctx.Recognize()
        if result is None:
            return []
        alts = result.AlternatesFromSelection(0, -1, 20)
        out = []
        for i in range(alts.Count):
            s = str(alts.Item(i).String).strip()
            if self._is_cjk_char(s) and s not in out:
                out.append(s)
            if len(out) >= self.MAX_CANDIDATES:
                break
        return out


class CNSRareCharApp:
    """全字庫小幫手主程式類別"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("全字庫小幫手")
        self.root.configure(bg=C_PAGE)
        # 依螢幕 DPI 換算像素尺寸（字型以 pt 計已自動縮放，畫布與間距需手動換算）
        try:
            self.scale = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        except Exception:
            self.scale = 1.0
        # 依螢幕「可用區域」（扣除工作列）決定視窗大小，避免狀態列被工作列遮住
        work_top, work_h = 0, self.root.winfo_screenheight() - self.px(48)
        if os.name == "nt":
            try:
                rect = ctypes.wintypes.RECT()
                if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                    work_top, work_h = rect.top, rect.bottom - rect.top
            except Exception:
                pass
        avail_h = work_h - self.px(40)   # 預留標題列高度
        win_w, win_h = self.px(800), min(self.px(880), avail_h)
        self.root.geometry(f"{win_w}x{win_h}+{self.px(40)}+{work_top + self.px(8)}")
        self.root.minsize(self.px(700), min(self.px(640), avail_h))

        # 取得正確的本機目錄 (完美支援 PyInstaller 打包後的 .exe 所在路徑)
        if getattr(sys, "frozen", False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))

        # 資料（快取、紀錄）預設放在程式旁；若程式安裝在使用者無寫入權限的位置
        # （例如 Program Files，可防止他人替換程式檔），改存到使用者自己的 AppData。
        self.data_dir = self._pick_data_dir()
        self.cache_dir = os.path.join(self.data_dir, "cache")
        self.history_file = os.path.join(self.data_dir, "history.json")
        os.makedirs(self.cache_dir, exist_ok=True)

        # 離線字庫（與 exe 同層的 cnsdata/；不存在時自動退回純線上模式）
        self.offline = OfflineCNSLibrary(os.path.join(self.base_dir, "cnsdata"))
        # 把全字庫字型「只給這支程式」使用（不安裝到系統）：部件多半是私用區字元，一般字型顯示成方框
        self.font_kai_family, self.font_plus_family = FONT_KAI, FONT_KAI
        if self.offline.available and os.name == "nt":
            try:
                add = ctypes.windll.gdi32.AddFontResourceExW
                fonts_dir = os.path.join(self.offline.data_dir, "fonts")
                if add(os.path.join(fonts_dir, OfflineCNSLibrary.FONT_BMP), 0x10, 0):        # FR_PRIVATE
                    self.font_kai_family = "TW-Kai"
                add(os.path.join(fonts_dir, OfflineCNSLibrary.FONT_EXTB), 0x10, 0)
                if add(os.path.join(fonts_dir, OfflineCNSLibrary.FONT_PLUS), 0x10, 0):
                    self.font_plus_family = "TW-Kai-Plus"
            except Exception:
                pass
        self._offline_thumbs = {}

        # 狀態變數初始化
        self.current_cns_code = None      # 標準化後的 CNS 代碼，如 "1-4442" 或 "12-6170"
        self.current_pil_image = None     # 目前載入的高解析度字圖 (PIL.Image)
        self.current_tk_image = None      # 供預覽畫布顯示的圖片物件
        self.current_source = None        # cache / online / sung-sub / fallback
        self.is_fetching = False          # 防止重複點擊下載鎖定旗標
        self.current_is_fallback = False  # 目前字圖是否為離線備援圖（非全字庫字形，禁止複製）
        self.current_style_is_sung = False  # 目前字圖載入時「選單所選」的字體（複製前比對用）
        self.pending_style_is_sung = False  # 下載中那一筆所選的字體
        self.copy_after_fetch = False     # 雙擊歷史紀錄：待字形載入完成後才複製
        self.ui_queue = queue.Queue()     # 背景執行緒 -> 主執行緒 的回呼佇列
        self._history_thumbs = []         # 歷史紀錄縮圖參照（避免被回收）
        self._copied_reset_job = None
        self.history_data = self.load_history()

        # 綁定變數
        self.code_var = tk.StringVar(value="3-3476")  # 堃
        self.font_style_var = tk.StringVar(value=STYLE_KAI)
        self.font_size_var = tk.StringVar(value="20")
        self.baseline_offset_var = tk.StringVar(value="0")  # 垂直對齊微調（%，-50 ~ 50，預設 0）
        self.status_var = tk.StringVar(value="就緒")

        # 建立 UI 介面
        self.setup_styles()
        self.build_ui()
        self.set_status("就緒")

        # 輸入框內容變更時，同步清除錯誤提示並更新複製按鈕狀態
        self.code_var.trace_add("write", lambda *a: self._on_code_changed())

        # 檢查環境套件狀態
        self.check_environment()

        # 預先載入歷史紀錄清單
        self.refresh_history_tree()
        self.update_copy_state()

        # 啟動主執行緒佇列輪詢（背景執行緒不直接操作 Tk）
        self.root.after(100, self._poll_ui_queue)
        self.root.after(200, self._update_version_label)

    def _update_version_label(self):
        """離線字庫在背景載入，完成後顯示版本"""
        if not self.offline._loaded.is_set():
            self.root.after(200, self._update_version_label)
            return
        if self.offline.ready(0):
            self.version_lbl.config(text=f"離線字庫 {self.offline.version or ''}".strip(), fg=C_MUTED)
            self.refresh_history_tree()          # 更新「離線」標示、縮圖與注音
            self._update_zhuyin_line()
            self._cand_render()
            self._fill_radical_options()         # 換成整庫 214 個部首
            self._wait_extras()                  # 倉頡、筆順、部件、拼音載入後更新
            self.refresh_query_results()         # 縮圖與整庫查詢
        elif self.offline.error:
            self.version_lbl.config(text="線上模式（離線字庫已停用）", fg=C_RED)
            self.set_status(self.offline.error, is_error=True)
        else:
            self.version_lbl.config(text="線上模式（未安裝離線字庫）", fg=C_MUTED)

    def check_cnsdata_update(self):
        """比對官網 release.txt 的字型版本；只提示，不自動下載（約 58MB）"""
        if not self.offline.available:
            self.set_status("未安裝離線字庫，目前每次都會連線官網下載字形。", is_error=True)
            return
        self.update_btn.set_enabled(False)
        self.set_status("正在檢查全字庫字庫版本…", is_loading=True)
        local = self.offline.version

        def worker():
            try:
                _ensure_requests()
                st, body = self._http_get(_get_shared_session(), f"{self.CNS_BASE}/opendata/release.txt",
                                          max_bytes=256 * 1024)
                if st != 200:
                    raise RuntimeError(f"HTTP {st}")
                m = re.search(r"Fonts_Kai\.zip\s*\n\s*版本：\s*(\d{8})", body.decode("utf-8-sig", errors="replace"))
                remote = m.group(1) if m else ""
                self.ui_queue.put(lambda: self._update_check_done(local, remote, None))
            except Exception as e:
                msg = str(e)
                self.ui_queue.put(lambda: self._update_check_done(local, "", msg))

        threading.Thread(target=worker, daemon=True).start()

    def _update_check_done(self, local: str, remote: str, error):
        self.update_btn.set_enabled(True)
        if error:
            self.set_status(f"無法檢查字庫更新：{error}", is_error=True)
        elif not remote:
            self.set_status("無法辨識官網的字庫版本資訊。", is_error=True)
        elif remote > local:
            self.set_status(f"官網有新版字庫 {remote}（目前 {local}）。請洽資訊人員以 tools/build_cnsdata.py 更新 cnsdata。",
                            is_error=True)
        else:
            self.set_status(f"離線字庫已是最新版（{local}）。")

    def _pick_data_dir(self) -> str:
        """程式資料夾可寫入就用它，否則用 %LOCALAPPDATA%\\全字庫小幫手"""
        try:
            os.makedirs(os.path.join(self.base_dir, "cache"), exist_ok=True)
            probe = os.path.join(self.base_dir, "cache", ".write_test")
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
            return self.base_dir
        except OSError:
            appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
            new_dir = os.path.join(appdata, "全字庫小幫手")
            old_dir = os.path.join(appdata, "全字庫罕用字小幫手")   # 改名前的資料夾
            if os.path.isdir(old_dir) and not os.path.exists(new_dir):
                try:
                    os.rename(old_dir, new_dir)                  # 沿用舊的快取與使用紀錄
                except OSError:
                    return old_dir
            return new_dir

    def px(self, n: float) -> int:
        """把設計稿的 CSS 像素換算成實際螢幕像素"""
        return int(round(n * self.scale))

    def setup_styles(self):
        """設定 ttk 元件為 shadcn 風格（細框、單色、無立體感）"""
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        self.root.option_add("*TCombobox*Listbox.font", (FONT_UI, 10))
        self.root.option_add("*TCombobox*Listbox.selectBackground", C_HOVER)
        self.root.option_add("*TCombobox*Listbox.selectForeground", C_FG)

        style.configure(
            "TCombobox", padding=self.px(6), foreground=C_FG,
            fieldbackground=C_CARD, background=C_CARD,
            bordercolor=C_INPUT, lightcolor=C_INPUT, darkcolor=C_INPUT,
            arrowcolor=C_MUTED, arrowsize=self.px(12),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", C_CARD)],
            selectbackground=[("readonly", C_CARD)],
            selectforeground=[("readonly", C_FG)],
            bordercolor=[("focus", C_FG)],
        )

        style.configure(
            "History.Treeview", background=C_CARD, fieldbackground=C_CARD, foreground=C_FG,
            font=(FONT_UI, 10), rowheight=self.px(48), borderwidth=0, relief="flat",   # 代碼欄兩行：代碼＋注音
        )
        style.layout("History.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.map(
            "History.Treeview",
            background=[("selected", C_HOVER)],
            foreground=[("selected", C_FG)],
        )
        style.configure(
            "History.Treeview.Heading", background=C_CARD, foreground=C_MUTED,
            font=(FONT_UI, 9, "bold"), relief="flat", borderwidth=0,
            bordercolor=C_BORDER, lightcolor=C_CARD, darkcolor=C_BORDER,
            padding=(self.px(8), self.px(8)),
        )
        style.map("History.Treeview.Heading", background=[("active", C_CARD)])
        style.configure(
            "Slim.Vertical.TScrollbar", background=C_BORDER, troughcolor=C_CARD,
            bordercolor=C_CARD, lightcolor=C_BORDER, darkcolor=C_BORDER,
            arrowcolor=C_MUTED, gripcount=0,
        )

    # --------------------------------------------------------------------------
    # 小型元件工廠
    # --------------------------------------------------------------------------
    def _label(self, parent, text="", size=10, weight="normal", fg=C_FG, font=FONT_UI, **kw):
        return tk.Label(parent, text=text, bg=parent.cget("bg"), fg=fg,
                        font=(font, size, weight), **kw)

    def _bordered(self, parent, color=C_BORDER, bg=C_CARD):
        """1px 細框容器：外層填框線色，內層留 1px"""
        outer = tk.Frame(parent, bg=color, padx=1, pady=1)
        inner = tk.Frame(outer, bg=bg)
        inner.pack(fill=tk.BOTH, expand=True)
        return outer, inner

    def _hline(self, parent, color=C_BORDER):
        line = tk.Frame(parent, bg=color, height=1)
        line.pack(fill=tk.X)
        return line

    def build_ui(self):
        """依 Claude Design「全字庫小幫手」稿建構介面"""
        outer = tk.Frame(self.root, bg=C_PAGE)
        outer.pack(fill=tk.BOTH, expand=True, padx=self.px(16), pady=self.px(16))
        card_border, card = self._bordered(outer)
        card_border.pack(fill=tk.BOTH, expand=True)

        # ------------------------------------------------------------------
        # 頁首：標題 + 副標 + 官網按鈕
        # ------------------------------------------------------------------
        header = tk.Frame(card, bg=C_CARD, padx=self.px(24), pady=self.px(16))
        header.pack(fill=tk.X)
        title_box = tk.Frame(header, bg=C_CARD)
        title_box.pack(side=tk.LEFT)
        self._label(title_box, "全字庫小幫手", size=15, weight="bold").pack(anchor="w")
        self._label(title_box, "免造字檔，複製後在 Word 直接貼上列印", size=10, fg=C_MUTED).pack(anchor="w")
        FlatButton(
            header, "全字庫官網查碼 ↗", variant="outline", small=True,
            command=lambda: webbrowser.open("https://www.cns11643.gov.tw/")
        ).pack(side=tk.RIGHT)
        self.version_lbl = self._label(header, "", size=9, weight="bold", fg=C_MUTED)
        self.version_lbl.pack(side=tk.RIGHT, padx=(0, self.px(12)))
        self._hline(card)

        # 頁尾狀態列（先 pack 在底部，確保視窗縮小時不被擠掉）
        footer = tk.Frame(card, bg=C_CARD)
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        self._hline(footer)
        about = self._label(footer, "關於與授權", size=9, weight="bold", fg=C_SUBTLE, cursor="hand2")
        about.pack(side=tk.RIGHT, padx=(0, self.px(24)))
        about.bind("<Button-1>", lambda e: self.show_about())
        self.status_lbl = tk.Label(
            footer, textvariable=self.status_var, bg=C_CARD, fg=C_SUBTLE,
            font=(FONT_UI, 9, "bold"), anchor="w", justify="left",
            padx=self.px(24), pady=self.px(8)
        )
        self.status_lbl.pack(fill=tk.X)
        self.status_lbl.bind("<Configure>", lambda e: self.status_lbl.config(wraplength=max(200, e.width - self.px(48))))

        # 主內容區可捲動（加入手繪造字後，內容高度超過一般螢幕）
        scroll_wrap = tk.Frame(card, bg=C_CARD)
        scroll_wrap.pack(fill=tk.BOTH, expand=True)
        self.main_canvas = tk.Canvas(scroll_wrap, bg=C_CARD, highlightthickness=0, bd=0)
        self.main_vsb = ttk.Scrollbar(scroll_wrap, orient=tk.VERTICAL, command=self.main_canvas.yview,
                                      style="Slim.Vertical.TScrollbar")
        self.main_canvas.configure(yscrollcommand=self.main_vsb.set)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main = tk.Frame(self.main_canvas, bg=C_CARD, padx=self.px(24), pady=self.px(20))
        main_win = self.main_canvas.create_window((0, 0), window=main, anchor="nw")

        def on_main_resize(event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
            # 內容放得下就隱藏捲軸
            needs = main.winfo_reqheight() > self.main_canvas.winfo_height()
            if needs and not self.main_vsb.winfo_ismapped():
                self.main_vsb.pack(side=tk.RIGHT, fill=tk.Y, before=self.main_canvas)
            elif not needs and self.main_vsb.winfo_ismapped():
                self.main_vsb.pack_forget()
                self.main_canvas.yview_moveto(0)

        main.bind("<Configure>", on_main_resize)
        self.main_canvas.bind("<Configure>", lambda e: (self.main_canvas.itemconfigure(main_win, width=e.width),
                                                        on_main_resize()))
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)

        # ------------------------------------------------------------------
        # 區塊 1：代碼輸入與設定
        # ------------------------------------------------------------------
        sec_input = tk.Frame(main, bg=C_CARD)
        sec_input.pack(fill=tk.X)

        # 查詢方式（分頁）：CNS 代碼／注音／拼音／部首筆畫／部件／倉頡／筆順／Unicode／複合
        self.query_mode = "cns"
        self.q_text_var = tk.StringVar()
        self.q_radical_var = tk.StringVar(value=Q_ALL_RADICALS)
        self.q_strokes_var = tk.StringVar(value=Q_ALL_STROKES)
        self.q_rzy_var = tk.StringVar()
        self.q_pinyin_sys_var = tk.StringVar(value=PINYIN_SYSTEMS[0][0])
        self.q_part_exact = False
        self._radical_widgets = []
        self._q_job = None
        self._q_results = []
        self._q_thumbs = {}
        self._q_cols = 0
        font_part = (self.font_plus_family, 14)      # 部件輸入框：可顯示私用區部件字形
        font_stroke = (self.font_kai_family, 14)     # 筆順輸入框：可顯示 ㇐㇑㇒㇔㇕

        self._label(sec_input, "查詢方式", size=10, weight="bold").pack(anchor="w")
        tabs = tk.Frame(sec_input, bg=C_HOVER, padx=self.px(4), pady=self.px(4))
        tabs.pack(anchor="w", pady=(self.px(8), 0))
        self.mode_tabs = {}
        for mode in QUERY_MODES:
            tab = tk.Label(tabs, text=mode["label"], font=(FONT_UI, 10, "bold"), cursor="hand2",
                           padx=self.px(10), pady=self.px(3))
            tab.pack(side=tk.LEFT, padx=(0 if not self.mode_tabs else self.px(2), 0))
            tab.bind("<Button-1>", lambda e, v=mode["v"]: self.set_query_mode(v))
            self.mode_tabs[mode["v"]] = tab
        self.mode_hint_lbl = self._label(sec_input, "", size=9, weight="bold", fg=C_MUTED, anchor="w", justify="left")
        self.mode_hint_lbl.pack(fill=tk.X, pady=(self.px(8), 0))
        self.mode_hint_lbl.bind("<Configure>", lambda e: self.mode_hint_lbl.config(wraplength=max(200, e.width)))

        self.input_row = tk.Frame(sec_input, bg=C_CARD)
        self.input_row.pack(fill=tk.X, pady=(self.px(8), 0))
        self.query_box = tk.Frame(self.input_row, bg=C_CARD)
        self.query_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # (1) CNS 代碼（可輸入範圍）
        self.code_entry_border, entry_inner = self._bordered(self.query_box, color=C_INPUT)
        self.code_entry = tk.Entry(
            entry_inner, textvariable=self.code_var, font=(FONT_MONO, 13, "bold"),
            bd=0, relief="flat", bg=C_CARD, fg=C_FG, insertbackground=C_FG,
        )
        self.code_entry.pack(fill=tk.BOTH, expand=True, padx=self.px(12), pady=self.px(7))
        self.code_entry.bind("<Return>", lambda event: self.query_load())
        self.code_entry.bind("<FocusIn>", lambda e: self.code_entry_border.config(bg=C_FG))
        self.code_entry.bind("<FocusOut>", lambda e: self.code_entry_border.config(bg=C_INPUT))

        # (2) 注音／拼音／部件／倉頡／筆順／Unicode 共用的文字輸入框，右側依分頁附加按鈕
        self.q_entry_border, self.q_entry = self._entry(self.query_box, self.q_text_var, (FONT_UI, 12),
                                                        on_return=self.query_load, padx=12, pady=7)
        self.q_text_var.trace_add("write", lambda *a: self._schedule_query())
        self.q_entry_fonts = {"part": font_part, "strokeseq": font_stroke}

        self.q_extra = {}
        # 拼音方式
        box = tk.Frame(self.query_box, bg=C_CARD)
        combo = ttk.Combobox(box, textvariable=self.q_pinyin_sys_var, values=[n for n, _ in PINYIN_SYSTEMS],
                             state="readonly", font=(FONT_UI, 10), width=11)
        combo.pack(fill=tk.Y, expand=True, ipady=self.px(3))
        combo.bind("<<ComboboxSelected>>", lambda e: (combo.selection_clear(), self.refresh_query_results()))
        self.q_extra["pinyin"] = box
        # 部件：模糊／精確、部件表
        box = tk.Frame(self.query_box, bg=C_CARD)
        self.part_mode_btns = {
            False: FlatButton(box, "模糊", variant="primary", small=True, command=lambda: self._set_part_exact(False)),
            True: FlatButton(box, "精確", variant="outline", small=True, command=lambda: self._set_part_exact(True)),
        }
        self.part_mode_btns[False].pack(side=tk.LEFT, fill=tk.Y)
        self.part_mode_btns[True].pack(side=tk.LEFT, fill=tk.Y)
        FlatButton(box, "部件表", variant="outline", small=True,
                   command=lambda: self.open_component_picker(self.q_entry)).pack(side=tk.LEFT, fill=tk.Y,
                                                                                padx=(self.px(6), 0))
        self.q_extra["part"] = box
        # 筆順：橫 豎 撇 點 折 ＊ ？
        box = tk.Frame(self.query_box, bg=C_CARD)
        for name, glyph in STROKE_KEYS:
            FlatButton(box, name, variant="outline", small=True, padx=8,
                       command=lambda g=glyph: self._insert_text(self.q_entry, g)).pack(side=tk.LEFT, fill=tk.Y,
                                                                                         padx=(self.px(3), 0))
        self.q_extra["strokeseq"] = box

        # (3) 部首＋筆畫＋注音（選填）
        self.q_radical_box = tk.Frame(self.query_box, bg=C_CARD)
        for col in range(3):
            self.q_radical_box.columnconfigure(col, weight=1, uniform="rs")
        rc, sc = self._radical_pair(self.q_radical_box, self.q_radical_box, self.q_radical_var, self.q_strokes_var)
        rc.grid(row=0, column=0, sticky="nsew", ipady=self.px(3))
        sc.grid(row=0, column=1, sticky="nsew", padx=(self.px(8), 0), ipady=self.px(3))
        self.q_radical_combo, self.q_strokes_combo = rc, sc
        rzy_border, self.q_rzy_entry = self._entry(self.q_radical_box, self.q_rzy_var, (FONT_UI, 11),
                                                   placeholder="注音（選填）", on_return=self.query_load)
        rzy_border.grid(row=0, column=2, sticky="nsew", padx=(self.px(8), 0))
        self.q_rzy_var.trace_add("write", lambda *a: self._schedule_query())

        # (4) 複合查詢：所有條件一起篩選
        self.q_compound_box = tk.Frame(self.query_box, bg=C_CARD)
        for col in range(4):
            self.q_compound_box.columnconfigure(col, weight=1, uniform="cp")
        self.c_rad_var = tk.StringVar(value=Q_ALL_RADICALS)
        self.c_stk_var = tk.StringVar(value=Q_ALL_STROKES)
        self.c_vars = {k: tk.StringVar() for k in ("zhuyin", "pinyin", "cangjie", "strokeseq", "part", "cns")}

        def cell(r, c, title):
            f = tk.Frame(self.q_compound_box, bg=C_CARD)
            f.grid(row=r, column=c, sticky="nsew", padx=(0 if c == 0 else self.px(8), 0),
                   pady=(0 if r == 0 else self.px(8), 0))
            self._label(f, title, size=9, weight="bold", fg=C_MUTED).pack(anchor="w", pady=(0, self.px(3)))
            return f

        f_rad, f_stk = cell(0, 0, "部首"), cell(0, 1, "總筆畫")
        rc, sc = self._radical_pair(f_rad, f_stk, self.c_rad_var, self.c_stk_var)
        rc.pack(fill=tk.X, ipady=self.px(3))
        sc.pack(fill=tk.X, ipady=self.px(3))
        self.c_part_entry = None
        for (r, c, title, key, ph, font) in (
                (0, 2, "注音", "zhuyin", "ㄒㄩㄢ", (FONT_UI, 11)),
                (0, 3, "拼音", "pinyin", "xuan", (FONT_UI, 11)),
                (1, 0, "倉頡", "cangjie", "FJMM", (FONT_UI, 11)),
                (1, 1, "筆順（1–5）", "strokeseq", "4334＊", (self.font_kai_family, 12)),
                (1, 2, "部件（包含）", "part", "火宣", (self.font_plus_family, 12)),
                (1, 3, "CNS 範圍", "cns", "3-4200~3-42FF", (FONT_UI, 11))):
            f = cell(r, c, title)
            border, entry = self._entry(f, self.c_vars[key], font, placeholder=ph, on_return=self.query_load)
            border.pack(fill=tk.X)
            self.c_vars[key].trace_add("write", lambda *a: self._schedule_query())
            if key == "part":
                self.c_part_entry = entry
                link = self._label(f, "部件表…", size=9, weight="bold", fg=C_SUBTLE, cursor="hand2")
                link.pack(anchor="e", pady=(self.px(2), 0))
                link.bind("<Button-1>", lambda e: self.open_component_picker(self.c_part_entry))
        self._fill_radical_options()

        self.fetch_btn = FlatButton(self.input_row, "載入", command=self.query_load)
        self.official_btn = FlatButton(self.input_row, "到官網查詢 ↗", variant="outline", command=self.open_official_query)
        self._pack_query_buttons(compact=False)

        # 查詢結果：字卡格狀排列，點選即載入
        self.q_results_wrap = tk.Frame(sec_input, bg=C_CARD)
        self.q_results_wrap.pack(fill=tk.X, pady=(self.px(12), 0))
        self.q_grid = tk.Frame(self.q_results_wrap, bg=C_CARD)
        self.q_grid.pack(fill=tk.X)
        self.q_grid.bind("<Configure>", lambda e: self._layout_query_cards())
        self.q_note_lbl = self._label(self.q_results_wrap, "", size=10, fg=C_MUTED, anchor="w", justify="left")
        self.q_note_lbl.bind("<Configure>", lambda e: self.q_note_lbl.config(wraplength=max(200, e.width)))

        self.input_error_lbl = self._label(sec_input, "", size=9, weight="bold", fg=C_RED, anchor="w")
        self.set_query_mode("cns", focus=False)

        self.settings_row = tk.Frame(sec_input, bg=C_CARD)
        self.settings_row.pack(fill=tk.X, pady=(self.px(14), 0))
        for i in range(3):
            self.settings_row.columnconfigure(i, weight=1, uniform="settings")

        def setting(col, title, var, values, on_change):
            box = tk.Frame(self.settings_row, bg=C_CARD)
            box.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else self.px(6), 0 if col == 2 else self.px(6)))
            self._label(box, title, size=9, weight="bold", fg=C_MUTED).pack(anchor="w", pady=(0, self.px(4)))
            combo = ttk.Combobox(box, textvariable=var, values=values, state="readonly", font=(FONT_UI, 10))
            combo.pack(fill=tk.X)
            combo.bind("<<ComboboxSelected>>", lambda e: (combo.selection_clear(), on_change()))
            return combo

        setting(0, "字體", self.font_style_var, [STYLE_KAI, STYLE_SUNG], self.fetch_character)
        setting(1, "Word 字級（pt）", self.font_size_var,
                ["12", "14", "16", "18", "20", "22", "24", "28", "36"], self._on_size_or_offset_changed)
        # 垂直對齊微調：數字輸入（-50 ~ 50），右側顯示 %
        offset_box = tk.Frame(self.settings_row, bg=C_CARD)
        offset_box.grid(row=0, column=2, sticky="ew", padx=(self.px(6), 0))
        self._label(offset_box, "垂直對齊微調", size=9, weight="bold", fg=C_MUTED).pack(anchor="w", pady=(0, self.px(4)))
        ob, oi = self._bordered(offset_box, color=C_INPUT)
        ob.pack(fill=tk.X)
        tk.Label(oi, text="%", bg=C_CARD, fg=C_MUTED, font=(FONT_UI, 10)).pack(side=tk.RIGHT, padx=(0, self.px(10)))
        vcmd = (self.root.register(lambda v: bool(re.fullmatch(r"-?\d{0,2}", v))), "%P")
        self.offset_spin = tk.Spinbox(
            oi, from_=OFFSET_MIN, to=OFFSET_MAX, increment=1, textvariable=self.baseline_offset_var,
            font=(FONT_UI, 10), bd=0, relief="flat", bg=C_CARD, fg=C_FG, insertbackground=C_FG,
            buttonbackground=C_CARD, validate="key", validatecommand=vcmd, width=6,
            command=self._on_size_or_offset_changed,
        )
        self.offset_spin.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(self.px(10), self.px(4)), pady=self.px(6))
        self.offset_spin.bind("<FocusOut>", lambda e: self._normalize_offset())
        self.offset_spin.bind("<Return>", lambda e: self._normalize_offset())

        # ------------------------------------------------------------------
        # 區塊 2：字形預覽與複製
        # ------------------------------------------------------------------
        preview_border, preview = self._bordered(main)
        preview_border.pack(fill=tk.X, pady=(self.px(20), 0))
        preview_body = tk.Frame(preview, bg=C_CARD, padx=self.px(20), pady=self.px(20))
        preview_body.pack(fill=tk.X)

        left = tk.Frame(preview_body, bg=C_CARD)
        left.pack(side=tk.LEFT, anchor="n")
        self.preview_size = self.px(168)
        canvas_border, canvas_inner = self._bordered(left)
        canvas_border.pack()
        self.preview_canvas = tk.Canvas(
            canvas_inner, width=self.preview_size, height=self.preview_size,
            bg=C_CARD, bd=0, highlightthickness=0
        )
        self.preview_canvas.pack()

        # 右欄（垂直置中）：CNS 11643 小標 → 大字代碼 → 字體・字級 + 狀態徽章 → 複製按鈕
        right_wrap = tk.Frame(preview_body, bg=C_CARD)
        right_wrap.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(self.px(24), 0))
        right = tk.Frame(right_wrap, bg=C_CARD)
        right.pack(fill=tk.X, expand=True)

        self._label(right, "CNS 11643", size=9, weight="bold", fg=C_MUTED).pack(anchor="w")
        self.cns_code_lbl = self._label(right, "—", size=22, weight="bold", font=FONT_MONO)
        self.cns_code_lbl.pack(anchor="w", pady=(self.px(2), 0))
        self.zhuyin_lbl = self._label(right, "", size=12, weight="bold", fg=C_SUBTLE)   # 有字時才顯示

        meta_row = tk.Frame(right, bg=C_CARD)
        meta_row.pack(fill=tk.X, pady=(self.px(14), 0))
        self.meta_lbl = self._label(meta_row, "", size=10, fg=C_SUBTLE)
        self.meta_lbl.pack(side=tk.LEFT)
        self.badge_lbl = tk.Label(meta_row, text="尚未載入", font=(FONT_UI, 9, "bold"),
                                  padx=self.px(8), pady=self.px(1))
        self.badge_lbl.pack(side=tk.LEFT, padx=(self.px(8), 0))
        self._set_badge("ok", "尚未載入")
        self._update_meta()
        self.font_style_var.trace_add("write", lambda *a: self._update_meta())
        self.font_size_var.trace_add("write", lambda *a: self._update_meta())

        self.copy_btn = FlatButton(right, "複製此字", command=self.copy_to_word_clipboard)
        self.copy_btn.pack(fill=tk.X, pady=(self.px(20), 0))
        self.copy_hint_lbl = self._label(right, "", size=9, weight="bold", fg=C_MUTED, anchor="w", justify="left")
        self.copy_hint_lbl.pack(fill=tk.X, pady=(self.px(8), 0))
        self.copy_hint_lbl.bind("<Configure>", lambda e: self.copy_hint_lbl.config(wraplength=max(150, e.width)))

        self.draw_grid_lines()
        self._draw_preview_placeholder("—")

        # ------------------------------------------------------------------
        # 區塊 3：手繪造字（全字庫查不到時使用）
        # ------------------------------------------------------------------
        sec_hand = tk.Frame(main, bg=C_CARD)
        sec_hand.pack(fill=tk.X, pady=(self.px(20), 0))
        self._label(sec_hand, "手繪造字", size=10, weight="bold").pack(anchor="w")
        self._label(sec_hand, "全字庫查不到時，直接在格子裡寫出這個字", size=9, fg=C_MUTED).pack(anchor="w")

        hand_border, hand = self._bordered(sec_hand)
        hand_border.pack(fill=tk.X, pady=(self.px(10), 0))
        hand_body = tk.Frame(hand, bg=C_CARD, padx=self.px(20), pady=self.px(20))
        hand_body.pack(fill=tk.X)

        self.pad_size = self.px(240)
        pad_border, pad_inner = self._bordered(hand_body)
        pad_border.pack(side=tk.LEFT, anchor="n")
        self.pad_canvas = tk.Canvas(pad_inner, width=self.pad_size, height=self.pad_size,
                                    bg=C_CARD, bd=0, highlightthickness=0, cursor="crosshair")
        self.pad_canvas.pack()
        self.pad_strokes = []            # 每筆：{"w": 筆寬(480 座標), "pts": [(x, y), ...]}
        self.pad_pen = PEN_DEFAULT
        self._pad_current = None
        self.pad_canvas.bind("<ButtonPress-1>", self._pad_down)
        self.pad_canvas.bind("<B1-Motion>", self._pad_move)
        self.pad_canvas.bind("<ButtonRelease-1>", self._pad_up)

        hand_right = tk.Frame(hand_body, bg=C_CARD)
        hand_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(self.px(24), 0))
        hand_inner = tk.Frame(hand_right, bg=C_CARD)
        hand_inner.pack(fill=tk.X, expand=True)

        self._label(hand_inner, "筆畫粗細", size=9, weight="bold", fg=C_MUTED).pack(anchor="w")
        pen_row = tk.Frame(hand_inner, bg=C_CARD)
        pen_row.pack(anchor="w", pady=(self.px(6), 0))
        self.pen_buttons = {}
        for label, width in PEN_OPTIONS:
            btn = FlatButton(pen_row, label, variant="primary" if width == self.pad_pen else "outline",
                             small=True, command=lambda w=width: self._pad_set_pen(w))
            btn.pack(side=tk.LEFT, padx=(0, self.px(8)))
            self.pen_buttons[width] = btn

        edit_row = tk.Frame(hand_inner, bg=C_CARD)
        edit_row.pack(anchor="w", pady=(self.px(20), 0))
        self.pad_undo_btn = FlatButton(edit_row, "↶  復原", variant="outline", small=True, command=self._pad_undo)
        self.pad_undo_btn.pack(side=tk.LEFT, padx=(0, self.px(8)))
        self.pad_clear_btn = FlatButton(edit_row, "清除", variant="outline", small=True, command=self._pad_clear)
        self.pad_clear_btn.pack(side=tk.LEFT)

        # 手寫辨識：按鈕 + 6 格候選字（停筆 0.5 秒也會自動辨識）
        self.recognizer = HandwritingRecognizer()
        self.cand_chars = []             # 目前候選字
        self.cand_codes = {}             # 字 → CNS 代碼（速查表或官網查到的結果）
        self.cand_picked = None
        self._recog_job = None
        self._cand_lookup_busy = False

        self.recog_btn = FlatButton(hand_inner, "辨識手寫", command=self._recognize_now)
        self.recog_btn.pack(fill=tk.X, pady=(self.px(20), 0))
        cand_head = tk.Frame(hand_inner, bg=C_CARD)
        cand_head.pack(fill=tk.X, pady=(self.px(10), self.px(6)))
        self._label(cand_head, "候選字", size=9, weight="bold", fg=C_MUTED).pack(side=tk.LEFT)
        self.cand_hint_lbl = self._label(cand_head, "", size=9, weight="bold", fg="#94A3B8")
        self.cand_hint_lbl.pack(side=tk.RIGHT)

        cand_border, cand_box = self._bordered(hand_inner, bg=C_SOFT)
        cand_border.pack(fill=tk.X)
        cand_grid = tk.Frame(cand_box, bg=C_SOFT, padx=self.px(6), pady=self.px(6))
        cand_grid.pack(fill=tk.X)
        self.cand_slots = []
        for i in range(HandwritingRecognizer.MAX_CANDIDATES):
            cand_grid.columnconfigure(i, weight=1, uniform="cand")
            slot_border = tk.Frame(cand_grid, bg=C_BORDER, padx=1, pady=1)
            slot_border.grid(row=0, column=i, sticky="nsew", padx=self.px(3))
            slot = tk.Frame(slot_border, bg=C_SOFT, height=self.px(80))
            slot.pack(fill=tk.BOTH, expand=True)
            slot.pack_propagate(False)
            num = tk.Label(slot, text="", bg=C_SOFT, fg="#94A3B8", font=(FONT_UI, 7))
            num.place(x=self.px(3), y=self.px(1))
            ch = tk.Label(slot, text="", bg=C_SOFT, fg=C_FG, font=(FONT_KAI, 18))
            ch.pack(pady=(self.px(6), 0))
            zy = tk.Label(slot, text="", bg=C_SOFT, fg=C_SUBTLE, font=(FONT_UI, 7, "bold"))
            zy.pack()
            code = tk.Label(slot, text="", bg=C_SOFT, fg="#94A3B8", font=(FONT_MONO, 7, "bold"))
            code.pack()
            parts = (slot, num, ch, zy, code)
            for w in parts:
                w.bind("<Button-1>", lambda e, idx=i: self._cand_pick(idx))
                w.bind("<Enter>", lambda e, idx=i: self._cand_hover(idx, True))
                w.bind("<Leave>", lambda e, idx=i: self._cand_hover(idx, False))
            self.cand_slots.append({"border": slot_border, "parts": parts, "num": num, "ch": ch, "zy": zy,
                                    "code": code})
        self.recog_error_lbl = self._label(hand_inner, "", size=9, weight="bold", fg=C_RED,
                                           anchor="w", justify="left")
        self.recog_error_lbl.bind("<Configure>", lambda e: self.recog_error_lbl.config(wraplength=max(150, e.width)))

        self.pad_use_btn = FlatButton(hand_inner, "↑  使用手繪字", variant="outline", command=self._pad_use)
        self.pad_use_btn.pack(fill=tk.X, pady=(self.px(20), 0))
        self._cand_render()
        self._label(hand_inner, "手繪字會載入上方預覽，之後可照常複製。", size=9, weight="bold",
                    fg=C_MUTED, anchor="w").pack(fill=tk.X, pady=(self.px(8), 0))
        self._pad_redraw()

        # ------------------------------------------------------------------
        # 區塊 4：使用紀錄
        # ------------------------------------------------------------------
        sec_hist = tk.Frame(main, bg=C_CARD)
        sec_hist.pack(fill=tk.X, pady=(self.px(20), 0))

        hist_head = tk.Frame(sec_hist, bg=C_CARD)
        hist_head.pack(fill=tk.X)
        head_text = tk.Frame(hist_head, bg=C_CARD)
        head_text.pack(side=tk.LEFT)
        self._label(head_text, "使用紀錄", size=10, weight="bold").pack(anchor="w")
        self._label(head_text, "點一下載入，點兩下直接複製", size=9, fg=C_MUTED).pack(anchor="w")
        self.clear_btn = FlatButton(hist_head, "全部清除", variant="ghost", small=True,
                                    command=self.clear_all_history)
        self.clear_btn.pack(side=tk.RIGHT)
        FlatButton(hist_head, "快取資料夾", variant="ghost", small=True,
                   command=self.open_cache_folder).pack(side=tk.RIGHT, padx=(0, self.px(4)))
        self.update_btn = FlatButton(hist_head, "檢查字庫更新", variant="ghost", small=True,
                                     command=self.check_cnsdata_update)
        self.update_btn.pack(side=tk.RIGHT, padx=(0, self.px(4)))

        table_border, table = self._bordered(sec_hist)
        table_border.pack(fill=tk.X, pady=(self.px(10), 0))

        columns = ("code", "style", "size", "time", "cache", "delete")
        self.history_tree = ttk.Treeview(
            table, columns=columns, show="tree headings", selectmode="browse",
            style="History.Treeview", height=5
        )
        self.history_tree.heading("#0", text="字形")
        self.history_tree.heading("code", text="代碼", anchor="w")
        self.history_tree.heading("style", text="字體", anchor="w")
        self.history_tree.heading("size", text="字級", anchor="w")
        self.history_tree.heading("time", text="最後使用", anchor="w")
        self.history_tree.heading("cache", text="快取", anchor="w")
        self.history_tree.heading("delete", text="")
        self.history_tree.column("#0", width=self.px(64), minwidth=self.px(56), stretch=False, anchor="center")
        self.history_tree.column("code", width=self.px(100), anchor="w")
        self.history_tree.column("style", width=self.px(80), anchor="w")
        self.history_tree.column("size", width=self.px(64), anchor="w")
        self.history_tree.column("time", width=self.px(140), anchor="w")
        self.history_tree.column("cache", width=self.px(80), anchor="w")
        self.history_tree.column("delete", width=self.px(40), minwidth=self.px(40), stretch=False, anchor="center")
        self.history_tree.tag_configure("row", foreground=C_FG)

        scrollbar = ttk.Scrollbar(table, orient=tk.VERTICAL, command=self.history_tree.yview,
                                  style="Slim.Vertical.TScrollbar")
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.history_empty_lbl = tk.Label(
            table, text="尚無紀錄。複製過的字會出現在這裡。",
            bg=C_CARD, fg=C_MUTED, font=(FONT_UI, 10)
        )

        self.history_tree.bind("<Button-1>", self._on_history_click, add=True)
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)
        self.history_tree.bind("<Double-1>", self.on_history_double_click)
        self.history_tree.bind("<Delete>", lambda e: self.delete_selected_history())
        self.history_tree.bind("<Motion>", self._on_history_motion)

    # --------------------------------------------------------------------------
    # 捲動與輸入輔助
    # --------------------------------------------------------------------------
    def _on_mousewheel(self, event):
        """滑鼠滾輪：在使用紀錄表內捲動表格，其他位置捲動主畫面；對話框不受影響"""
        widget = event.widget
        if not isinstance(widget, tk.Misc) or widget.winfo_toplevel() is not self.root:
            return None
        if widget is self.history_tree:
            return None                           # Treeview 自行處理捲動
        if widget is self.offset_spin:
            return None
        if self.main_vsb.winfo_ismapped():
            self.main_canvas.yview_scroll(int(-event.delta / 120), "units")
        return None

    def _normalize_offset(self):
        """把垂直對齊輸入整理成 -50 ~ 50 的整數字串"""
        try:
            v = int(float(self.baseline_offset_var.get()))
        except ValueError:
            v = 0
        v = max(OFFSET_MIN, min(OFFSET_MAX, v))
        if self.baseline_offset_var.get() != str(v):
            self.baseline_offset_var.set(str(v))
        self._on_size_or_offset_changed()

    # --------------------------------------------------------------------------
    # 手繪造字
    # --------------------------------------------------------------------------
    def _pad_xy(self, event):
        """畫布像素 → 480 座標系"""
        k = PAD_UNITS / self.pad_size
        x = max(0, min(self.pad_size, event.x)) * k
        y = max(0, min(self.pad_size, event.y)) * k
        return (x, y)

    def _pad_down(self, event):
        self._pad_current = {"w": self.pad_pen, "pts": [self._pad_xy(event)]}
        self.pad_strokes.append(self._pad_current)
        self._pad_redraw()

    def _pad_move(self, event):
        if self._pad_current is None:
            return
        pt = self._pad_xy(event)
        last = self._pad_current["pts"][-1]
        if abs(pt[0] - last[0]) + abs(pt[1] - last[1]) >= 1.0:
            self._pad_current["pts"].append(pt)
            self._pad_draw_stroke(self._pad_current, segment_only=True)

    def _pad_up(self, event):
        self._pad_current = None
        self._schedule_recognize()

    def _pad_draw_stroke(self, stroke, segment_only=False):
        """在畫布上畫一筆（segment_only：只補畫最後一段，拖曳時較流暢）"""
        k = self.pad_size / PAD_UNITS
        w = max(1.0, stroke["w"] * k)
        pts = stroke["pts"]
        if len(pts) == 1:
            x, y = pts[0][0] * k, pts[0][1] * k
            self.pad_canvas.create_oval(x - w / 2, y - w / 2, x + w / 2, y + w / 2,
                                        fill="#000000", outline="", tags="ink")
            return
        seq = pts[-2:] if segment_only else pts
        coords = [c * k for p in seq for c in p]
        self.pad_canvas.create_line(*coords, fill="#000000", width=w, capstyle=tk.ROUND,
                                    joinstyle=tk.ROUND, tags="ink")

    def _pad_redraw(self):
        c, s = self.pad_canvas, self.pad_size
        c.delete("all")
        c.create_line(0, s / 2, s, s / 2, fill="#FCA5A5", dash=(4, 3))
        c.create_line(s / 2, 0, s / 2, s, fill="#FCA5A5", dash=(4, 3))
        c.create_line(0, 0, s, s, fill="#FEE2E2", dash=(4, 3))
        c.create_line(s, 0, 0, s, fill="#FEE2E2", dash=(4, 3))
        for stroke in self.pad_strokes:
            self._pad_draw_stroke(stroke)
        has = bool(self.pad_strokes)
        for btn in (self.pad_undo_btn, self.pad_clear_btn, self.pad_use_btn):
            btn.set_enabled(has)
        if hasattr(self, "recog_btn"):
            self.recog_btn.set_enabled(has and not self._cand_lookup_busy)

    def _pad_set_pen(self, width):
        self.pad_pen = width
        for w, btn in self.pen_buttons.items():
            btn.set_variant("primary" if w == width else "outline")

    def _pad_undo(self):
        if self.pad_strokes:
            self.pad_strokes.pop()
            self._pad_redraw()
            if self.pad_strokes:
                self._schedule_recognize()
            else:
                self._cand_reset()

    def _pad_clear(self):
        self.pad_strokes = []
        self._pad_redraw()
        self._cand_reset()

    # --------------------------------------------------------------------------
    # 手寫辨識 → 候選字 → 全字庫代碼
    # --------------------------------------------------------------------------
    def _schedule_recognize(self):
        """停筆 0.5 秒後自動辨識（期間再下筆就重新計時）"""
        if self._recog_job:
            self.root.after_cancel(self._recog_job)
        self._recog_job = self.root.after(500, self._recognize_now)

    def _recognize_now(self):
        self._recog_job = None
        if not self.pad_strokes:
            self._cand_reset()
            return
        if not self.recognizer.available:
            self._show_recog_error(self.recognizer.error + "。請到「設定 > 時間與語言 > 語言與地區 > "
                                   "中文(台灣) > 語言選項」安裝「手寫」，或改用「使用手繪字」。")
            return
        try:
            cands = self.recognizer.recognize(self.pad_strokes)
        except Exception as e:
            self._show_recog_error(f"辨識失敗：{e}")
            return
        self.cand_chars = cands
        self.cand_picked = None
        for ch in cands:                                  # 離線字庫／速查表有的字，直接顯示代碼
            if ch in self.cand_codes:
                continue
            offline_codes = self.offline.codes_for(ch)
            known = next((c["cns"] for c in COMMON_RARE_CHARS if c["char"] == ch), None)
            if offline_codes:
                self.cand_codes[ch] = offline_codes
            elif known:
                self.cand_codes[ch] = [known]
        if cands:
            self._hide_recog_error()
        else:
            self._show_recog_error("辨識不出來，請寫清楚一點，或改用「使用手繪字」。")
        self._cand_render()

    def _cand_reset(self):
        if self._recog_job:
            self.root.after_cancel(self._recog_job)
            self._recog_job = None
        self.cand_chars, self.cand_picked = [], None
        self._hide_recog_error()
        self._cand_render()

    def _show_recog_error(self, msg: str):
        self.recog_error_lbl.config(text=msg)
        if not self.recog_error_lbl.winfo_ismapped():
            self.recog_error_lbl.pack(fill=tk.X, pady=(self.px(8), 0), before=self.pad_use_btn)

    def _hide_recog_error(self):
        if self.recog_error_lbl.winfo_ismapped():
            self.recog_error_lbl.pack_forget()

    def _cand_render(self):
        """重畫 6 格候選字：字、編號、注音、代碼（未查過顯示「點選查碼」）"""
        if self._cand_lookup_busy:
            hint = "查詢代碼中…"
        elif self.cand_chars:
            hint = "點選載入全字庫字形"
        elif self.pad_strokes:
            hint = "辨識中…"
        else:
            hint = "寫完後自動辨識"
        self.cand_hint_lbl.config(text=hint)
        for i, slot in enumerate(self.cand_slots):
            ch = self.cand_chars[i] if i < len(self.cand_chars) else ""
            filled = bool(ch)
            bg = C_CARD if filled else C_SOFT
            picked = filled and ch == self.cand_picked
            slot["border"].config(bg=C_FG if picked else C_BORDER)
            for w in slot["parts"]:
                w.config(bg=bg, cursor="hand2" if filled else "")
            slot["num"].config(text=str(i + 1) if filled else "")
            slot["ch"].config(text=ch)
            zys = self.zhuyin_for_char(ch) if filled else []
            slot["zy"].config(text=zys[0] if zys else "")
            codes = self.cand_codes.get(ch) if filled else None
            if not filled:
                slot["code"].config(text="")
            elif codes:
                slot["code"].config(text=codes[0] if len(codes) == 1 else f"{codes[0]}…", fg=C_FG)
            elif codes == []:
                slot["code"].config(text="查無", fg=C_RED)
            else:
                slot["code"].config(text="點選查碼", fg="#94A3B8")

    def _cand_hover(self, idx: int, entering: bool):
        if idx >= len(self.cand_chars):
            return
        bg = C_HOVER if entering else C_CARD
        for w in self.cand_slots[idx]["parts"]:
            w.config(bg=bg)

    def _cand_pick(self, idx: int):
        """點選候選字：查全字庫代碼（背景執行）→ 自動載入官方字圖"""
        if idx >= len(self.cand_chars) or self._cand_lookup_busy:
            return
        if self.is_fetching:
            self.set_status("字形下載中，請稍候再選。", is_loading=True)
            return
        ch = self.cand_chars[idx]
        self.cand_picked = ch
        if ch in self.cand_codes:
            self._cand_apply_codes(ch, self.cand_codes[ch])
            return
        self._cand_lookup_busy = True
        self.recog_btn.set_enabled(False)
        self._cand_render()
        self.set_status(f"正在查詢「{ch}」的全字庫代碼…", is_loading=True)

        def worker():
            try:
                codes = self.lookup_codes_by_char(ch)
                self.ui_queue.put(lambda: self._cand_lookup_done(ch, codes, None))
            except Exception as e:
                msg = str(e)
                self.ui_queue.put(lambda: self._cand_lookup_done(ch, None, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _cand_lookup_done(self, ch: str, codes, error):
        self._cand_lookup_busy = False
        self.recog_btn.set_enabled(bool(self.pad_strokes))
        if error is not None:
            self._cand_render()
            self.set_status(f"查詢「{ch}」的代碼失敗：{error}", is_error=True)
            return
        self.cand_codes[ch] = codes
        self._cand_render()
        self._cand_apply_codes(ch, codes)

    def _cand_apply_codes(self, ch: str, codes):
        self._cand_render()
        if not codes:
            self.set_status(f"全字庫查無「{ch}」，可改用「使用手繪字」。", is_error=True)
            return
        if len(codes) == 1:
            self._load_code_from_candidate(codes[0])
            return
        dlg = self._make_dialog(f"「{ch}」有多個全字庫代碼", "請選擇要載入的代碼。", width=380)
        row = tk.Frame(dlg.body, bg=C_CARD)
        row.pack(fill=tk.X, pady=(self.px(16), 0))
        for code in codes:
            FlatButton(row, code, variant="outline",
                       command=lambda c=code: (dlg.destroy(), self._load_code_from_candidate(c))
                       ).pack(side=tk.LEFT, padx=(0, self.px(8)))
        self._center_dialog(dlg)

    def _load_code_from_candidate(self, code: str):
        self.code_var.set(code)
        self.fetch_character()
        self.main_canvas.yview_moveto(0)

    def lookup_codes_by_char(self, ch: str) -> list:
        """
        以官網快速查詢（searchQ.jsp）取得某字的 CNS 代碼。
        結果頁中的 wordView.jsp?ID=n，n = 字面×65536 + 碼位。連線失敗拋出例外。
        """
        import urllib.parse
        offline_codes = self.offline.codes_for(ch)   # 離線對照表優先（即時、不需網路）
        if offline_codes:
            return offline_codes
        _ensure_requests()
        session = _get_shared_session()
        url = f"{self.CNS_BASE}/searchQ.jsp?WORD={urllib.parse.quote(ch)}"
        try:
            status, body, location = self._http_get(session, url, max_bytes=self.MAX_PAGE_BYTES,
                                                    want_location=True)
        except requests.exceptions.SSLError:
            _reset_shared_session()
            raise RuntimeError("全字庫官網的安全憑證驗證失敗（可能遭攔截）") from None
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
            _reset_shared_session()
            raise RuntimeError("無法連線全字庫官網") from None
        if status in (301, 302, 303, 307, 308):
            # 只對到一個字時，官網轉址到 wordView.jsp?ID=n：只讀出 ID，不跟隨轉址
            m = re.fullmatch(r"(?:https://www\.cns11643\.gov\.tw/)?wordView\.jsp\?ID=(\d+)", location or "")
            if not m:
                raise RuntimeError("官網回應異常（未預期的轉址）")
            n = int(m.group(1))
            code = f"{n // 65536}-{n % 65536:04X}"
            return [code] if self._is_std_code(code) else []
        if status != 200:
            raise RuntimeError(f"官網回應異常（HTTP {status}）")
        codes = []
        for m in re.finditer(r"wordView\.jsp\?ID=(\d+)", body.decode("utf-8", errors="replace")):
            n = int(m.group(1))
            code = f"{n // 65536}-{n % 65536:04X}"
            if self._is_std_code(code) and code not in codes:
                codes.append(code)
        # 多筆結果時，清單會混入偏旁或相關字形（例：查「金」會出現 钅、糹、釒），
        # 逐一核對字形資訊頁標題，只保留真的是這個字的代碼
        exact = []
        for code in codes[:12]:
            p, h = code.split("-")
            st, page = self._http_get(session, f"{self.CNS_BASE}/wordView.jsp?ID={int(p) * 65536 + int(h, 16)}",
                                      max_bytes=self.MAX_PAGE_BYTES)
            if st != 200:
                raise RuntimeError(f"全字庫官網暫時忙碌（HTTP {st}），請稍後再試")
            t = re.search(r"<title>\s*字形資訊\s*-\s*\[(.{1,2}?)\]", page.decode("utf-8", errors="replace"))
            if t and t.group(1) == ch:
                exact.append(code)
        return exact

    def render_pad_image(self) -> Image.Image:
        """把手繪筆畫以 512×512 高解析度重新繪製（向量重繪，不是放大螢幕截圖）"""
        size = 512
        k = size / PAD_UNITS
        img = Image.new("RGB", (size, size), (255, 255, 255))
        draw = ImageDraw.Draw(img)
        for stroke in self.pad_strokes:
            w = max(1, int(round(stroke["w"] * k)))
            pts = [(x * k, y * k) for x, y in stroke["pts"]]
            if len(pts) > 1:
                draw.line(pts, fill=(0, 0, 0), width=w, joint="curve")
            r = w / 2
            for x, y in pts:   # 圓頭：每個節點補圓，轉角平滑
                draw.ellipse([x - r, y - r, x + r, y + r], fill=(0, 0, 0))
        return img

    def _pad_use(self):
        """把手繪字載入上方預覽，之後可照常複製"""
        if not self.pad_strokes:
            return
        if self.is_fetching:
            self.set_status("字形下載中，請稍候再使用手繪字。", is_loading=True)
            return
        self.current_pil_image = self.normalize_glyph_image(self.render_pad_image().convert("RGBA"))
        self.current_cns_code = "手繪"
        self.current_source = "hand"
        self.current_is_fallback = False
        self.current_style_is_sung = "宋體" in self.font_style_var.get()
        self.update_preview_display()
        self.set_status("已載入手繪字，可直接複製")
        self.main_canvas.yview_moveto(0)

    # --------------------------------------------------------------------------
    # 狀態顯示
    # --------------------------------------------------------------------------
    def set_status(self, message: str, is_error: bool = False, is_loading: bool = False):
        """更新底部狀態列：一般 ✓、讀取中 ⋯、錯誤 ⚠（紅字）"""
        if is_error:
            prefix, fg = "⚠  ", C_RED_TEXT
        elif is_loading:
            prefix, fg = "⋯  ", C_SUBTLE
        else:
            prefix, fg = "✓  ", C_SUBTLE
        self.status_lbl.config(fg=fg)
        self.status_var.set(prefix + message)

    def _set_badge(self, tone: str, text: str):
        bg, fg = BADGE_TONES.get(tone, BADGE_TONES["ok"])
        self.badge_lbl.config(text=text, bg=bg, fg=fg)

    def _show_input_error(self, message: str):
        self.input_error_lbl.config(text=message)
        self.input_error_lbl.pack(fill=tk.X, pady=(self.px(6), 0), after=self.q_results_wrap)

    def _clear_input_error(self):
        if self.input_error_lbl.winfo_ismapped():
            self.input_error_lbl.pack_forget()

    def _on_code_changed(self):
        self._clear_input_error()
        self.update_copy_state()
        if self.query_mode == "cns":
            self._schedule_query()        # 依代碼前綴篩選下方字卡

    def _on_size_or_offset_changed(self):
        self.update_copy_state()

    # --------------------------------------------------------------------------
    # 注音
    # --------------------------------------------------------------------------
    def _offline_loaded(self) -> bool:
        """離線字庫已載入完成（不等待，避免卡住介面）"""
        return self.offline.available and self.offline._loaded.is_set()

    def char_for_code(self, code: str):
        item = next((c for c in COMMON_RARE_CHARS if c["cns"] == code), None)
        if item:
            return item["char"]
        return self.offline.char_for(code) if self._offline_loaded() else None

    def zhuyin_for_code(self, code: str) -> list:
        """速查表優先（人名讀音排第一），其次離線字庫的官網注音"""
        item = next((c for c in COMMON_RARE_CHARS if c["cns"] == code), None)
        if item:
            return list(item["zhuyin"])
        return self.offline.zhuyin_for(code) if self._offline_loaded() else []

    def zhuyin_for_char(self, ch: str) -> list:
        item = next((c for c in COMMON_RARE_CHARS if c["char"] == ch), None)
        if item:
            return list(item["zhuyin"])
        if not self._offline_loaded():
            return []
        for code in self.offline.codes_for(ch):
            zys = self.offline.zhuyin_for(code)
            if zys:
                return zys
        return []

    def _update_zhuyin_line(self):
        """預覽區代碼下方的「字・注音」，例如「煊・ㄒㄩㄢ／ㄋㄨㄢˇ」"""
        code = self.current_cns_code
        ch = self.char_for_code(code) if code and self.current_source != "hand" else None
        if not ch:
            if self.zhuyin_lbl.winfo_ismapped():
                self.zhuyin_lbl.pack_forget()
            return
        zys = self.zhuyin_for_code(code)
        self.zhuyin_lbl.config(text=f"{ch}・{'／'.join(zys[:3]) if zys else '—'}")
        if not self.zhuyin_lbl.winfo_ismapped():
            self.zhuyin_lbl.pack(anchor="w", pady=(self.px(2), 0), after=self.cns_code_lbl)

    # --------------------------------------------------------------------------
    # 查詢方式（CNS 代碼／注音／拼音／部首筆畫／部件／倉頡／筆順／Unicode／複合）
    # --------------------------------------------------------------------------
    def _entry(self, parent, var, font, placeholder="", on_return=None, padx=10, pady=5):
        """細框輸入框＋灰色提示字（tk.Entry 沒有 placeholder）"""
        border, inner = self._bordered(parent, color=C_INPUT)
        entry = tk.Entry(inner, textvariable=var, font=font, width=6, bd=0, relief="flat",
                         bg=C_CARD, fg=C_FG, insertbackground=C_FG)
        entry.pack(fill=tk.BOTH, expand=True, padx=self.px(padx), pady=self.px(pady))
        entry.bind("<FocusIn>", lambda e: border.config(bg=C_FG))
        entry.bind("<FocusOut>", lambda e: border.config(bg=C_INPUT))
        if on_return:
            entry.bind("<Return>", lambda e: on_return())
        ph = tk.Label(inner, text=placeholder, bg=C_CARD, fg="#94A3B8", font=font, cursor="xterm")
        ph.bind("<Button-1>", lambda e: entry.focus_set())

        def toggle(*_):
            if var.get() or not ph.cget("text"):
                ph.place_forget()
            else:
                ph.place(x=self.px(padx), rely=0.5, anchor="w")

        var.trace_add("write", toggle)
        toggle()
        entry.placeholder, entry.toggle_placeholder = ph, toggle
        return border, entry

    @staticmethod
    def _insert_text(entry, text: str):
        try:
            entry.insert(tk.INSERT, text)
            entry.icursor(tk.INSERT)
        except tk.TclError:
            pass

    def _radical_pair(self, parent_r, parent_s, rad_var, stk_var):
        """部首（可選可打字）＋總筆畫（只列該部首實際有的筆畫數）"""
        rc = ttk.Combobox(parent_r, textvariable=rad_var, font=(FONT_UI, 10), height=16, width=6)
        sc = ttk.Combobox(parent_s, textvariable=stk_var, state="readonly", font=(FONT_UI, 10), height=16, width=6)

        def changed(typing=False):
            self._on_radical_changed(rad_var, stk_var, sc, typing=typing)

        rc.bind("<<ComboboxSelected>>", lambda e: (rc.selection_clear(), changed()))
        rc.bind("<KeyRelease>", lambda e: changed(True))
        rc.bind("<Return>", lambda e: (self.refresh_query_results(), self.query_load()))
        rc.bind("<FocusIn>", lambda e: rc.selection_range(0, tk.END))
        sc.bind("<<ComboboxSelected>>", lambda e: (sc.selection_clear(), self.refresh_query_results()))
        self._radical_widgets.append((rc, rad_var, stk_var, sc))
        return rc, sc

    def _pack_query_buttons(self, compact: bool):
        """一般分頁：按鈕與輸入框等高；複合分頁：按鈕對齊表單底部"""
        for b in (self.fetch_btn, self.official_btn):
            b.pack_forget()
        for b in (self.fetch_btn, self.official_btn):
            if compact:
                b.pack(side=tk.LEFT, anchor="s", padx=(self.px(8), 0))
            else:
                b.pack(side=tk.LEFT, fill=tk.Y, padx=(self.px(8), 0))

    def _set_part_exact(self, exact: bool):
        self.q_part_exact = exact
        for v, btn in self.part_mode_btns.items():
            btn.set_variant("primary" if v == exact else "outline")
        self.refresh_query_results()

    def _mode_info(self) -> dict:
        return next(m for m in QUERY_MODES if m["v"] == self.query_mode)

    def set_query_mode(self, mode: str, focus: bool = True):
        self.query_mode = mode
        info = self._mode_info()
        for v, tab in self.mode_tabs.items():
            on = v == mode
            tab.config(bg=C_CARD if on else C_HOVER, fg=C_FG if on else C_SUBTLE)
        self.mode_hint_lbl.config(text=info["hint"])
        for w in (self.code_entry_border, self.q_entry_border, self.q_radical_box, self.q_compound_box,
                  *self.q_extra.values()):
            w.pack_forget()
        self._clear_input_error()
        target = None
        if mode == "cns":
            self.code_entry_border.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            target = self.code_entry
        elif mode == "radical":
            self.q_radical_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            target = self.q_radical_combo
        elif mode == "compound":
            self.q_compound_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            self.q_text_var.set("")
            font = self.q_entry_fonts.get(mode, (FONT_UI, 12))
            self.q_entry.config(font=font)
            self.q_entry.placeholder.config(text=info.get("ph", ""), font=font)
            self.q_entry.toggle_placeholder()
            self.q_entry_border.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            if mode in self.q_extra:
                self.q_extra[mode].pack(side=tk.LEFT, fill=tk.Y, padx=(self.px(8), 0))
            target = self.q_entry
        self._pack_query_buttons(compact=mode == "compound")
        self.refresh_query_results()
        if focus and target is not None:
            target.focus_set()

    def _schedule_query(self):
        """打字時稍等一下再查（部分查詢要掃整個離線字庫）"""
        if self._q_job:
            self.root.after_cancel(self._q_job)
        self._q_job = self.root.after(150, self.refresh_query_results)

    def open_official_query(self):
        webbrowser.open(self._mode_info()["url"])

    def _wait_extras(self):
        """倉頡、筆順、部件、拼音資料在背景載入，完成後更新目前的查詢結果"""
        if not self.offline.extras_loaded.is_set():
            self.root.after(300, self._wait_extras)
            return
        self._pinyin_cache = None
        if self.query_mode in ("pinyin", "part", "cangjie", "strokeseq", "compound"):
            self.refresh_query_results()

    def _require(self, data, what: str):
        """查詢需要的資料不存在時，丟出給使用者看的說明"""
        if data:
            return
        if not self.offline.available:
            raise QueryInputError(f"需要離線字庫（cnsdata）才能用{what}查詢，可按「到官網查詢」。")
        if not self.offline._loaded.is_set() or not self.offline.extras_loaded.is_set():
            raise QueryInputError("字庫資料載入中，請稍候…")
        raise QueryInputError(f"離線字庫缺少{what}資料，請以 tools/build_cnsdata.py 重新產生 cnsdata。")

    # ---- 注音 ----
    def _zy_index(self) -> list:
        """離線注音索引：[(代碼, 注音, 去聲調注音)]，第一次查詢時才建立"""
        idx = getattr(self, "_zy_index_cache", None)
        if idx is None and self._offline_loaded():
            idx = [(code, zy, re.sub("[ˊˇˋ˙]", "", zy))
                   for code, zys in self.offline.code2zy.items() for zy in zys]
            self._zy_index_cache = idx
        return idx or []

    def _parse_zhuyin_query(self, qn: str):
        """回傳 (要比對的讀音清單, 說明, 錯誤訊息)。可輸入注音符號（ˉ＝只要一聲），或一個同音的字。"""
        if re.fullmatch(r"[ㄅ-ㄩˊˇˋ˙ˉ]+", qn):
            if "ˉ" in qn:
                return [qn.replace("ˉ", "") + "ˉ"], None, None
            return [qn], None, None
        if len(qn) == 1:
            readings = self.zhuyin_for_char(qn)
            if not readings:
                return [], None, f"查不到「{qn}」的讀音，請改輸入注音符號，例如 ㄒㄩㄢ。"
            wanted = [z if re.search("[ˊˇˋ˙]", z) else z + "ˉ" for z in readings]    # 同音＝讀音與聲調都相同
            return wanted, f"以「{qn}」的讀音 {'／'.join(readings)} 查同音字。", None
        return [], None, "請輸入注音符號（例如 ㄒㄩㄢ），或輸入一個同音的字。"

    @staticmethod
    def _zy_matcher(wanted: list):
        """
        讀音比對規則：
          ㄒㄩㄢˉ → 只要一聲；ㄒㄩㄢ~ → 這個音的各聲調（拼音不加聲調時）；
          有聲調符號 → 讀音完全相同；沒有聲調符號 → 前綴比對（各聲調都算，可只打一部分）
        """
        rules = []
        for w in wanted:
            if w.endswith("ˉ"):
                rules.append(("eq", w[:-1]))
            elif w.endswith("~"):
                rules.append(("base", w[:-1]))
            elif re.search("[ˊˇˋ˙]", w):
                rules.append(("eq", w))
            else:
                rules.append(("prefix", w))

        def match(zy: str, base: str) -> bool:
            for kind, w in rules:
                if (kind == "eq" and zy == w) or (kind == "base" and base == w) \
                        or (kind == "prefix" and base.startswith(w)):
                    return True
            return False

        return match

    def _zy_codes(self, wanted: list) -> dict:
        """符合讀音的代碼 → 符合的讀音（內建速查表的人名讀音優先）"""
        match = self._zy_matcher(wanted)
        zymap = {}
        for code, zy, base in self._zy_index():
            if code not in zymap and match(zy, base):
                zymap[code] = zy
        for c in COMMON_RARE_CHARS:
            z = next((z for z in c["zhuyin"] if match(z, re.sub("[ˊˇˋ˙]", "", z))), None)
            if z:
                zymap[c["cns"]] = z
        return zymap

    def _c_zhuyin(self, text: str):
        wanted, note, err = self._parse_zhuyin_query(re.sub(r"\s", "", text))
        if err:
            raise QueryInputError(err)
        return self._zy_codes(wanted), note

    # ---- 拼音 ----
    @staticmethod
    def _norm_pinyin(s: str) -> str:
        s = s.strip().lower()
        for a, b in (("’", "'"), ("‘", "'"), ("`", "'"), ("u:", "ü"), ("v", "ü")):
            s = s.replace(a, b)
        return re.sub(r"\s+", "", s)

    def _pinyin_index(self) -> dict:
        """拼音 → [(拼音方式欄位, 要比對的讀音)]；第一次查詢時才建立"""
        idx = getattr(self, "_pinyin_cache", None)
        if idx:
            return idx
        idx = {}
        for r in self.offline.pinyin_rows:
            zy = r[0]
            base = re.sub("[ˊˇˋ˙]", "", zy)
            tone = 5 if "˙" in zy else next((t for m, t in (("ˊ", 2), ("ˇ", 3), ("ˋ", 4)) if m in zy), 1)
            exact = zy if tone != 1 else zy + "ˉ"
            for col in range(1, 6):
                num = self._norm_pinyin(r[col])
                b = re.sub(r"[1-5]$", "", num)
                if not b:
                    continue
                idx.setdefault(b + str(tone), []).append((col, exact))
                idx.setdefault(b, []).append((col, base + "~"))
                mark = self._norm_pinyin(r[col + 5])
                if mark and not re.fullmatch(r"[a-z']+", mark):      # 帶調號的寫法（純字母留給「不分聲調」）
                    idx.setdefault(mark, []).append((col, exact))
        self._pinyin_cache = idx
        return idx

    def _c_pinyin(self, text: str, sys_col):
        self._require(self.offline.pinyin_rows, "拼音")
        q = self._norm_pinyin(text)
        hits = [w for col, w in self._pinyin_index().get(q, []) if sys_col is None or col == sys_col]
        if not hits:
            raise QueryInputError(f"看不懂拼音「{text.strip()}」，請確認拼法，例如 xuan、xuan2、xuān、hsüan2。")
        wanted = list(dict.fromkeys(hits))
        shown = "／".join(dict.fromkeys(
            w[:-1] + "（各聲調）" if w.endswith("~") else w.rstrip("ˉ") for w in wanted))
        return self._zy_codes(wanted), f"「{text.strip()}」＝{shown}"

    # ---- 部首/筆畫 ----
    def _radical_options(self) -> list:
        if self.offline.has_radicals():
            return [Q_ALL_RADICALS] + [f"{n:03d} {name}" for n, name in sorted(self.offline.rad_names.items())]
        return [Q_ALL_RADICALS] + [f"{r} 部" for r in sorted({c["radical"] for c in COMMON_RARE_CHARS})]

    def _parse_radical(self, text: str):
        """None＝不限部首；整數＝部首序號（整庫）；單字＝內建速查表的部首；"invalid"＝看不懂"""
        t = text.strip()
        if not t or t == Q_ALL_RADICALS:
            return None
        full = self.offline.has_radicals()
        m = re.match(r"(\d{1,3})", t)
        if m:
            n = int(m.group(1))
            return n if full and 1 <= n <= 214 else "invalid"
        ch = t[0]
        if full:
            names = self.offline.rad_names
            n = next((n for n, name in names.items() if name.startswith(ch)), None) \
                or next((n for n, name in names.items() if ch in name), None)
            return n if n else "invalid"
        return ch if any(c["radical"] == ch for c in COMMON_RARE_CHARS) else "invalid"

    @staticmethod
    def _parse_strokes(text: str):
        m = re.match(r"(\d+)", text.strip())
        return int(m.group(1)) if m else None

    def _stroke_options(self, rad) -> list:
        if self.offline.has_radicals():
            if isinstance(rad, int):
                strokes = {self.offline.code2rs[c][1] for c in self.offline.rad2codes.get(rad, [])}
            else:
                strokes = set(self.offline.stk2codes)
        else:
            strokes = {int(c["strokes"]) for c in COMMON_RARE_CHARS if not isinstance(rad, str) or c["radical"] == rad}
        return [Q_ALL_STROKES] + [f"{n} 畫" for n in sorted(strokes)]

    def _fill_radical_options(self):
        opts = self._radical_options()
        for rc, rad_var, stk_var, sc in self._radical_widgets:
            rc.config(values=opts)
            self._on_radical_changed(rad_var, stk_var, sc, refresh=False)

    def _on_radical_changed(self, rad_var, stk_var, sc, typing: bool = False, refresh: bool = True):
        """部首改變時，筆畫清單只留該部首實際有的筆畫數"""
        rad = self._parse_radical(rad_var.get())
        values = self._stroke_options(rad if rad != "invalid" else None)
        sc.config(values=values)
        if stk_var.get() not in values:
            stk_var.set(Q_ALL_STROKES)
        if refresh:
            self._schedule_query() if typing else self.refresh_query_results()

    def _radical_sub(self, code: str) -> str:
        r, s = self.offline.code2rs.get(code, (0, 0))
        return f"{self.offline.rad_names.get(r, '')[:1]}・{s}畫" if r else code

    # ---- 倉頡、筆順、部件、CNS 範圍 ----
    def _c_cangjie(self, text: str) -> set:
        lib = self.offline
        self._require(lib.code2cj, "倉頡")
        q = "".join(CJ_ROOTS.get(ch, ch) for ch in re.sub(r"\s", "", text)).upper()
        q = q.replace("＊", "*").replace("？", "?")
        if not re.fullmatch(r"[A-Z*?]{1,10}", q):
            raise QueryInputError("倉頡碼只能是英文字母 A–Z 或字根（日月金木水火土…），可加 ＊ ？。")
        if not re.search("[A-Z]", q):
            raise QueryInputError("請至少輸入一個倉頡碼字母或字根。")
        if "*" in q or "?" in q:
            rx = re.compile(q.replace("?", ".").replace("*", ".*"))
            return {c for c, v in lib.code2cj.items() if rx.fullmatch(v)}
        return set(lib.cj2codes.get(q, []))

    def _c_strokeseq(self, text: str) -> set:
        lib = self.offline
        self._require(lib.code2seq, "筆順")
        q = "".join(STROKE_DIGITS.get(ch, "5" if 0x31C0 <= ord(ch) <= 0x31E3 else ch)    # 其餘筆形都屬「折」
                    for ch in re.sub(r"\s", "", text))
        q = q.replace("＊", "*").replace("？", "?")
        if not re.fullmatch(r"[1-5*?]{1,99}", q):
            raise QueryInputError("筆順只能用 橫 豎 撇 點 折（或數字 1–5），可加 ＊ ？。")
        if not re.search("[1-5]", q):
            raise QueryInputError("請至少輸入一筆。")
        if "*" in q or "?" in q:
            rx = re.compile(q.replace("?", ".").replace("*", ".*"))
            return {c for c, v in lib.code2seq.items() if rx.fullmatch(v)}
        return set(lib.seq2codes.get(q, []))

    def _part_tokens(self, text: str) -> list:
        """把輸入的字／部件拆成部件集合序列；＊ ？ 保留為萬用字元"""
        lib = self.offline
        tokens = []
        for ch in text:
            if ch in " +＋、,，":
                continue
            if ch in "*＊":
                tokens.append("*")
                continue
            if ch in "?？":
                tokens.append("?")
                continue
            cp, groups = ord(ch), None
            for code in lib.cp2codes.get(cp, []):
                ids = lib.code2comp.get(code)
                if ids:
                    groups = [lib.comp_group(i) for i in ids]
                    break
            if groups is None and cp in lib.cp2comp:
                groups = [lib.comp_group(lib.cp2comp[cp])]
            if groups is None:
                raise QueryInputError(f"「{ch}」查不到部件資料，請換一個字或改用部件表。")
            tokens.extend(groups)
        return tokens

    def _c_part(self, text: str, exact: bool) -> set:
        """模糊：包含這些部件（同一部件的不同寫法都算）；精確：剛好由這些部件依序組成（可用 ＊ ？）"""
        from collections import Counter
        lib = self.offline
        self._require(lib.code2comp, "部件")
        tokens = self._part_tokens(text)
        need = Counter(t for t in tokens if isinstance(t, int))
        if not need:
            raise QueryInputError("請至少輸入一個字或部件。")
        cands = None
        for g in sorted(need, key=lambda g: len(lib.group2codes.get(g, ()))):
            s = lib.group2codes.get(g, set())
            cands = set(s) if cands is None else cands & s
            if not cands:
                return set()
        if exact:
            rx = re.compile("".join(f"{t}," if isinstance(t, int) else (r"-?\d+," if t == "?" else r"(?:-?\d+,)*")
                                    for t in tokens))
            return {c for c in cands if rx.fullmatch("".join(f"{lib.comp_group(i)}," for i in lib.code2comp[c]))}
        out = set()
        for c in cands:
            have = Counter(lib.comp_group(i) for i in lib.code2comp[c])
            if all(have[g] >= n for g, n in need.items()):
                out.add(c)
        return out

    def _cns_range(self, text: str):
        """CNS 範圍：3-4200~3-42FF、3-4200~42FF、只打前幾碼（3-42）或單一代碼 → (字面, 起, 迄)；看不懂回傳 None"""
        s = re.sub(r"^CNS(11643)?\s*[-_:\s]*", "", self.clean_input_string(text))
        parts = re.split(r"\s*[~～至到]\s*", s)
        if len(parts) == 2:
            a = self.parse_cns_code(parts[0])
            if not a:
                return None
            if re.fullmatch(r"[0-9A-F]{4}", parts[1]):
                hi = int(parts[1], 16)
            else:
                b = self.parse_cns_code(parts[1])
                if not b or b[0] != a[0]:
                    return None
                hi = int(b[1], 16)
            lo = int(a[1], 16)
            return a[0], min(lo, hi), max(lo, hi)
        m = re.fullmatch(r"([0-9A-F]{1,2})[-_\s:]+([0-9A-F]{0,3})", s)
        if m:
            plane = self._parse_plane_number(m.group(1))
            if not plane or not 1 <= plane <= 19:
                return None
            h = m.group(2)
            return plane, int(h.ljust(4, "0"), 16), int(h.ljust(4, "F"), 16)
        p = self.parse_cns_code(s)
        if p:
            v = int(p[1], 16)
            return p[0], v, v
        return None

    def _codes_in_range(self, rng) -> set:
        plane, lo, hi = rng
        prefix = f"{plane}-"
        if self._offline_loaded():
            return {c for c in self.offline.code2cp
                    if c.startswith(prefix) and lo <= int(c[len(prefix):], 16) <= hi}
        return {c["cns"] for c in COMMON_RARE_CHARS
                if c["cns"].startswith(prefix) and lo <= int(c["cns"][len(prefix):], 16) <= hi}

    # ---- 組合條件 ----
    def _filter_codes(self, f: dict):
        """
        f：rad、stk、zhuyin、pinyin、pinyin_sys、cangjie、strokeseq、part、part_exact、cns（都可省略）
        回傳 (代碼集合, 讀音對照, 說明清單)；沒有任何條件時回傳 None
        """
        lib = self.offline
        sets, zymap, notes = [], None, []
        rad = self._parse_radical(f.get("rad", ""))
        if rad == "invalid":
            raise QueryInputError("找不到這個部首，請從清單選擇，或輸入部首字（如 火）或序號（如 86）。")
        if rad is not None:
            self._require(lib.code2rs, "部首")
            sets.append(set(lib.rad2codes.get(rad, [])))
        stk = self._parse_strokes(f.get("stk", ""))
        if stk is not None:
            self._require(lib.code2rs, "筆畫")
            sets.append(set(lib.stk2codes.get(stk, [])))
        for key in ("zhuyin", "pinyin"):
            text = f.get(key, "").strip()
            if not text:
                continue
            zm, note = self._c_zhuyin(text) if key == "zhuyin" else self._c_pinyin(text, f.get("pinyin_sys"))
            sets.append(set(zm))
            zymap = zm if zymap is None else {c: z for c, z in zymap.items() if c in zm}
            if note:
                notes.append(note)
        if f.get("cangjie", "").strip():
            sets.append(self._c_cangjie(f["cangjie"]))
        if f.get("strokeseq", "").strip():
            sets.append(self._c_strokeseq(f["strokeseq"]))
        if f.get("part", "").strip():
            sets.append(self._c_part(f["part"], f.get("part_exact", False)))
        if f.get("cns", "").strip():
            rng = self._cns_range(f["cns"])
            if not rng:
                raise QueryInputError("CNS 範圍格式不對，例如 3-4200~3-42FF、3-42 或 3-424E。")
            sets.append(self._codes_in_range(rng))
        if not sets:
            return None
        sets.sort(key=len)
        codes = set(sets[0])
        for s in sets[1:]:
            codes &= s
        return codes, zymap, notes

    def _assemble(self, codes, zymap=None, order="strokes", sub=None, notes=None, more_hint="可再加條件縮小範圍。"):
        """代碼 → 字卡資料：排序、同一個字只列一次、超過 3 排只列前面"""
        lib = self.offline
        builtin = {c["cns"]: c for c in COMMON_RARE_CHARS}
        rows = []
        for code in codes:
            b = builtin.get(code)
            ch = b["char"] if b else (chr(lib.code2cp[code]) if code in lib.code2cp else None)
            if not ch:
                continue
            if zymap and code in zymap:
                zy = zymap[code]
            else:
                zys = (b["zhuyin"] if b else None) or lib.code2zy.get(code, [])
                zy = zys[0] if zys else ""
            plane, hexpart = code.split("-")
            stroke = lib.code2rs.get(code, (0, int(b["strokes"]) if b else 99))[1]
            tail = (0 if int(plane) <= 2 else 1, int(plane), int(hexpart, 16))
            if order == "zhuyin":
                key = (zhuyin_sort_key(zy),) + tail          # 與官網相同：同音字常用字面（1、2）在前
            elif order == "code":
                key = (int(plane), int(hexpart, 16))
            else:
                key = (stroke,) + tail
            rows.append((key, code, ch, zy))
        rows.sort(key=lambda r: r[0])
        results, seen = [], set()
        for _, code, ch, zy in rows:
            if ch in seen:                  # 同一個字有多個代碼時只列第一個
                continue
            seen.add(ch)
            results.append({"cns": code, "char": ch, "zhuyin": zy, "sub": sub(code) if sub else code})
        notes = [n for n in (notes or []) if n]
        limit = max(self._q_cols, 4) * Q_MAX_ROWS
        if len(results) > limit:
            notes.append(f"共 {len(results)} 個字，只列出前 {limit} 個；{more_hint}")
            results = results[:limit]
        return results, (" ".join(notes) or None)

    def compute_query_results(self):
        """回傳 (結果清單, 說明文字或 None)"""
        mode = self.query_mode
        q = self.q_text_var.get().strip()
        sub_of = {
            "unicode": lambda c: f"U+{ord(c['char']):04X}",
            "radical": lambda c: f"{c['radical']}・{c['strokes']}畫",
            "part": lambda c: c["desc"],
        }.get(mode, lambda c: c["cns"])

        def builtin(items):
            items = sorted(items, key=lambda c: zhuyin_sort_key(c["zhuyin"][0]))
            return [{"cns": c["cns"], "char": c["char"], "zhuyin": c["zhuyin"][0], "sub": sub_of(c)} for c in items]

        try:
            if mode == "cns":
                return self._cns_results(builtin)
            if mode == "unicode":
                return self._unicode_results(q, builtin)
            if mode in ("zhuyin", "pinyin", "cangjie", "strokeseq", "part") and not q:
                return builtin(COMMON_RARE_CHARS), None
            if mode == "zhuyin":
                zm, note = self._c_zhuyin(q)
                return self._assemble(zm, zm, order="zhuyin", notes=[note], more_hint="加上聲調可以縮小範圍。")
            if mode == "pinyin":
                sys_col = dict(PINYIN_SYSTEMS).get(self.q_pinyin_sys_var.get())
                zm, note = self._c_pinyin(q, sys_col)
                return self._assemble(zm, zm, order="zhuyin", notes=[note], more_hint="加上聲調可以縮小範圍。")
            if mode == "cangjie":
                return self._assemble(self._c_cangjie(q), sub=lambda c: self.offline.code2cj.get(c, c))
            if mode == "strokeseq":
                return self._assemble(self._c_strokeseq(q), more_hint="再多輸入幾筆可以縮小範圍。")
            if mode == "part":
                if not self.offline.code2comp and not self.offline.available:
                    return self._builtin_part(q, builtin), None
                return self._assemble(self._c_part(q, self.q_part_exact), sub=self._radical_sub)
            if mode == "radical":
                return self._radical_results(builtin)
            if mode == "compound":
                return self._compound_results(builtin)
        except QueryInputError as e:
            return [], str(e)
        return [], None

    def _cns_results(self, builtin):
        """CNS 代碼：空白或單一代碼 → 速查表（依前綴篩選）；範圍或只打前幾碼 → 整個字庫"""
        ci = self.code_var.get().strip()
        if ci and not self.parse_cns_code(ci):
            rng = self._cns_range(ci)
            if rng:
                return self._assemble(self._codes_in_range(rng), order="code",
                                      more_hint="範圍縮小一點就能看到更多。")
        cu = ci.upper()
        return builtin([c for c in COMMON_RARE_CHARS if cu and c["cns"].startswith(cu)] or COMMON_RARE_CHARS), None

    def _unicode_results(self, q: str, builtin):
        if not q:
            return builtin(COMMON_RARE_CHARS), None
        qu = re.sub(r"^U\+?", "", q.upper())
        cp = None
        if len(q) == 1 and ord(q) > 0x7F:          # 直接貼上字
            cp = ord(q)
        elif re.fullmatch(r"[0-9A-F]{4,6}", qu):
            cp = int(qu, 16)
        if cp is not None and self._offline_loaded():
            codes = self.offline.cp2codes.get(cp, [])
            if codes:
                return [self._code_result(c, sub=f"U+{cp:04X}") for c in codes], None
        if cp is not None:
            return builtin([c for c in COMMON_RARE_CHARS if ord(c["char"]) == cp]), None
        return builtin([c for c in COMMON_RARE_CHARS if qu in f"{ord(c['char']):04X}"]), None

    @staticmethod
    def _builtin_part(q: str, builtin):
        """沒有離線字庫時：只比對內建速查表的拆字說明"""
        parts = list(re.sub(r"[+＋\s]", "", q))

        def has_parts(c):
            hay = c["desc"] + c["radical"]
            return all(p in hay or (p == "王" and "玉" in hay) for p in parts)

        return builtin([c for c in COMMON_RARE_CHARS if has_parts(c)])

    def _radical_results(self, builtin):
        rad_text, stk_text, zq = self.q_radical_var.get(), self.q_strokes_var.get(), self.q_rzy_var.get().strip()
        if not self.offline.has_radicals():           # 沒有部首資料（舊版 cnsdata）：只查內建速查表
            rad = self._parse_radical(rad_text)
            if rad == "invalid":
                raise QueryInputError("找不到這個部首，請從清單選擇。")
            stk = self._parse_strokes(stk_text)
            zm = self._c_zhuyin(zq)[0] if zq else None
            return builtin([c for c in COMMON_RARE_CHARS
                            if (rad is None or c["radical"] == rad) and (stk is None or int(c["strokes"]) == stk)
                            and (zm is None or c["cns"] in zm)]), None
        got = self._filter_codes({"rad": rad_text, "stk": stk_text, "zhuyin": zq})
        if got is None:
            return builtin(COMMON_RARE_CHARS), "選擇部首、筆畫或輸入注音，就會從整個字庫查詢；下方先列出常見人名用字。"
        codes, zymap, notes = got
        return self._assemble(codes, zymap, sub=self._radical_sub, notes=notes,
                              more_hint="可再選筆畫或加注音縮小範圍（依筆畫排序）。")

    def _compound_results(self, builtin):
        v = {k: var.get() for k, var in self.c_vars.items()}
        got = self._filter_codes({"rad": self.c_rad_var.get(), "stk": self.c_stk_var.get(), **v})
        if got is None:
            return builtin(COMMON_RARE_CHARS), "填任一個條件就會從整個字庫查詢；條件越多範圍越小。下方先列出常見人名用字。"
        codes, zymap, notes = got
        order = "zhuyin" if (v["zhuyin"].strip() or v["pinyin"].strip()) else "strokes"
        return self._assemble(codes, zymap, order=order, notes=notes)

    def _code_result(self, code: str, zhuyin: str = None, sub: str = None) -> dict:
        zys = self.zhuyin_for_code(code)
        return {"cns": code, "char": self.char_for_code(code) or "?",
                "zhuyin": zhuyin if zhuyin is not None else (zys[0] if zys else ""), "sub": sub or code}

    # ---- 部件表 ----
    def open_component_picker(self, target_entry):
        """部件表：依筆畫列出全部部件（同一部件的不同寫法只列一個），點選即加到查詢框"""
        lib = self.offline
        if not lib.comp_info:
            busy = lib.available and not lib.extras_loaded.is_set()
            self.set_status("部件資料載入中，請稍候再開部件表。" if busy else "需要離線字庫（cnsdata）才能使用部件表。",
                            is_error=not busy, is_loading=busy)
            return
        dlg = self._make_dialog("部件表", "點一下部件就會加到查詢框，可以連續點選。同一部件的不同寫法（例如「火」與火字旁）算同一個。",
                                width=620)
        groups = {}
        for cid, (_cns, cp, g, strokes) in lib.comp_info.items():
            if strokes > 0 and (g not in groups or cid < groups[g][0]):
                groups[g] = (cid, cp, strokes)
        items = sorted(groups.values(), key=lambda t: (t[2], t[0]))

        wrap = tk.Frame(dlg.body, bg=C_CARD)
        wrap.pack(fill=tk.BOTH, expand=True, pady=(self.px(12), 0))
        canvas = tk.Canvas(wrap, bg=C_CARD, highlightthickness=0, bd=0, width=self.px(560), height=self.px(380))
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=canvas.yview, style="Slim.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        grid = tk.Frame(canvas, bg=C_CARD)
        canvas.create_window((0, 0), window=grid, anchor="nw")
        grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def wheel(e):
            canvas.yview_scroll(int(-e.delta / 120), "units")
            return "break"

        canvas.bind("<MouseWheel>", wheel)
        cols, row, col, last = 12, 0, 0, None
        picked = self._label(dlg.body, "", size=9, weight="bold", fg=C_SUBTLE, anchor="w")
        for cid, cp, strokes in items:
            if strokes != last:
                if col:
                    row += 1
                hdr = self._label(grid, f"{strokes} 畫", size=9, weight="bold", fg=C_MUTED, anchor="w")
                hdr.grid(row=row, column=0, columnspan=cols, sticky="w", pady=(self.px(8) if row else 0, self.px(2)))
                hdr.bind("<MouseWheel>", wheel)
                row, col, last = row + 1, 0, strokes
            cell = tk.Label(grid, text=chr(cp), font=(self.font_plus_family, 18), bg=C_CARD, fg=C_FG,
                            width=2, cursor="hand2", highlightthickness=1, highlightbackground=C_BORDER)
            cell.grid(row=row, column=col, padx=self.px(2), pady=self.px(2), sticky="nsew")
            cell.bind("<Enter>", lambda e, w=cell: w.config(bg=C_HOVER))
            cell.bind("<Leave>", lambda e, w=cell: w.config(bg=C_CARD))
            cell.bind("<Button-1>", lambda e, ch=chr(cp): (self._insert_text(target_entry, ch),
                                                           picked.config(text=f"已加入：{target_entry.get()}")))
            cell.bind("<MouseWheel>", wheel)
            col += 1
            if col >= cols:
                row, col = row + 1, 0
        picked.config(font=(self.font_plus_family, 11, "bold"))
        picked.pack(fill=tk.X, pady=(self.px(10), 0))
        btns = tk.Frame(dlg.body, bg=C_CARD)
        btns.pack(fill=tk.X, pady=(self.px(10), 0))
        FlatButton(btns, "完成", command=dlg.destroy).pack(side=tk.RIGHT)
        self._center_dialog(dlg)

    # ---- 結果字卡 ----
    def _query_thumb(self, code: str):
        size = self.px(36)
        if code in self._q_thumbs:
            return self._q_thumbs[code]
        photo = None
        if self._offline_loaded():
            img = self.offline.thumb(code, size)
            if img is not None:
                photo = ImageTk.PhotoImage(img)
        if photo is not None or self._offline_loaded():
            self._q_thumbs[code] = photo      # 離線字庫還在載入時不快取，載入後再畫
        return photo

    def refresh_query_results(self):
        if self._q_job:
            self.root.after_cancel(self._q_job)
            self._q_job = None
        results, note = self.compute_query_results()
        self._q_results = results
        for w in self.q_grid.winfo_children():
            w.destroy()
        self._q_cards = [self._make_query_card(r) for r in results]
        self._layout_query_cards(force=True)
        if not results and not note:
            note = "沒有符合的字，可按「到官網查詢」用全字庫完整查詢。" if self._offline_loaded() \
                else "速查表中沒有符合的字，可按「到官網查詢」用全字庫完整查詢。"
        if note:
            self.q_note_lbl.config(text=note)
            self.q_note_lbl.pack(fill=tk.X, pady=(self.px(8) if results else 0, 0))
        else:
            self.q_note_lbl.pack_forget()
        self._update_load_btn()

    def _make_query_card(self, r: dict):
        border = tk.Frame(self.q_grid, bg=C_BORDER, padx=1, pady=1)
        card = tk.Frame(border, bg=C_CARD)
        card.pack(fill=tk.BOTH, expand=True)
        photo = self._query_thumb(r["cns"])
        if photo is not None:
            glyph = tk.Label(card, image=photo, bg=C_CARD)
        else:
            glyph = tk.Label(card, text=r["char"], bg=C_CARD, fg=C_FG, font=(FONT_KAI, 22))
        glyph.pack(pady=(self.px(8), 0))
        zy = tk.Label(card, text=r["zhuyin"], bg=C_CARD, fg=C_FG, font=(FONT_UI, 9, "bold"))
        zy.pack()
        sub = tk.Label(card, text=r["sub"], bg=C_CARD, fg=C_SUBTLE, font=(FONT_MONO, 8, "bold"))
        sub.pack(pady=(0, self.px(8)))
        widgets = (card, glyph, zy, sub)

        def paint(color):
            for w in widgets:
                w.config(bg=color)

        for w in widgets:
            w.config(cursor="hand2")
            w.bind("<Enter>", lambda e: paint(C_HOVER))
            w.bind("<Leave>", lambda e: paint(C_CARD))
            w.bind("<Button-1>", lambda e, code=r["cns"]: self._pick_query_result(code))
        return border

    def _layout_query_cards(self, force: bool = False):
        """字卡寬度至少約 84px，依可用寬度決定每列幾張"""
        width = self.q_grid.winfo_width()
        if width <= 1:
            width = self.px(700)
        cols = max(4, min(10, (width + self.px(8)) // self.px(92)))
        if cols == self._q_cols and not force:
            return
        for i in range(max(cols, self._q_cols, 10)):
            self.q_grid.columnconfigure(i, weight=1 if i < cols else 0, uniform="q" if i < cols else "")
        self._q_cols = cols
        for i, card in enumerate(getattr(self, "_q_cards", [])):
            card.grid(row=i // cols, column=i % cols, sticky="nsew",
                      padx=(0 if i % cols == 0 else self.px(8), 0), pady=(0 if i < cols else self.px(8), 0))

    def _update_load_btn(self):
        ok = bool(self._q_results) or (self.query_mode == "cns" and bool(self.parse_cns_code(self.code_var.get())))
        self.fetch_btn.set_enabled(not self.is_fetching and ok)

    def _pick_query_result(self, code: str):
        if self.is_fetching:
            self.set_status("字形下載中，請稍候再選。", is_loading=True)
            return
        self.code_var.set(code)
        self.fetch_character()

    def query_load(self):
        """「載入」：CNS 模式輸入單一代碼時直接載入；其他情況載入第一個結果"""
        if self._q_job:
            self.refresh_query_results()
        ci = self.code_var.get().strip()
        if self.query_mode == "cns" and (not ci or self.parse_cns_code(ci) or not self._q_results):
            self.fetch_character()
        elif self._q_results:
            self._pick_query_result(self._q_results[0]["cns"])

    def _update_meta(self):
        """預覽區右欄的「標楷體・20 pt」"""
        style_short = "宋體" if "宋體" in self.font_style_var.get() else "標楷體"
        self.meta_lbl.config(text=f"{style_short}・{self.font_size_var.get()} pt")

    def _set_fetching_ui(self, fetching: bool, std_code: str = ""):
        """切換下載中狀態：按鈕文字、預覽畫布、狀態徽章"""
        self.fetch_btn.set_text("下載中" if fetching else "載入")
        self._update_load_btn()
        if fetching:
            self.draw_grid_lines()
            self.preview_canvas.create_text(
                self.preview_size / 2, self.preview_size / 2, text="下載中…",
                fill="#94A3B8", font=(FONT_UI, 11, "bold")
            )
            self.cns_code_lbl.config(text=std_code or "—")
            if self.zhuyin_lbl.winfo_ismapped():
                self.zhuyin_lbl.pack_forget()      # 載入完成後依新字重新顯示
            self._set_badge("ok", "下載中…")
        self.update_copy_state()

    def get_copy_state(self):
        """回傳 (可否複製, 按鈕下方提示文字)"""
        if self.is_fetching:
            return False, "字形下載中，請稍候。"
        if not self.current_pil_image or not self.current_cns_code:
            return False, "請先輸入代碼並按「載入」。"
        if self.current_is_fallback:
            return False, "這不是全字庫字形，請連線後重新下載。"
        if self.current_source == "hand":
            return True, "手繪字：複製後到 Word 將「?」反白，按 Ctrl+V 貼上。"
        parsed = self.parse_cns_code(self.code_var.get())
        style_changed = ("宋體" in self.font_style_var.get()) != self.current_style_is_sung
        if (parsed and parsed[2] != self.current_cns_code) or style_changed:
            return False, "代碼或字體已變更，請先按「載入」。"
        return True, "複製後到 Word 將「?」反白，按 Ctrl+V 貼上，即成為對齊的行內字。"

    def update_copy_state(self):
        ok, hint = self.get_copy_state()
        self.copy_btn.set_enabled(ok)
        self.copy_hint_lbl.config(text=hint)

    def _flash_copied(self):
        """複製成功後，按鈕短暫顯示「已複製」"""
        self.copy_btn.set_text("✓  已複製")
        if self._copied_reset_job:
            self.root.after_cancel(self._copied_reset_job)
        self._copied_reset_job = self.root.after(1800, lambda: self.copy_btn.set_text("複製此字"))

    # --------------------------------------------------------------------------
    # 預覽繪製
    # --------------------------------------------------------------------------
    def draw_grid_lines(self):
        """在預覽畫布繪製米字格（紅色虛線），輔助檢視字面比例與重心"""
        c, s = self.preview_canvas, self.preview_size
        c.delete("all")
        c.create_line(0, s / 2, s, s / 2, fill="#FCA5A5", dash=(4, 3))
        c.create_line(s / 2, 0, s / 2, s, fill="#FCA5A5", dash=(4, 3))
        c.create_line(0, 0, s, s, fill="#FEE2E2", dash=(4, 3))
        c.create_line(s, 0, 0, s, fill="#FEE2E2", dash=(4, 3))

    def _draw_preview_placeholder(self, code_text: str):
        """尚無字形時，畫一個灰色代碼方框"""
        c, s = self.preview_canvas, self.preview_size
        tid = c.create_text(s / 2, s / 2, text=code_text, fill="#94A3B8", font=(FONT_MONO, 13, "bold"))
        x1, y1, x2, y2 = c.bbox(tid)
        pad_x, pad_y = self.px(12), self.px(14)
        c.create_rectangle(x1 - pad_x, y1 - pad_y, x2 + pad_x, y2 + pad_y, outline=C_INPUT, width=2)

    def update_preview_display(self, from_cache: bool = False):
        """更新米字格預覽、代碼與狀態徽章"""
        self.draw_grid_lines()
        if not self.current_pil_image:
            self._draw_preview_placeholder(self.current_cns_code or "—")
            self.cns_code_lbl.config(text=self.current_cns_code or "—")
            self._update_zhuyin_line()
            self._set_badge("ok", "尚未載入")
            self.update_copy_state()
            return

        glyph_px = self.px(148)
        preview_resized = self.current_pil_image.resize((glyph_px, glyph_px), Image.Resampling.LANCZOS)
        self.current_tk_image = ImageTk.PhotoImage(preview_resized)
        self.preview_canvas.create_image(self.preview_size / 2, self.preview_size / 2, image=self.current_tk_image)
        self.cns_code_lbl.config(text=self.current_cns_code)
        self._update_zhuyin_line()

        badge = {
            "fallback": ("bad", "備援圖・不可複製"),
            "sung-sub": ("warn", "宋體替代"),
            "cache": ("ok", "本機快取"),
            "online": ("ok", "已下載並快取"),
            "hand": ("warn", "手繪字"),
            "offline": ("ok", "離線字庫"),
        }.get(self.current_source, ("ok", "已載入"))
        self._set_badge(*badge)
        self.update_copy_state()

    # --------------------------------------------------------------------------
    # 使用紀錄
    # --------------------------------------------------------------------------
    @staticmethod
    def _history_iid(code: str, style: str) -> str:
        return f"{code}|{style}"

    @staticmethod
    def _split_iid(iid: str):
        code, _, style = iid.partition("|")
        return code, style

    def _history_thumb(self, code: str, is_sung: bool):
        """取得歷史紀錄的字形縮圖（楷體優先用離線字庫，其次本機快取）"""
        size = self.px(28)
        if not is_sung and self.offline.available and self.offline._loaded.is_set():
            small = self._offline_thumbs.get(code)
            if small is None and code not in self._offline_thumbs:
                img = self.offline.render(code)
                small = img.resize((size, size), Image.Resampling.LANCZOS) if img is not None else None
                self._offline_thumbs[code] = small
            if small is not None:
                photo = ImageTk.PhotoImage(small)
                self._history_thumbs.append(photo)
                return photo
        path = self.get_cache_filepath(code, is_sung)
        if not (HAS_PIL_REQ and os.path.exists(path)):
            return None
        try:
            img = self._open_cache_image(path)
            if img is None:
                return None
            img = img.resize((size, size), Image.Resampling.LANCZOS)
            bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
            photo = ImageTk.PhotoImage(Image.alpha_composite(bg, img))
            self._history_thumbs.append(photo)
            return photo
        except Exception:
            return None

    def refresh_history_tree(self):
        """重新整理使用紀錄列表（保留原本的選取項目）"""
        sel = self.history_tree.selection()
        keep_iid = sel[0] if sel else None

        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        self._history_thumbs = []

        for item in self.history_data:
            code = item.get("cns_code", "")
            style_str = item.get("font_style", STYLE_KAI)
            iid = self._history_iid(code, style_str)
            if not code or self.history_tree.exists(iid):
                continue
            is_sung = "宋體" in style_str
            offline_ready = self.offline.available and self.offline._loaded.is_set()
            if not is_sung and offline_ready and self.offline.has(code):
                cache_label = "離線"
            elif os.path.exists(self.get_cache_filepath(code, is_sung)):
                cache_label = "已快取"
            else:
                cache_label = "需下載"
            thumb = self._history_thumb(code, is_sung)
            char = self.char_for_code(code) or "?"
            kwargs = {"image": thumb, "text": ""} if thumb else {"text": char}
            zys = self.zhuyin_for_code(code)
            self.history_tree.insert(
                "", tk.END, iid=iid, tags=("row",),
                values=(
                    f"{code}\n{zys[0]}" if zys else code, "宋體" if is_sung else "標楷體",
                    item.get("font_size", "20 pt"), item.get("last_used", ""),
                    cache_label, "✕",
                ),
                **kwargs,
            )

        if keep_iid and self.history_tree.exists(keep_iid):
            # 隨後觸發的 <<TreeviewSelect>> 因代碼與字體相同不會重複載入
            self.history_tree.selection_set(keep_iid)

        empty = not self.history_tree.get_children()
        if empty:
            self.history_empty_lbl.place(relx=0.5, rely=0.6, anchor="center")
        else:
            self.history_empty_lbl.place_forget()
        self.clear_btn.set_enabled(not empty)

    def _is_delete_cell(self, event) -> bool:
        return (self.history_tree.identify_region(event.x, event.y) == "cell"
                and self.history_tree.identify_column(event.x) == "#6")

    def _on_history_motion(self, event):
        self.history_tree.config(cursor="hand2" if self._is_delete_cell(event) else "")

    def _on_history_click(self, event):
        """點到每列最右側的 ✕ 時刪除該筆，不觸發載入"""
        if self._is_delete_cell(event):
            iid = self.history_tree.identify_row(event.y)
            if iid:
                self.delete_history_item(iid)
            return "break"
        return None

    def on_history_select(self, event):
        """點選使用紀錄時，以當時的字體載入該字形"""
        selected = self.history_tree.selection()
        if not selected:
            return
        code, style = self._split_iid(selected[0])
        style_changed = self._apply_history_style(style)
        if style_changed or code != self.current_cns_code or self.current_pil_image is None:
            self.code_var.set(code)
            self.fetch_character()

    def on_history_double_click(self, event):
        """雙擊使用紀錄時，載入完成後自動複製到剪貼簿"""
        if self._is_delete_cell(event):
            return "break"
        iid = self.history_tree.identify_row(event.y)
        if not iid:
            return None
        code, style = self._split_iid(iid)
        self._apply_history_style(style)
        self.code_var.set(code)
        self.copy_after_fetch = True   # 字形載入完成後才複製
        if self.is_fetching:
            return None                # 單擊已觸發下載，完成後會自動複製
        self.fetch_character()
        return None

    def delete_history_item(self, iid: str):
        code, style = self._split_iid(iid)
        self.history_data = [
            item for item in self.history_data
            if not (item.get("cns_code") == code and item.get("font_style", STYLE_KAI) == style)
        ]
        self.save_history()
        self.refresh_history_tree()
        self.set_status(f"已移除紀錄 [{code}]")

    def delete_selected_history(self):
        """刪除選取的單筆使用紀錄（鍵盤 Delete）"""
        selected = self.history_tree.selection()
        if selected:
            self.delete_history_item(selected[0])

    def clear_all_history(self):
        """確認後清空所有使用紀錄"""
        if not self.history_data:
            return
        dlg = self._make_dialog("清除所有使用紀錄？", "紀錄清單會清空，本機快取的字圖不受影響。", width=420)
        btns = tk.Frame(dlg.body, bg=C_CARD)
        btns.pack(fill=tk.X, pady=(self.px(20), 0))

        def do_clear():
            dlg.destroy()
            self.history_data = []
            self.save_history()
            self.refresh_history_tree()
            self.set_status("已清除所有使用紀錄")

        FlatButton(btns, "清除", variant="destructive", command=do_clear).pack(side=tk.RIGHT)
        FlatButton(btns, "取消", variant="outline", command=dlg.destroy).pack(side=tk.RIGHT, padx=(0, self.px(8)))
        self._center_dialog(dlg)

    # --------------------------------------------------------------------------
    # 對話框
    # --------------------------------------------------------------------------
    def _make_dialog(self, title: str, description: str, width: int = 480):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg=C_CARD)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.body = tk.Frame(dlg, bg=C_CARD, padx=self.px(24), pady=self.px(24), width=self.px(width))
        dlg.body.pack(fill=tk.BOTH, expand=True)
        self._label(dlg.body, title, size=13, weight="bold").pack(anchor="w")
        self._label(dlg.body, description, size=10, fg=C_MUTED, justify="left",
                    wraplength=self.px(width - 48)).pack(anchor="w", pady=(self.px(6), 0))
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        return dlg

    def _center_dialog(self, dlg):
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        dlg.grab_set()
        dlg.focus_set()

    def _poll_ui_queue(self):
        """在主執行緒執行背景執行緒送來的 UI 回呼"""
        try:
            while True:
                callback = self.ui_queue.get_nowait()
                try:
                    callback()
                except Exception as e:
                    # 單一回呼失敗不可中斷輪詢，並解除下載鎖定
                    self.is_fetching = False
                    self.copy_after_fetch = False
                    try:
                        self._set_fetching_ui(False)
                        self.update_preview_display()
                        self.set_status(f"處理下載結果時發生錯誤：{e}", is_error=True)
                    except Exception:
                        pass
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_ui_queue)

    def check_environment(self):
        """檢查執行環境並顯示友善提示"""
        missing = []
        if not HAS_PIL_REQ:
            missing.append("pillow / requests")
        if not HAS_WIN32 and os.name == "nt":
            missing.append("pywin32")

        if missing:
            self.set_status(
                f"提示：尚未安裝 {', '.join(missing)}。請執行 pip install requests pillow pywin32 以啟用完整功能。",
                is_error=True
            )
        else:
            # 啟動時預設載入範例字或最近一筆紀錄
            if self.history_data:
                first_cns = self.history_data[0].get("cns_code", "3-3476")
                self.code_var.set(first_cns)
            self.root.after(250, self.fetch_character)

    # ==========================================================================
    # 全字庫代碼解析與網路非同步抓取 / 本機快取核心邏輯
    # ==========================================================================
    @staticmethod
    def clean_input_string(raw_input: str) -> str:
        """自動將全形數字、全形英文字母與各種破折號轉為標準半形字元，並去除隱形空白"""
        fullwidth_map = str.maketrans(
            "０１２３４５６７８９ＡＢＣＤＥＦａｂｃｄｅｆ－—–─＿：",
            "0123456789ABCDEFabcdef----_:"
        )
        cleaned = raw_input.translate(fullwidth_map)
        # 移除零寬度空白或特殊控制字元
        cleaned = re.sub(r"[\u200b-\u200f\ufeff]", "", cleaned)
        return cleaned.strip().upper()

    def parse_cns_code(self, raw_input: str):
        """
        解析使用者輸入的全字庫代碼（全面支援第 1 ~ 19 字面及十進位/十六進位表示法）。
        支援格式：
          - 常規與深層十進位字面："1-4442", "12-6170", "15-2121"
          - 全形輸入："１２－６１７０" (自動清洗轉半形)
          - 十六進位字面："C-6170" (自動轉為 12-6170), "F-2121" (轉為 15-2121)
          - 含前綴或特殊分隔符："CNS 12-6170", "CNS11643 01-4442", "12_6170"
          - 緊湊格式："126170", "014442", "C6170"
        回傳：(plane_int, hex_code_str, standard_code_str) 或 None
        """
        cleaned = self.clean_input_string(raw_input)
        cleaned = re.sub(r"^CNS(11643)?\s*[-_:\s]*", "", cleaned)

        # 1. 匹配「字面 + 分隔符 + 四位十六進位碼」(例如 12-6170 或 C-6170)
        match = re.match(r"^([0-9A-F]{1,2})[-_\s:]+([0-9A-F]{4})$", cleaned)
        if match:
            plane_str, hex_part = match.group(1), match.group(2)
            plane = self._parse_plane_number(plane_str)
            if plane and 1 <= plane <= 19:
                return plane, hex_part, f"{plane}-{hex_part}"

        # 2. 匹配無分隔符的緊湊格式 (例如 126170、014442、14442、C6170)
        match_compact = re.match(r"^([0-9A-F]{1,2})([0-9A-F]{4})$", cleaned)
        if match_compact:
            plane_str, hex_part = match_compact.group(1), match_compact.group(2)
            plane = self._parse_plane_number(plane_str)
            if plane and 1 <= plane <= 19:
                return plane, hex_part, f"{plane}-{hex_part}"

        return None

    @staticmethod
    def _parse_plane_number(plane_str: str):
        """
        將字面字串轉為十進位整數 (1~19)。
        若輸入為純數字 (如 "12") 優先視為十進位第 12 字面；
        若輸入為 A~F (如 "C") 則視為十六進位轉為 12。
        """
        if plane_str.isdigit():
            val = int(plane_str, 10)
            if 1 <= val <= 19:
                return val
        try:
            val_hex = int(plane_str, 16)
            if 1 <= val_hex <= 19:
                return val_hex
        except ValueError:
            pass
        return None

    @staticmethod
    def _open_cache_image(path: str):
        """讀取快取字圖：先檢查尺寸再解碼（防「解壓縮炸彈」），異常的快取檔直接刪除，之後重新下載"""
        img = Image.open(path)
        if img.width > 2048 or img.height > 2048:
            img.close()
            try:
                os.remove(path)
            except OSError:
                pass
            return None
        return img.convert("RGBA")

    def get_cache_filepath(self, standard_code: str, is_sung: bool = False) -> str:
        """取得該全字庫代碼對應的本機快取 PNG 路徑"""
        if not self._is_std_code(standard_code):
            raise ValueError(f"不合法的全字庫代碼：{standard_code!r}")   # 防止路徑穿越
        style_tag = "sung" if is_sung else "kai"
        safe_code = standard_code.replace("-", "_")
        return os.path.join(self.cache_dir, f"cns_{style_tag}_{safe_code}.png")

    def fetch_character(self):
        """
        點擊「下載 / 載入字形」時觸發：
          1. 自動清洗輸入並驗證格式。
          2. 若本機快取已有檔案，秒速讀取並顯示。
          3. 若需連線下載，啟動背景執行緒 (threading.Thread) 非同步抓取，防止視窗凍結沒反應。
        """
        if self.is_fetching:
            self.set_status("正在下載字形中，請稍候完成後再操作。", is_loading=True)
            return

        if not HAS_PIL_REQ:
            self.set_status("缺少 Pillow 或 Requests 套件，請先執行 pip install requests pillow", is_error=True)
            return

        raw_code = self.code_var.get()
        parsed = self.parse_cns_code(raw_code)
        if not parsed:
            self.copy_after_fetch = False
            self._show_input_error(f"代碼格式不正確：「{raw_code}」。請輸入如 1-4442 或 12-6170。")
            self.set_status("代碼格式不正確", is_error=True)
            return

        plane, hex_code, std_code = parsed
        self.code_var.set(std_code)
        # 注意：current_cns_code 要等字圖真正載入後才更新，避免代碼與字圖不一致

        is_sung = "宋體" in self.font_style_var.get()
        cache_path = self.get_cache_filepath(std_code, is_sung)

        # 第一層（楷體）：離線字庫直接繪製，不需網路（約 10 毫秒）。
        # 放在快取之前：日後更新 cnsdata（官網修正字形）時，以新版字庫為準。
        if not is_sung and self.offline.available:
            offline_img = self.offline.render(std_code)
            if offline_img is not None:
                self.current_pil_image = self.normalize_glyph_image(offline_img.convert("RGBA"))
                self.current_cns_code = std_code
                self.current_is_fallback = False
                self.current_style_is_sung = False
                self.current_source = "offline"
                self.update_preview_display()
                self.refresh_history_tree()
                ch = self.offline.char_for(std_code)
                self.set_status(f"已從離線字庫載入 [{std_code}]，字：{self._describe_char(ch) if ch else '—'}")
                self._run_pending_copy()
                return

        # 第二層：檢查本機快取是否已存在 (零延遲載入)
        if os.path.exists(cache_path):
            try:
                # 快取檔存的已是正規化後的字圖，不再重複處理（避免筆畫被二次加濃）
                img = self._open_cache_image(cache_path)
                if img is None:
                    raise ValueError("快取字圖異常")
                if img.size != (512, 512):
                    img = img.resize((512, 512), Image.Resampling.LANCZOS)
                self.current_pil_image = img
                self.current_cns_code = std_code
                self.current_is_fallback = False
                self.current_style_is_sung = is_sung
                self.current_source = "cache"
                self.update_preview_display(from_cache=True)
                self.refresh_history_tree()
                self.set_status(f"已從本機快取載入 [{std_code}]")
                self._run_pending_copy()
                return
            except Exception:
                pass

        # 第三層：啟動背景執行緒連線下載（宋體、離線字庫沒有的字），確保 UI 絕不卡頓
        self.is_fetching = True
        self.pending_style_is_sung = is_sung
        self._set_fetching_ui(True, std_code)
        self.set_status(f"正在連線全字庫下載 [{std_code}]（第 {plane} 字面）…", is_loading=True)

        worker = threading.Thread(
            target=self._background_fetch_worker,
            args=(plane, hex_code, std_code, is_sung, cache_path),
            daemon=True
        )
        worker.start()

    def _background_fetch_worker(self, plane: int, hex_code: str, std_code: str, is_sung: bool, cache_path: str):
        """在背景執行緒中進行網路請求與影像標準化處理"""
        try:
            downloaded_img, used_fallback_sung, source_name = self.download_with_fallback(plane, hex_code, is_sung)

            if downloaded_img is not None:
                norm_img = self.normalize_glyph_image(downloaded_img)
                # 依「實際取得的字體」存檔：宋體替代品不可存成楷體快取
                save_path = self.get_cache_filepath(std_code, is_sung or used_fallback_sung)
                try:
                    norm_img.save(save_path, "PNG")
                except Exception:
                    pass
                self.ui_queue.put(lambda: self._on_fetch_complete(
                    std_code, norm_img, True, used_fallback_sung, source_name
                ))
            else:
                # 若遠端均無回應，使用本機標楷體備援繪製。
                # 備援圖「不」寫入快取，避免日後連線時仍讀到假字圖。
                fallback_img = self.generate_fallback_glyph(std_code)
                self.ui_queue.put(lambda: self._on_fetch_complete(
                    std_code, fallback_img, False, False, "本機備援字圖"
                ))
        except LookupError as err:
            msg = str(err)
            self.ui_queue.put(lambda: self._on_fetch_error(std_code, msg, not_found=True))
        except Exception as err:
            msg = f"下載 [{std_code}] 發生錯誤：{err}"
            self.ui_queue.put(lambda: self._on_fetch_error(std_code, msg))

    def _on_fetch_complete(self, std_code: str, norm_img: Image.Image, is_online_success: bool,
                           used_fallback_sung: bool, source_name: str):
        """背景下載完成後，回到主執行緒更新 UI 畫面"""
        self.is_fetching = False

        self.current_pil_image = norm_img
        self.current_cns_code = std_code
        self.current_is_fallback = not is_online_success
        self.current_style_is_sung = self.pending_style_is_sung
        if not is_online_success:
            self.current_source = "fallback"
        elif used_fallback_sung:
            self.current_source = "sung-sub"
        else:
            self.current_source = "online"
        self._set_fetching_ui(False)
        self.update_preview_display(from_cache=False)
        self.refresh_history_tree()

        if is_online_success:
            if used_fallback_sung:
                self.set_status(
                    f"[{std_code}] 找不到楷書，目前是宋體字形（來源：{source_name}），與公文標楷體不同。",
                    is_error=True
                )
            else:
                self.set_status(f"已下載 [{std_code}] 並存入本機快取（來源：{source_name}）")
        else:
            self.copy_after_fetch = False
            if self.offline.available and not self.offline.has(std_code) and not self.pending_style_is_sung:
                msg = (f"離線字庫（版本 {self.offline.version or '—'}）查無 [{std_code}]，也無法連線官網確認。"
                       "請確認代碼，或連上網路後重試。")
            else:
                msg = f"遠端字庫無回應，[{std_code}] 僅有預覽用備援圖，已停用複製。請檢查網路後重試。"
            self.set_status(
                msg,
                is_error=True
            )
            return

        self._run_pending_copy()

    def _run_pending_copy(self):
        """雙擊歷史紀錄後，字形載入完成才執行複製"""
        if self.copy_after_fetch:
            self.copy_after_fetch = False
            self.copy_to_word_clipboard()

    def _on_fetch_error(self, std_code: str, err_msg: str, not_found: bool = False):
        """背景下載失敗（查無此字或其他錯誤）時的 UI 回饋"""
        self.copy_after_fetch = False
        self.is_fetching = False
        self._set_fetching_ui(False)
        self.update_preview_display()
        if not_found:
            self._show_input_error(err_msg)
        self.set_status(err_msg, is_error=True)

    def download_with_fallback(self, plane: int, hex_code: str, is_sung: bool):
        """
        向全字庫官網下載字圖；選標楷體但官網沒有楷書字形時，自動改抓宋體。
        回傳：(PIL.Image 或 None, 是否使用了宋體替代 bool, 來源說明 str)
          - 回傳 None：連不上全字庫（呼叫端會產生不可複製的備援預覽圖）
          - 拋出 LookupError：連得上，但全字庫查無此代碼或此字無字圖
        """
        _ensure_requests()
        session = _get_shared_session()
        try:
            info = self._lookup_cns_char(session, plane, hex_code)
            if info is None:
                return None, False, ""
            char, uni_hex = info
            label = f"全字庫官網，字：{self._describe_char(char)}"

            img = self._download_cns_glyph(session, uni_hex, "sung" if is_sung else "kai")
            if img is not None:
                return img, False, label

            if not is_sung:
                sung_img = self._download_cns_glyph(session, uni_hex, "sung")
                if sung_img is not None:
                    return sung_img, True, label
        except TransientServerError as err:
            raise RuntimeError(f"全字庫官網暫時忙碌（{err}），請稍後再按「載入」") from None
        except requests.exceptions.SSLError:
            # 憑證驗證失敗：可能遭中間人攔截或網路被導向假網站，絕不可退回不驗證
            _reset_shared_session()
            raise RuntimeError("全字庫官網的安全憑證驗證失敗，已停止下載（可能遭攔截）。請洽資訊人員") from None
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError) as err:
            _reset_shared_session()
            if "CERTIFICATE_VERIFY_FAILED" in str(err) or "SSLError" in repr(err):
                raise RuntimeError("全字庫官網的安全憑證驗證失敗，已停止下載（可能遭攔截）。請洽資訊人員") from None
            return None, False, ""   # 重試後仍連不上

        raise LookupError(f"全字庫有 [{plane}-{hex_code}]，但官網沒有提供這個字的字圖")

    CNS_BASE = "https://www.cns11643.gov.tw"
    HTTP_HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
    }
    # 官網對不存在的字形仍回傳 200 與佔位圖（「豆腐框」與「1」），需排除
    PLACEHOLDER_MD5 = {
        "2dc7ad617edc65d439c2405e50c7f599",
        "72850ff1b5d5efc09501d141605f07cc",
    }

    @staticmethod
    def _describe_char(char: str) -> str:
        """私用區字元（如戶政第 12 字面）無法正常顯示，改以說明文字呈現"""
        cp = ord(char) if len(char) == 1 else 0
        if 0xE000 <= cp <= 0xF8FF or cp >= 0xF0000:
            return f"U+{cp:X}（私用區字，需用字圖）"
        return f"「{char}」"

    def _lookup_cns_char(self, session, plane: int, hex_code: str):
        """
        查詢全字庫字形資訊頁，取得 (字元, Unicode 十六進位小寫)。
        字形資訊頁網址：wordView.jsp?ID=字面×65536+碼位
        連線失敗回傳 None；查無此字拋出 LookupError。
        """
        word_id = plane * 65536 + int(hex_code, 16)
        url = f"{self.CNS_BASE}/wordView.jsp?ID={word_id}"
        # 連線/憑證錯誤不在此攔截，交由 download_with_fallback 統一判斷
        status, body = self._http_get(session, url, max_bytes=self.MAX_PAGE_BYTES)
        if status != 200:
            raise TransientServerError(status)     # 不能當成「查無此字」或「離線」
        text = body.decode("utf-8", errors="replace")
        title = re.search(r"<title>\s*字形資訊\s*-\s*\[(.{1,2}?)\]\s*([0-9A-Fa-f]{1,2}-[0-9A-Fa-f]{4})", text)
        if not title or title.group(2).upper() != f"{plane}-{hex_code}".upper():
            raise LookupError(f"全字庫查無 [{plane}-{hex_code}] 這個代碼，請確認字面與碼位")
        img_path = re.search(r"/char/(?:kai|sung)/\d+/[0-9a-fA-F]{2,4}__/([0-9a-fA-F]{4,6})\.png", text)
        if not img_path:
            raise LookupError(f"全字庫有 [{plane}-{hex_code}]，但官網沒有提供這個字的字圖")
        return title.group(1), img_path.group(1).lower()

    def _download_cns_glyph(self, session, uni_hex: str, face: str):
        """下載官網 512px 字圖：/char/{kai|sung}/512/{前綴}__/{unicode}.png"""
        if not re.fullmatch(r"[0-9a-f]{4,6}", uni_hex) or face not in ("kai", "sung"):
            return None
        url = f"{self.CNS_BASE}/char/{face}/512/{uni_hex[:-2]}__/{uni_hex}.png"
        status, data = self._http_get(session, url, max_bytes=self.MAX_IMAGE_BYTES)
        if status == 404:
            return None                            # 確定沒有這個字體的字圖
        if status != 200:
            # 429/5xx 等暫時性錯誤（重試後仍失敗）：不可誤判為「沒有楷書」而改用宋體
            raise TransientServerError(status)
        if data[:8] != b"\x89PNG\r\n\x1a\n":
            return None
        if hashlib.md5(data).hexdigest() in self.PLACEHOLDER_MD5:
            return None
        try:
            img = Image.open(io.BytesIO(data))
            # 先檢查尺寸再解碼，避免惡意「解壓縮炸彈」圖片耗盡記憶體
            if img.width > 2048 or img.height > 2048:
                return None
            img = img.convert("RGBA")
        except Exception:
            return None
        return img if self._is_valid_glyph_image(img) else None

    MAX_PAGE_BYTES = 512 * 1024      # 字形資訊頁實際約 12KB
    MAX_IMAGE_BYTES = 1024 * 1024    # 512px 字圖實際約 3~8KB

    def _http_get(self, session, url: str, max_bytes: int, want_location: bool = False):
        """
        受限的 GET：只連全字庫官網、不跟隨轉址、限制回應大小。
        回傳 (HTTP 狀態碼, 內容 bytes)；轉址一律視為失敗（狀態碼照實回傳）。
        want_location=True 時回傳 (狀態碼, 內容, 轉址目標字串或 None)，只讀取不跟隨。
        """
        if not url.startswith(self.CNS_BASE + "/"):
            raise ValueError("只允許連線全字庫官網")
        # 連線逾時含 TLS 握手（官網偶爾 3~5 秒），讀取逾時為每次收資料的等待上限。
        # 標頭階段的失敗由 Retry 自動重試；內容傳到一半中斷則在這裡再重試一次。
        body_errors = (requests.exceptions.ChunkedEncodingError,
                       requests.exceptions.ConnectionError,
                       requests.exceptions.Timeout)
        for attempt in range(2):
            with session.get(url, headers=self.HTTP_HEADERS, timeout=self.HTTP_TIMEOUT,
                             allow_redirects=False, stream=True) as resp:
                location = resp.headers.get("Location") if resp.is_redirect else None
                done = (lambda st, body: (st, body, location)) if want_location else (lambda st, body: (st, body))
                if resp.status_code != 200:
                    return done(resp.status_code, b"")
                declared = resp.headers.get("Content-Length")
                if declared and declared.isdigit() and int(declared) > max_bytes:
                    return done(413, b"")
                chunks, total = [], 0
                try:
                    for chunk in resp.iter_content(chunk_size=16384):
                        total += len(chunk)
                        if total > max_bytes:
                            return done(413, b"")
                        chunks.append(chunk)
                except body_errors:
                    if attempt == 1:
                        raise
                    time.sleep(0.5)
                    continue
                return done(200, b"".join(chunks))

    HTTP_TIMEOUT = (6.0, 10.0)   # (連線含 TLS 握手, 讀取)

    @staticmethod
    def _is_valid_glyph_image(img: Image.Image) -> bool:
        """檢查下載的圖片是否確實含有筆畫像素（排除空白佔位圖）"""
        try:
            # 先將透明背景合成白底再檢測灰階極值，支援透明底黑字 PNG 與白底黑字 PNG
            bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
            composited = Image.alpha_composite(bg, img.convert("RGBA")).convert("L")
            extrema = composited.getextrema()
            if extrema and (extrema[1] - extrema[0] < 15):
                return False
            return True
        except Exception:
            return True

    def normalize_glyph_image(self, img: Image.Image) -> Image.Image:
        """
        標準化字形圖片：
        1. 自動相容「白底黑字」與「透明底黑字」圖片，將背景統一去背並強化黑色筆畫對比度。
        2. 統一輸出為 512x512 標準 Em-Square 正方形，確保縮小為 20pt 時與標楷體字面率完全一致。
        """
        target_size = 512
        # 先合成於純白背景上，統一計算亮度，避免透明底黑字 (RGB=0,0,0) 被誤判
        white_canvas = Image.new("RGBA", img.size, (255, 255, 255, 255))
        composited = Image.alpha_composite(white_canvas, img.convert("RGBA"))

        # 亮度 > 235（淺色背景）→ 全透明；其餘強化筆畫濃度，列印時邊緣銳利無灰階毛邊。
        # 以查表（point）一次處理整張圖，取代逐像素迴圈（512×512 由約 0.3 秒降到數毫秒）。
        lum = composited.convert("RGB").convert("L")
        alpha_lut = [0 if v > 235 else max(0, min(255, int((255 - v) * 1.15))) for v in range(256)]
        out_img = Image.new("RGBA", img.size, (0, 0, 0, 0))
        out_img.putalpha(lum.point(alpha_lut))

        if out_img.size != (target_size, target_size):
            out_img = out_img.resize((target_size, target_size), Image.Resampling.LANCZOS)

        return out_img

    def generate_fallback_glyph(self, std_code: str) -> Image.Image:
        """
        當處於完全斷網環境時，使用系統標楷體 (kaiu.ttf) 渲染 512x512 高解析度備援字圖。
        """
        target_size = 512
        img = Image.new("RGBA", (target_size, target_size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)

        char_to_draw = None
        for item in COMMON_RARE_CHARS:
            if item["cns"] == std_code:
                char_to_draw = item["char"]
                break

        font_path = "C:/Windows/Fonts/kaiu.ttf"
        if not os.path.exists(font_path):
            font_path = "C:/Windows/Fonts/msjh.ttc"

        try:
            if char_to_draw:
                font = ImageFont.truetype(font_path, int(target_size * 0.84))
                bbox = draw.textbbox((0, 0), char_to_draw, font=font)
                tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
                tx = (target_size - tw) / 2 - bbox[0]
                ty = (target_size - th) / 2 - bbox[1]
                draw.text((tx, ty), char_to_draw, font=font, fill=(0, 0, 0, 255))
            else:
                font = ImageFont.truetype(font_path, int(target_size * 0.20))
                draw.rectangle([28, 28, target_size - 28, target_size - 28], outline=(0, 0, 0, 255), width=10)
                draw.text((55, 195), std_code, font=font, fill=(0, 0, 0, 255))
        except Exception:
            draw.rectangle([30, 30, target_size - 30, target_size - 30], outline=(0, 0, 0, 255), width=8)

        return img

    # ==========================================================================
    # 核心關鍵：寫入 Windows 剪貼簿 (精準控制 20pt 尺寸與垂直置中對齊)
    # ==========================================================================
    def copy_to_word_clipboard(self):
        """
        將目前的難字寫入 Windows 剪貼簿：
          1. Rich Text Format：圖片尺寸以 twip 指定（\\picwgoal），垂直降低以半點指定（\\dn），
             Word 會轉成「字元位置：降低 N 點」，可精準對齊。
          2. CF_DIB：供小畫家等純圖片軟體貼上。
        不放 HTML 格式：實測 Word 會優先採用 HTML，但忽略 CSS 的 vertical-align，
        且把寬度捨去成整數像素（20pt 會變 19.5pt），導致垂直對齊微調完全無效。
        """
        if self.is_fetching:
            self.set_status("字形仍在下載中，請等下載完成再複製。", is_loading=True)
            return
        if not self.current_pil_image or not self.current_cns_code:
            self.set_status("請先查詢或載入一個罕用字後再進行複製。", is_error=True)
            return
        if self.current_is_fallback:
            self.set_status("目前為離線備援圖（非全字庫字形），禁止複製。請連線後重新下載。", is_error=True)
            return
        is_hand = self.current_source == "hand"
        if not is_hand:
            if self.code_var.get().strip() and self.parse_cns_code(self.code_var.get()) \
                    and self.parse_cns_code(self.code_var.get())[2] != self.current_cns_code:
                self.set_status("輸入框代碼與目前預覽字形不同，請先按「載入」。", is_error=True)
                return
            if ("宋體" in self.font_style_var.get()) != self.current_style_is_sung:
                self.set_status("字體選單與目前預覽字形不同，請先按「載入」。", is_error=True)
                return

        try:
            pt_size = float(self.font_size_var.get())
        except ValueError:
            self.set_status("Word 字級請輸入數字（例如 20）。", is_error=True)
            return
        if not 6 <= pt_size <= 200:
            self.set_status("Word 字級需介於 6 ~ 200 pt。", is_error=True)
            return

        lower_pct = self.get_lower_fraction()

        # 將透明背景合成為純白底 512x512 PNG（20pt 時約 1843 DPI，列印邊緣銳利）
        white_bg = Image.new("RGB", self.current_pil_image.size, (255, 255, 255))
        white_bg.paste(self.current_pil_image, mask=self.current_pil_image.split()[3])

        calculated_dpi = int(round(512.0 / (pt_size / 72.0)))
        png_buffer = io.BytesIO()
        white_bg.save(png_buffer, format="PNG", dpi=(calculated_dpi, calculated_dpi))

        # 準備 DIB 格式緩衝區 (BMP 去除前 14 bytes 檔頭)
        bmp_buffer = io.BytesIO()
        white_bg.save(bmp_buffer, format="BMP", dpi=(calculated_dpi, calculated_dpi))
        dib_data = bmp_buffer.getvalue()[14:]

        rtf_bytes = self.build_rtf_payload(png_buffer.getvalue(), white_bg.size, pt_size, lower_pct)

        if HAS_WIN32 and os.name == "nt":
            try:
                # 剪貼簿可能正被其他程式（如 Word）短暫佔用，重試數次
                for attempt in range(20):
                    try:
                        win32clipboard.OpenClipboard()
                        break
                    except Exception:
                        if attempt == 19:
                            raise
                        time.sleep(0.05)
                win32clipboard.EmptyClipboard()

                # Word 採用 RTF：精準尺寸 + 字元位置降低
                cf_rtf = win32clipboard.RegisterClipboardFormat("Rich Text Format")
                win32clipboard.SetClipboardData(cf_rtf, rtf_bytes)

                # 同時註冊 CF_DIB 供其他應用程式貼上
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib_data)

                win32clipboard.CloseClipboard()

                if is_hand:
                    # 手繪字沒有全字庫代碼，不寫入使用紀錄
                    self.set_status(f"已複製手繪字（{int(pt_size)} pt）。到 Word 反白「?」後按 Ctrl+V。")
                else:
                    self.record_to_history(
                        self.current_cns_code,
                        self.font_style_var.get(),
                        f"{int(pt_size)} pt"
                    )
                    self.set_status(
                        f"已複製 [{self.current_cns_code}]（{int(pt_size)} pt）。到 Word 反白「?」後按 Ctrl+V。"
                    )
                self._flash_copied()
                return
            except Exception as e:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
                self.set_status(f"寫入 Windows 剪貼簿時發生錯誤：{e}", is_error=True)
                return
        else:
            self.record_to_history(
                self.current_cns_code,
                self.font_style_var.get(),
                f"{int(pt_size)} pt"
            )
            self.set_status(
                f"已準備 [{self.current_cns_code}] 字圖快取（若需直接 Ctrl+V 貼入 Word，請於 Windows 安裝 pywin32）。"
            )

    def get_lower_fraction(self) -> float:
        """把「垂直對齊微調」轉成降低比例：-27 → 0.27（往下）；+10 → -0.10（往上）"""
        try:
            value = float(self.baseline_offset_var.get().strip().rstrip("%") or 0)
        except ValueError:
            value = 0.0
        value = max(OFFSET_MIN, min(OFFSET_MAX, value))
        return -value / 100.0

    @staticmethod
    def build_rtf_payload(png_bytes: bytes, px_size, pt_size: float, lower_fraction: float) -> bytes:
        """
        產生只含一張行內圖片的 RTF：
          \\fs   圖片所在文字的字級（半點），避免影響行高
          \\dn / \\up   字元位置降低／提高量（半點）
          \\picwgoal/\\pichgoal  圖片顯示尺寸（twip，1pt = 20 twip）
        """
        half_points = int(round(pt_size * 2))
        shift_half_points = int(round(lower_fraction * pt_size * 2))
        position = f"\\dn{shift_half_points}" if shift_half_points >= 0 else f"\\up{-shift_half_points}"
        goal_twips = int(round(pt_size * 20))
        w_px, h_px = px_size
        rtf = (
            "{\\rtf1\\ansi\\deff0"
            f"{{\\fs{half_points}{position}"
            f"{{\\pict\\pngblip\\picw{w_px}\\pich{h_px}"
            f"\\picwgoal{goal_twips}\\pichgoal{goal_twips} "
            f"{png_bytes.hex()}"
            "}}}"
        )
        return rtf.encode("ascii")

    @staticmethod
    def build_cf_html_payload(fragment: str) -> bytes:
        """
        依照微軟官方 CF_HTML (HTML Clipboard Format) 規範封裝 UTF-8 位元組標頭。
        確保 Word 能正確解析 StartHTML / EndHTML 與行內樣式。
        """
        html_template = (
            "<html>\r\n"
            "<head><meta charset=\"utf-8\"></head>\r\n"
            "<body>\r\n"
            "<!--StartFragment-->"
            f"{fragment}"
            "<!--EndFragment-->\r\n"
            "</body>\r\n"
            "</html>"
        )

        header_template = (
            "Version:0.9\r\n"
            "StartHTML:{start_html:010d}\r\n"
            "EndHTML:{end_html:010d}\r\n"
            "StartFragment:{start_frag:010d}\r\n"
            "EndFragment:{end_frag:010d}\r\n"
        )

        dummy_header = header_template.format(start_html=0, end_html=0, start_frag=0, end_frag=0)
        start_html = len(dummy_header.encode("utf-8"))

        full_payload_str = dummy_header + html_template
        full_bytes = full_payload_str.encode("utf-8")

        start_frag_marker = b"<!--StartFragment-->"
        end_frag_marker = b"<!--EndFragment-->"

        start_frag = full_bytes.index(start_frag_marker) + len(start_frag_marker)
        end_frag = full_bytes.index(end_frag_marker)
        end_html = len(full_bytes)

        real_header = header_template.format(
            start_html=start_html,
            end_html=end_html,
            start_frag=start_frag,
            end_frag=end_frag
        )
        return (real_header + html_template).encode("utf-8")

    # ==========================================================================
    # 歷史使用紀錄管理 (自動儲存、點擊回放、單筆/全部清除)
    # ==========================================================================
    def load_history(self) -> list:
        """從本機 history.json 讀取歷史使用紀錄"""
        if not os.path.exists(self.history_file):
            return []
        try:
            if os.path.getsize(self.history_file) > 1024 * 1024:
                return []                     # 異常大的檔案不讀取
            with open(self.history_file, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return []
        if not isinstance(data, list):
            return []
        # 逐筆驗證：檔案可能被手動或他人修改，代碼會用來組成快取檔路徑
        clean = []
        for item in data[:50]:
            if not isinstance(item, dict):
                continue
            code = str(item.get("cns_code", ""))
            if not self._is_std_code(code):
                continue
            style = item.get("font_style")
            clean.append({
                "cns_code": code,
                "font_style": style if style in (STYLE_KAI, STYLE_SUNG) else STYLE_KAI,
                "font_size": str(item.get("font_size", "20 pt"))[:10],
                "last_used": str(item.get("last_used", ""))[:20],
            })
        return clean

    @staticmethod
    def _is_std_code(code: str) -> bool:
        """是否為標準化後的全字庫代碼（字面 1~19 + 四位十六進位大寫）"""
        m = re.fullmatch(r"(\d{1,2})-([0-9A-F]{4})", code or "")
        return bool(m) and 1 <= int(m.group(1)) <= 19

    def save_history(self):
        """將歷史使用紀錄寫入本機 history.json"""
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def record_to_history(self, cns_code: str, font_style: str, font_size: str):
        """新增或更新一筆歷史紀錄，並保持最近使用的排在最上方"""
        now_str = datetime.now().strftime("%Y/%m/%d %H:%M")

        self.history_data = [
            item for item in self.history_data
            if not (item.get("cns_code") == cns_code and item.get("font_style") == font_style)
        ]

        new_entry = {
            "cns_code": cns_code,
            "font_style": font_style,
            "font_size": font_size,
            "last_used": now_str
        }
        self.history_data.insert(0, new_entry)

        self.history_data = self.history_data[:50]
        self.save_history()
        self.refresh_history_tree()

    def _apply_history_style(self, style_str: str) -> bool:
        """把歷史紀錄的字體套回選單；回傳選單是否因此改變"""
        target = "宋體 (全字庫標準明體)" if "宋體" in style_str else "標楷體 (全字庫標準楷書)"
        if self.is_fetching or self.font_style_var.get() == target:
            return False
        self.font_style_var.set(target)
        return True

    def show_about(self):
        """關於與授權：GPL-3.0 第 5 條要求互動介面顯示著作權、無擔保與授權資訊"""
        win = getattr(self, "_about_win", None)
        if win is not None and win.winfo_exists():
            win.lift()
            return
        win = self._about_win = tk.Toplevel(self.root)
        win.title("關於與授權")
        win.configure(bg=C_CARD)
        win.resizable(False, False)
        win.transient(self.root)
        body = tk.Frame(win, bg=C_CARD, padx=self.px(24), pady=self.px(20))
        body.pack(fill=tk.BOTH, expand=True)
        wrap = self.px(440)
        self._label(body, f"全字庫小幫手 v{APP_VERSION}", size=13, weight="bold").pack(anchor="w")
        for text, size, color, gap in (
            (f"Copyright (C) 2026 {APP_AUTHOR}", 10, C_FG, 8),
            ("本程式為自由軟體，免費提供，依 GNU 通用公眾授權條款第 3 版或更新版本（GPL-3.0-or-later）"
             "授權。任何人散布本程式（包括收費散布）都必須附上完整原始碼，並沿用相同授權。", 10, C_FG, 8),
            ("本程式不提供任何擔保，包括可售性或特定用途適用性的默示擔保。詳見授權全文。", 10, C_FG, 8),
            ("字型與字屬性資料：數位發展部，CNS11643中文標準交換碼全字庫網站，https://www.cns11643.gov.tw。"
             "字型採 SIL Open Font License 1.1／政府資料開放授權條款－第1版；資料採政府資料開放授權條款－第1版。",
             9, C_MUTED, 12),
            ("本程式非官方軟體，與數位發展部無關。", 9, C_MUTED, 4),
        ):
            self._label(body, text, size=size, fg=color, anchor="w", justify="left",
                        wraplength=wrap).pack(anchor="w", fill=tk.X, pady=(self.px(gap), 0))

        def open_file(name):
            path = os.path.join(self.base_dir, name)
            if os.path.isfile(path):
                os.startfile(path) if os.name == "nt" else webbrowser.open(f"file://{path}")
            else:
                webbrowser.open(f"{PROJECT_URL}/blob/main/LICENSE")

        btns = tk.Frame(body, bg=C_CARD)
        btns.pack(anchor="w", pady=(self.px(16), 0))
        FlatButton(btns, "原始碼 ↗", variant="outline", small=True,
                   command=lambda: webbrowser.open(PROJECT_URL)).pack(side=tk.LEFT)
        FlatButton(btns, "授權全文", variant="outline", small=True,
                   command=lambda: open_file("LICENSE.txt")).pack(side=tk.LEFT, padx=(self.px(6), 0))
        FlatButton(btns, "第三方元件授權", variant="outline", small=True,
                   command=lambda: open_file("THIRD-PARTY-NOTICES.txt")).pack(side=tk.LEFT, padx=(self.px(6), 0))
        FlatButton(btns, "關閉", small=True, command=win.destroy).pack(side=tk.LEFT, padx=(self.px(6), 0))
        win.bind("<Escape>", lambda e: win.destroy())
        win.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(0, x)}+{max(0, y)}")

    def open_cache_folder(self):
        """開啟本機快取資料夾"""
        if os.name == "nt":
            os.startfile(self.cache_dir)
        else:
            webbrowser.open(f"file://{self.cache_dir}")

    # ==========================================================================
    # 輔助功能：部首與常見人名難字速查子視窗
    # ==========================================================================
if __name__ == "__main__":
    # 啟用 Windows 高 DPI 螢幕清晰顯示支援，避免視窗字體模糊
    if os.name == "nt":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    root = tk.Tk()
    app = CNSRareCharApp(root)
    root.mainloop()