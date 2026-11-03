# 全字庫小幫手（CNS11643 Helper）

公文或申請書上，常有姓名用到電腦打不出、印不出的罕用字。這個工具不用安裝造字檔（EUDC），就能從 [CNS11643 中文標準交換碼全字庫](https://www.cns11643.gov.tw/) 找到字，複製後在 Word 直接貼上，也能正常列印。

Windows 桌面程式，查詢全部在本機離線完成，不需要網路。

<img width="802" height="912" alt="image" src="https://github.com/user-attachments/assets/23930f5a-3202-444f-9bc9-af271b7fe601" />
<img width="802" height="912" alt="image" src="https://github.com/user-attachments/assets/69c16b38-75a8-4944-8607-cf5adc4b553a" />


## 功能

### 找字
九種查詢方式。除了 Unicode，每一種都對應全字庫官網的查詢功能，並且查整個字庫。

| 查詢方式 | 怎麼輸入 |
|---|---|
| CNS 代碼 | `1-4442`、`12-6170`。也可以查一段範圍（`3-4200~3-42FF`），或只打前幾碼（`3-42`） |
| 注音 | `ㄒㄩㄢ` 會列出各聲調。加 `ˉ` 只查一聲。也可以打一個同音字 |
| 拼音 | 漢語、注音二式、耶魯、威妥瑪、通用拼音。可以打 `xuan2` 或 `xuán` |
| 部首／筆畫 | 選部首或總筆畫，也可以再加注音一起篩選，例如火部＋ㄒㄩㄢ |
| 部件 | 打組成的字（`火宣`），或從部件表點選。分「模糊」和「精確」兩種，可以用 `＊` `？` |
| 倉頡 | 英文字母（`FJMM`）或字根（`火十一一`） |
| 筆順 | 橫豎撇點折（1–5），可以用 `＊` `？` |
| Unicode | `U+714A`，或直接貼上字 |
| 複合 | 以上條件一起篩選，跟官網的「複合查詢」相同 |

查不到的時候，按「到官網查詢」會開啟官網對應的查詢頁。

### 複製到 Word
- 在 Word 把「?」反白，按 Ctrl+V，就會變成大小剛好、跟前後文字對齊的字。
- 可以選楷體或宋體、Word 字級（12–36 pt），也可以微調上下位置。
- 也能貼到小畫家等圖片軟體。

### 其他
- **手繪造字**：全字庫也查不到的字，可以直接手寫。程式用 Windows 內建的中文手寫辨識列出候選字，辨識在本機完成，筆跡不會傳出去。
- **使用紀錄**：點一下載入，點兩下直接複製。

### 字形從哪裡來
1. **離線字庫**：用全字庫正楷體字型直接畫出字形，不用連網。
2. **本機快取**：之前下載過的字。
3. **全字庫官網**：宋體，以及離線字庫沒有的字。下載後會存進快取。

## 給同仁：安裝與使用

會收到兩個壓縮檔（`版本` 例如 `v1.0.0`）：

| 檔案 | 大小 | 內容 |
|---|---|---|
| `cns11643-helper-版本-app.zip` | 約 16 MB | 程式本體 |
| `cns11643-helper-版本-cnsdata.zip` | 約 60 MB | 離線字庫（可以不裝） |

### 安裝步驟
1. **核對檔案**：確認收到的檔案沒有被換掉。在命令提示字元執行：
   ```
   certutil -hashfile cns11643-helper-v1.0.0-app.zip SHA256
   ```
   把算出來的結果跟承辦人另外公告的核對碼（或 Release 附的 `SHA256SUMS.txt`）比對，一致才使用。
2. **解除封鎖**：在 zip 上按右鍵 →「內容」→ 勾選最下面的「**解除封鎖**」→「確定」，兩個 zip 都要做。
3. **解壓縮**：在 zip 上按右鍵 →「**全部解壓縮**」。兩個 zip 解壓到**同一個位置**，`cnsdata` 會自動放到程式旁邊。
4. **執行**：打開解壓出來的 `全字庫小幫手` 資料夾，執行 `全字庫小幫手.exe`。

- **放置位置**：請放在自己的資料夾，例如桌面或文件。不要放在大家都能修改的共用資料夾裡直接執行。
- 沒裝字庫也能用，但字形要從官網下載，注音、部首、部件等查詢也不能用。
- 系統需求：Windows 10／11。手繪辨識需要 Windows 內建的「中文（繁體，台灣）」手寫功能。

### 常見問題

**出現「Windows 已保護您的電腦」（Microsoft Defender SmartScreen）**
程式目前沒有數位簽章，Windows 認不出發行者，所以會先擋下來。確認核對碼一致後，可以按「**其他資訊**」→「**仍要執行**」。如果安裝時已經先做了「解除封鎖」，通常不會看到這個警告。

**出現「Failed to load Python DLL」**
- **最常見的原因**：沒有解壓縮，直接在 zip 裡點兩下 exe。請先「全部解壓縮」，再到解壓出來的資料夾裡執行。
- 檢查 `全字庫小幫手\_internal\python314.dll` 還在不在。如果不見了，可能被防毒軟體隔離，請聯絡資訊單位把這支程式加入允許清單。
- Windows 11 開啟「智慧型應用程式控制」時，可能會擋下沒有數位簽章的程式，也請聯絡資訊單位協助。
- 只支援 Windows 10／11。

**右上角顯示「線上模式（離線字庫已停用）」**
字庫檔案跟程式版本不符，或檔案被修改過。請重新下載同一個版本的 cnsdata.zip，解壓並覆蓋原本的資料夾。

## 開發

### 環境
- Windows
- Python 3（目前用 3.14 開發）

```
pip install requests certifi pillow pywin32 pyinstaller
python 全字庫小幫手.py
```

### 產生離線字庫
`cnsdata/` 大約 114 MB，沒有放進 git，要從官網開放資料重新產生：

```
python tools/build_cnsdata.py
```

會下載官網的 `Fonts_Kai.zip`、`MapingTables.zip`、`Properties.zip`，產生字型、對照表和查詢資料，同時更新 `cnsdata_manifest.py`。

### 打包

```
pyinstaller --noconfirm 全字庫小幫手.spec
```

打包時會自動在 exe 旁邊放上 `LICENSE.txt`（GPL-3.0 全文）和 `THIRD-PARTY-NOTICES.txt`（第三方元件授權全文），由 `tools/collect_licenses.py` 產生，**散布時一定要附上**。

打包完成後：
1. 把 `cnsdata/` 複製到 `dist/全字庫小幫手/` 裡。
2. 分別壓縮程式（exe、`_internal`、`LICENSE.txt`、`THIRD-PARTY-NOTICES.txt`）與 `cnsdata`，命名為 `cns11643-helper-版本-app.zip` 和 `cns11643-helper-版本-cnsdata.zip`。GitHub Release 的附件檔名不能用中文。
3. 用 `certutil -hashfile` 算出兩個 zip 的 SHA-256，再透過公文或內部網站等**另一個管道**公告。

> ⚠️ **字庫核對**：`cnsdata_manifest.py` 記錄每個字庫檔的 SHA-256，會打包進 exe。程式每次啟動都會比對，字庫被修改或版本不符，就停用離線字庫，改用線上模式。
> 所以**更新字庫後一定要重新打包程式**，兩個 zip 也要一起發。只想依現有的 `cnsdata` 重算核對碼：
> ```
> python tools/build_cnsdata.py --manifest-only
> ```

### 專案結構

```
全字庫小幫手.py         主程式（Tkinter）
全字庫小幫手.spec       PyInstaller 打包設定
cnsdata_manifest.py     字庫核對碼（自動產生，請勿手動修改）
LICENSE                 本專案授權（GPL-3.0 全文）
tools/build_cnsdata.py  從官網產生 cnsdata/
tools/collect_licenses.py  產生第三方元件授權聲明（打包時自動執行）
cnsdata/                離線字庫（不放進 git）
```

### 資料存放
程式資料夾可以寫入時，快取（`cache/`）和使用紀錄（`history.json`）放在程式旁邊。不能寫入時（例如裝在 Program Files），改放在 `%LOCALAPPDATA%\全字庫小幫手\`。

## 安全
- 只連全字庫官網，只走 HTTPS（TLS 1.2 以上），完整驗證憑證，不跟隨轉址，並限制下載大小。
- 全字庫字型只給這支程式使用，不會安裝到電腦上。
- 啟動時比對離線字庫的核對碼（見上方「字庫核對」）。
- 快取圖片和使用紀錄在讀取前會先檢查。

## 授權

### 本專案
Copyright (C) 2026 71658-dev

本專案是**免費**的自由軟體，以 [GNU 通用公眾授權條款第 3 版或更新版本（GPL-3.0-or-later）](LICENSE) 授權：

- ✅ 可以免費使用、複製、修改、再散布，個人、機關、學校、企業都可以用。
- ✅ 修改後再散布，必須**一樣以 GPL 授權並公開完整原始碼**，不能改成封閉軟體。
- ⚠️ GPL 不禁止收費散布。但任何人散布本程式（包括收費散布），都必須同時提供完整原始碼，並讓收到的人享有同樣的權利，所以不能把本程式變成私有商品販售。
- ❌ 本程式**不提供任何擔保**。使用前請自行確認字形正確，正式文件以全字庫官網為準。

本專案是個人開發的**非官方**工具，與數位發展部及全字庫網站無關。

### 字型與資料
離線字庫（`cnsdata/`）和線上取得的字形，都來自數位發展部「CNS11643 中文標準交換碼全字庫」網站的開放資料，各自依原授權條款授權，**不適用**本專案的 GPL：

| 內容 | 來源檔案 | 授權 | 本專案的處理 |
|---|---|---|---|
| 全字庫正楷體字型（BMP、Ext-B、Plus）<br>`cnsdata/fonts/*.ttf` | `Fonts_Kai.zip` | (c) 數位發展部。字型內載明可擇一適用 [SIL Open Font License 1.1](https://openfontlicense.org) 或 [政府資料開放授權條款－第 1 版](https://data.gov.tw/license) | 原檔未修改，依 OFL-1.1 散布，附授權全文 `cnsdata/LICENSE-OFL-1.1.txt` |
| CNS ↔ Unicode 對照表<br>`cnsdata/cns2unicode.tsv` | `MapingTables.zip` | [政府資料開放授權條款－第 1 版](https://data.gov.tw/license) | 轉換為 tsv 格式 |
| 注音、部首、筆畫、倉頡、筆順、部件、拼音<br>`cnsdata/cns_*.tsv` | `Properties.zip` | [政府資料開放授權條款－第 1 版](https://data.gov.tw/license) | 轉換為 tsv 格式 |
| 線上字形圖片（宋體、離線字庫沒有的字） | 全字庫網站字形資訊頁 | 依全字庫網站授權說明 | 只存在使用者自己電腦的快取，不隨本專案散布 |

**出處標示**（依[全字庫網站授權說明](https://www.cns11643.gov.tw/pageView.jsp?ID=59)）：

> 數位發展部，CNS11643中文標準交換碼全字庫網站，https://www.cns11643.gov.tw。

全字庫網站以無償、非專屬方式授權，可以不限時間和地點重製、編輯、公開傳輸，或開發各種產品和服務，但不包含商標權和專利權。散布字型時，必須附上著作權聲明和 OFL-1.1 全文。本專案都已附上。

### 第三方元件
打包後的程式內含下列開放原始碼元件，授權全文都附在 `THIRD-PARTY-NOTICES.txt`。這些授權都可以跟 GPL-3.0 一起使用。

| 元件 | 用途 | 授權 |
|---|---|---|
| [Python](https://www.python.org/)（含 OpenSSL、libffi、zlib、Tcl/Tk 等） | 執行環境 | PSF License 等 |
| [requests](https://github.com/psf/requests) | 連線全字庫官網 | Apache-2.0 |
| [urllib3](https://github.com/urllib3/urllib3) | 連線元件 | MIT |
| [idna](https://github.com/kjd/idna) | 網域名稱處理 | BSD-3-Clause |
| [charset-normalizer](https://github.com/jawah/charset_normalizer) | 文字編碼偵測 | MIT |
| [certifi](https://github.com/certifi/python-certifi) | HTTPS 根憑證 | MPL-2.0 |
| [Pillow](https://github.com/python-pillow/Pillow)（含 FreeType、libjpeg、libpng 等） | 字形圖片處理 | MIT-CMU 等 |
| [pywin32](https://github.com/mhammond/pywin32) | 剪貼簿、手寫辨識 | PSF |
| [PyInstaller](https://github.com/pyinstaller/pyinstaller) | 打包成 exe | GPL-2.0-or-later（附例外條款，允許打包任何授權的程式） |

手寫辨識使用 Windows 內建的「中文（繁體，台灣）」手寫辨識功能，不隨本程式散布。

## 參考資料
- CNS11643 中文標準交換碼全字庫：https://www.cns11643.gov.tw
- 全字庫授權說明（政府網站資料開放宣告）：https://www.cns11643.gov.tw/pageView.jsp?ID=59
- 全字庫開放資料下載（政府資料開放平臺）：https://data.gov.tw/dataset/5961
- 全字庫查詢功能（本程式的查詢方式依此設計，結果以官網為準）：
  - 複合查詢：https://www.cns11643.gov.tw/search.jsp?ID=99
  - 部首查詢：https://www.cns11643.gov.tw/searchR.jsp
  - 注音、拼音、部件、倉頡、筆順、CNS 碼、Unicode 查詢：`https://www.cns11643.gov.tw/search.jsp?ID=` 加上 3、5、7、4、8、6、12
- 政府資料開放授權條款－第 1 版：https://data.gov.tw/license
- SIL Open Font License 1.1：https://openfontlicense.org
- GNU General Public License v3.0：https://www.gnu.org/licenses/gpl-3.0.html
