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

1. 解壓 app.zip，執行 `全字庫小幫手\全字庫小幫手.exe`。
2. 要用離線查詢的話，把字庫 zip 解壓到**同一個位置**，`cnsdata` 資料夾會自動放到程式旁邊。
3. 沒裝字庫也能用，但字形要從官網下載，注音、部首、部件等查詢也不能用。

**收到檔案後，請先確認沒有被換掉。** 在命令提示字元執行：

```
certutil -hashfile cns11643-helper-v1.0.0-app.zip SHA256
```

把算出來的結果跟承辦人另外公告的核對碼比對，一致才使用。

**放置位置**：請放在自己的資料夾，例如桌面或文件。不要放在大家都能修改的共用資料夾裡直接執行。

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

打包完成後：
1. 把 `cnsdata/` 複製到 `dist/全字庫小幫手/` 裡。
2. 分別壓縮程式（exe 和 `_internal`）與 `cnsdata`，命名為 `cns11643-helper-版本-app.zip` 和 `cns11643-helper-版本-cnsdata.zip`。GitHub Release 的附件檔名不能用中文。
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
tools/build_cnsdata.py  從官網產生 cnsdata/
cnsdata/                離線字庫（不放進 git）
```

### 資料存放
程式資料夾可以寫入時，快取（`cache/`）和使用紀錄（`history.json`）放在程式旁邊。不能寫入時（例如裝在 Program Files），改放在 `%LOCALAPPDATA%\全字庫小幫手\`。

## 安全
- 只連全字庫官網，只走 HTTPS（TLS 1.2 以上），完整驗證憑證，不跟隨轉址，並限制下載大小。
- 全字庫字型只給這支程式使用，不會安裝到電腦上。
- 啟動時比對離線字庫的核對碼（見上方「字庫核對」）。
- 快取圖片和使用紀錄在讀取前會先檢查。

## 資料來源與授權
資料來自數位發展部「CNS11643 中文標準交換碼全字庫」網站（https://www.cns11643.gov.tw）：

- 全字庫正楷體字型：[SIL Open Font License 1.1](https://openfontlicense.org)
- 中文碼對照表，以及注音、部首、筆畫、倉頡、筆順、部件、拼音等字屬性：[政府資料開放授權條款－第 1 版](https://data.gov.tw/license)

各檔來源與授權全文，請看產生後的 `cnsdata/來源與授權.txt` 和 `cnsdata/LICENSE-OFL-1.1.txt`。
