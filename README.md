# LTC 重疊檢查工具

居家長照服務時間重疊自動檢查系統，支援個案、居服員、支援與上課四種衝突類型檢查。

---

## 下載執行檔

右側 **Actions** → 最新成功的 workflow run → **LTC_Checker_Windows** artifact 下載。

解壓縮後直接執行：

| 檔案 | 說明 |
|---|---|
| `LTC_Checker.exe` | 命令提示字元版，雙擊或直接執行 |
| `LTC_Checker_UI.exe` | 圖形介面版，雙擊執行 |

---

## 資料夾結構

exe 執行時，會在**同一層目錄**尋找以下資料夾：

```
LTC_Checker.exe
LTC_Checker_UI.exe
csv/                        ← 各公司服務紀錄 xlsx（檔名不限）
support/
    CS/                     ← 公司子資料夾（名稱自訂）
    CZ/
    YH/
    ...（每個子資料夾放排班表 + 單頁服務紀錄，檔名不限）
class/
    CS/
    CZ/
    YH/
    ...（每個子資料夾放上課排班表 + 單頁服務紀錄 + 名單 .txt）
save/                       ← 結果輸出（自動建立）
```

> **檔名不需要改**：工具依檔案內容自動辨識類型，不依賴命名規則。

---

## 檢查項目

| 類型 | 說明 |
|---|---|
| 個案重疊 | 同一個案在同一時段有兩筆服務紀錄 |
| 居服員重疊 | 同一居服員在同一時段服務不同個案，或跨公司連續服務 |
| 支援範圍 | 居服員服務時間超出支援排班範圍 |
| 上課衝突 | 居服員上課期間仍有服務紀錄 |

---

## 結果輸出

```
save/
    Patient/
        <個案>_patient_overlap.txt
    Worker/
        <居服員>_worker_overlap_<原因>_<公司>.txt
    support/
        <公司>/output_support.txt
    class/
        <公司>/output_class.txt
```

每筆衝突紀錄包含服務時間、個案、居服員、公司、地址，以及**來源檔案名稱與 Excel 列號**。

---

## 本機開發

```bash
pip install pandas numpy openpyxl rich customtkinter

# CLI 版
python main.py

# UI 版
python app.py
```

### 打包 Windows exe（需在 Windows 上執行）

```bat
pip install pyinstaller
pyinstaller ltc_checker.spec --clean       # CLI 版
pyinstaller ltc_checker_ui.spec --clean    # UI 版
```

或直接 push 到 main branch，GitHub Actions 會自動建置。
