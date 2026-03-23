# ARCHITECTURE — Meeting Recorder

## 工具總覽

錄製 Windows 系統音訊（WASAPI Loopback），儲存為 MP3。
tkinter GUI 視窗，按鈕開始 / 結束，支援連續錄多段、自訂檔名。

## 檔案清單

| 檔案 | 用途 |
|------|------|
| `meeting_recorder啟動器.bat` | 使用者雙擊的入口，2 行，呼叫 launcher.ps1 |
| `launcher.ps1` | 環境檢查（Python / uv / venv）+ 啟動 main.py |
| `main.py` | 主程式：錄音控制、MP3 轉換、儲存 |
| `requirements.txt` | Python 套件清單 |
| `.gitignore` | 版本控制排除清單 |
| `README.md` | 專案說明 |
| `ARCHITECTURE.md` | 本檔案 |
| `CHANGELOG.md` | 更新紀錄 |
| `TODO.md` | 待辦清單 |
| `PITFALLS.md` | 已知地雷 |

## 執行流程

```
使用者雙擊 .bat
    └── launcher.ps1
            ├── [1/3] 檢查 Python（沒有就用 winget 安裝）
            ├── [2/3] 檢查 uv
            ├── [3/3] 檢查 venv（沒有就建立並安裝套件）
            └── python main.py
                    ├── cls + CTH Banner
                    ├── tkinter 選擇儲存資料夾
                    └── 錄音迴圈
                            ├── Enter → 開始錄音（背景執行緒）
                            ├── 計時器顯示（背景執行緒）
                            ├── Enter → 停止錄音
                            ├── lameenc 轉 MP3
                            ├── 儲存 meeting_YYYY-MM-DD_HH-MM-SS.mp3
                            └── 回到等待下一段
```

## 音訊技術細節

- **錄音方式**：WASAPI Loopback（捕捉系統輸出，非麥克風）
- **採樣格式**：PCM Int16
- **採樣率**：跟隨系統預設輸出裝置（通常 44100 或 48000 Hz）
- **聲道數**：最多 2ch（MP3 限制）
- **MP3 位元率**：128 kbps

## 音效卡相容性（靜音錄音行為）

WASAPI Loopback 的「靜音是否影響錄音」取決於音效卡驅動的截取點位置：

| 音效卡類型 | 靜音能錄音 | 說明 |
|-----------|----------|------|
| **Realtek HDA**（大多數 PC/筆電） | ✅ 可以 | Pre-volume tap，市佔最高 |
| **Intel HDA** | ✅ 可以 | 行為同 Realtek |
| **Creative Sound Blaster** | ❓ 未知 | 少數，需測試 |
| **USB 外接音效卡** | ❓ 未知 | 部分廠商 post-volume |
| **虛擬機音效卡** | ❌ 不支援 | 無真實 WASAPI 支援 |

開發者電腦音效卡：**Realtek High Definition Audio + Intel Smart Sound Technology**
驗證結果：靜音狀態下 WASAPI Loopback 仍可正常錄音（pre-volume tap）。

程式已內建靜音偵測：連續 10 秒 RMS < 100 時顯示橘色警告橫幅。

## 關鍵設定變數（main.py）

| 變數 | 位置 | 說明 |
|------|------|------|
| `SILENCE_RMS_THRESHOLD` | `_record_worker()` | 靜音判斷閾值，預設 100（Int16 最大 32767） |
| `SILENCE_WARNING_SECS` | `_record_worker()` | 靜音幾秒後顯示警告，預設 10 秒 |
| `chunk` | `_record_worker()` | 每次讀取的音訊幀數，預設 512 |
| `bit_rate` | `_save_after_stop()` | MP3 位元率，預設 128 kbps |
| `quality` | `_save_after_stop()` | lameenc 編碼品質，2=高品質 |
