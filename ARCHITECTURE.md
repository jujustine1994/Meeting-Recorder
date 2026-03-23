# ARCHITECTURE — Meeting Recorder

## 工具總覽

錄製 Windows 系統音訊（WASAPI Loopback），儲存為 MP3。
終端機介面，Enter 開始 / Enter 結束，支援連續錄多段。

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

## 關鍵設定變數（main.py）

| 變數 | 位置 | 說明 |
|------|------|------|
| `CHUNK` | `record_session()` | 每次讀取的音訊幀數，預設 512 |
| `bit_rate` | `record_session()` | MP3 位元率，預設 128 kbps |
| `quality` | `record_session()` | lameenc 編碼品質，2=高品質 |
