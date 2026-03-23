```
/*  ================================  *\
 *                                    *
 *          C  T  H                   *
 *        created by CTH              *
 *                                    *
\*  ================================  */
```

規則檔: windows-tool.md
類型: Windows 工具

# Meeting Recorder

按 Enter 開始/結束，自動錄製電腦系統音訊並儲存為 MP3。主要用於線上會議錄音。

## 功能

- 錄製 Windows 系統輸出音訊（WASAPI Loopback）
- 儲存為 MP3 格式（128kbps）
- 自選儲存資料夾（預設 Desktop）
- 檔名自動加時間戳記（`meeting_YYYY-MM-DD_HH-MM-SS.mp3`）
- 可連續錄多段，不需重新啟動

## 系統需求

- Windows 10 / 11
- Python 3.8+（首次執行自動安裝）
- 音效卡支援 WASAPI（一般 Windows 電腦皆支援）

## 執行方式

雙擊 `meeting_recorder啟動器.bat`

首次執行會自動安裝所需套件，之後直接進入錄音介面。

## 技術棧

- Python 3
- `pyaudiowpatch` — WASAPI Loopback 系統音訊捕捉
- `lameenc` — MP3 編碼（純 Python，不需 ffmpeg）
- `tkinter` — 資料夾選擇視窗（Python 內建）

## .gitignore 規則

- `venv/`
- `__pycache__/`
- `*.pyc`
- `*.log`
- `.env`
