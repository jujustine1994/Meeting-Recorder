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

**錄音控制**
- 三種錄音模式：電腦聲音（WASAPI Loopback）/ 麥克風 / 兩者同時
- 暫停 / 繼續錄音（保留已錄資料，計時器累計）
- 停止不儲存（含確認彈窗防誤觸）
- 可連續錄多段，不需重新啟動

**輸出**
- 儲存為 MP3，位元率可選 128 / 192 / 320 kbps
- 「系統＋麥克風」模式下可選：合併一軌 / 獨立兩軌（`_system.mp3` + `_mic.mp3`）/ 兩個都要
- 自訂檔名，自動加時間戳記（`meeting_YYYY-MM-DD_HH-MM-SS.mp3`）
- 同名自動流水號，不覆蓋舊檔

**音質**
- 自動等化音量（混音前拉齊兩軌響度，可設增益上限與靜音過濾）
- 麥克風使用 MME API 錄音，繞過 Discord / Teams / Zoom 的 WASAPI 音訊增強干擾

**UI / 監控**
- VU Meter 即時顯示系統音訊與麥克風音量
- 連續靜音 10 秒自動顯示警告橫幅
- 存檔時 log 即時顯示 MP3 編碼進度
- 音訊裝置中斷自動重試（最多 30 秒），失敗才通知使用者
- 裝置設定與測試：選擇輸出 / 輸入裝置，測試麥克風 VU meter
- 自選儲存資料夾（預設 Desktop）

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
