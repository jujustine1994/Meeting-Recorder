# ARCHITECTURE — Meeting Recorder

## 工具總覽

錄製 Windows 系統音訊（WASAPI Loopback），儲存為 MP3。
tkinter GUI 視窗，支援開始 / 暫停 / 繼續 / 停止，可選擇停止後儲存或捨棄，支援連續錄多段、自訂檔名。

## 檔案清單

| 檔案 | 用途 |
|------|------|
| `meeting_recorder啟動器.bat` | 使用者雙擊的入口，2 行，呼叫 launcher.ps1 |
| `launcher.ps1` | 環境檢查（Python / uv / venv）+ 啟動 main.py |
| `main.py` | 主程式：錄音控制、MP3 轉換、儲存 |
| `i18n.py` | 介面文字多語言查表（`t()` / `set_lang()`），四語清單在 `LANGUAGES` |
| `config.py` | `config.json` 讀寫；目前只存 `language` 一個欄位 |
| `logtext.py` | `logs/app.log` 的訊息字串，**固定繁中**，不進 i18n |
| `locales/zh_tw.py` | 繁體中文（母表，其他語言的翻譯來源），108 條 |
| `locales/zh_cn.py` | 简体中文 |
| `locales/en.py` | English |
| `locales/ja.py` | 日本語 |
| `config.json` | 使用者的語言設定（執行時產生，`.gitignore` 排除） |
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
                    ├── _pick_language_on_first_run()（config.json 沒語言才跳）
                    └── tkinter 主視窗（建 widget 前先 i18n.set_lang）
                            ├── 選擇儲存位置 / 錄音模式 / 檔名
                            ├── 按「開始錄音」→ loopback + mic 背景執行緒
                            ├── VU Meter 即時更新、計時器每秒更新
                            ├── 按「暫停」→ stop_stream() 喚醒 thread → pa.terminate() → 保留資料
                            ├── 按「繼續錄音」→ 新 PyAudio + 新執行緒，資料接續累加
                            ├── 按「停止並儲存」→ force_stop → join → pa.terminate() → lameenc → 存 MP3
                            ├── 按「停止不儲存」→ 確認彈窗 → force_stop → join → 丟棄資料
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

## 架構決策紀錄

### `_save_after_stop` 保持單一方法（2026-05-13）

`_save_after_stop` 約 140 行，mode × output_mode 組合產生 5–6 條 code path，全部寫在一個方法裡。

**為何不重構：**
- 功能正常，沒有 bug
- 重構（抽 helper 或改 dict dispatch）本身有引入錯誤的風險，且測試無法覆蓋所有組合
- 只有在「真的要加新輸出格式（WAV、FLAC 等）」時才值得動

**若未來要擴充：**
先將 `else: # both` 的 ~80 行抽成 `_encode_and_save_both()`，讓 `_save_after_stop` 只負責 setup + dispatch。
抽出後再新增格式，不要在現有巢狀結構裡直接硬塞。

### 多語言採「重開生效」而非即時切換（2026-08-15）

即時切換要建 widget 登記表、切語言時逐一 `config(text=...)`，漏掉任一元件就是
中英混雜，三個 popup 還要各自處理。改動幅度大好幾倍，換來的只是省一次重開。

**哪些字串不翻（資料，不是介面文字）：**

| 字串 | 為什麼 |
|---|---|
| 錄音模式代號 `system` / `mic` / `both` | Radiobutton 的 value、`_save_mode` 的值，程式自己 `==` 比對 |
| 輸出模式 `merge` / `separate` / `both` | 同上 |
| `msg_queue` 的訊息類型（`saved` / `error` / `vu_mic`...） | 執行緒間的協定 |
| 存檔後綴 `_system` / `_mic` | 落進檔名，翻了使用者的舊檔就對不起來 |
| `_make_progress_cb("system"/"mic")` 的標籤 | 刻意與上面的檔名後綴一致，讓 log 的「system 編碼完成」對得上 `_system.mp3` |
| `logs/app.log` 的全部內容 | 跟著使用者語言變，維護者看不懂自己的 log |

**唯一的例外是 `系統預設`（`gui.dev.default`）**：它同時是裝置下拉選單的顯示
文字，也被 `sys_worker()` 拿去 `!= "系統預設"` 比對。判定為可翻，因為
(1) 語言在單次執行內固定不變，顯示端與比對端用同一個 `t()` 呼叫；
(2) 存進 `selected_output_name` 的是 `None` 或真實裝置名，不是這個顯示字串，
所以不落檔、不跨執行比對。**日後若要把裝置選擇存進 config.json，這條必須先
改成「顯示值 / 儲存值分離」再存**，否則就是資料污染。

---

## 關鍵設定變數（main.py）

| 變數 | 位置 | 說明 |
|------|------|------|
| `_DEFAULT_BIT_RATE` | 模組頂部 | MP3 位元率預設值，可在進階設定中更改 |
| `_MIX_WEIGHT_SYSTEM / _MIX_WEIGHT_MIC` | 模組頂部 | 混音權重，進階使用者直接改此常數 |
| `_SILENCE_RMS_THRESHOLD` | 模組頂部 | 靜音閾值，影響警告橫幅與等化計算 |
| `_SILENCE_WARNING_SECS` | 模組頂部 | 靜音警告倒計時 |
| `chunk` | `_record_worker()` | 每次讀取的音訊幀數，預設 512 |
| `_record_stream` | `__init__` | 活躍的 loopback stream 參照，供 `_force_stop_streams()` 解除 read() 阻塞 |
| `_mic_stream` | `__init__` | 活躍的 mic stream 參照，同上 |
| `_elapsed_before_pause` | `__init__` | 暫停前已累計秒數，resume 後計時器從此繼續 |
