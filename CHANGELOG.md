# CHANGELOG — Meeting Recorder

## 現狀總覽（2026-05-13）

### 已完成功能
- [x] WASAPI Loopback 系統音訊錄製
- [x] lameenc MP3 編碼（無需 ffmpeg）
- [x] tkinter 資料夾選擇視窗
- [x] 錄音計時器顯示
- [x] 時間戳記自動命名
- [x] 連續錄音（不重啟）
- [x] 自動安裝啟動器（launcher.ps1）
- [x] 三種錄音模式（電腦聲音 / 麥克風 / 兩者混音）
- [x] 靜音警告橫幅（連續 10 秒 RMS < 100 時顯示）
- [x] 同名自動流水號避免覆蓋
- [x] 裝置設定與測試（選擇 WASAPI 裝置、麥克風測試）
- [x] VU Meter 即時音量顯示（線性換算，正常音量約 20~50%；峰值接近滿刻度時短暫變紅警示爆音）
- [x] 獨立音軌輸出（合併一軌 / 獨立兩軌 / 兩個都要）
- [x] 自動等化音量（混音前拉齊兩軌響度，含增益上限與靜音過濾）
- [x] 進階設定彈窗
- [x] 停止不儲存（含確認彈窗防誤觸）
- [x] 暫停 / 繼續錄音（保留已錄資料，計時器累計）
- [x] 主畫面內嵌輸出格式（系統＋麥克風模式動態顯示）
- [x] MP3 位元率可設定（128 / 192 / 320 kbps）
- [x] 模組級 config 常數（混音權重、靜音閾值等）
- [x] 麥克風以 MME 開啟（繞過 Discord/Teams WASAPI 音訊增強，裝置測試同步）
- [x] 存檔時 log 即時顯示 MP3 編碼進度（both/both 最多 3 段）
- [x] 音訊裝置中斷自動重試（30 秒）並通知使用者（系統音訊 + 麥克風均適用）
- [x] 系統音訊裝置不可恢復時自動儲存已錄部分（不遺失資料）
- [x] 麥克風斷線即時警示標籤（VU meter 旁紅字，與「靜音」區分）
- [x] 合併音訊混音補零（mic 中途斷線不截短 merged MP3）

### 未完成 / 待規劃
- [ ] 支援選擇錄音裝置（多音效卡環境）
- [ ] 錄音完成後自動開啟資料夾

---

## 更新記錄

### 2026-07-23
- 調整：VU meter 除數 650→120——實測回報 650 過鬆導致大聲僅約 10%，回推調回 120，目標讓大聲落在 50~60%（見 TODO 待驗證）

### 2026-07-17 — 導入執行紀錄（log）規範
- 新增：`logs/app.log`——launcher.ps1 與 main.py 共用同一檔案，靠標籤區分來源，事後可回查任務失敗原因
- 新增：`launcher.ps1` 加入 `Write-Log` / `Write-LogHeader`，環境檢查失敗（Python / uv / 套件安裝）與主程式異常結束（exit code）落檔；`trap` 攔截到的閃退也落檔
- 新增：`main.py` 加入 `_find_project_root()` / `_write_log()` / `_write_log_header()`；`_log()` 改為支援 `to_file` 參數（預設 `False`，fail-closed）
- 落檔範圍：任務起始（開始錄音，含模式／輸出方式／位元率）、錯誤（例外類型 + 重試次數，不落例外訊息全文避免夾帶敏感內容）、任務結果（成功／失敗／使用者捨棄 + 耗時）；`_log_progress()` 的即時進度訊息維持只推 UI 不落檔
- 新增：`.gitignore` 加入 `logs/`

### 2026-07-03
- 調整：VU meter 顯示公式除數改小，正常說話音量從原本的 1~10% 提升到約 20~50%，更容易一眼判斷是否有在收音
- 新增：音訊峰值超過滿刻度 90% 時，對應指示條與百分比文字短暫變紅警示爆音（保持 1.5 秒）
- 套用範圍：正式錄音畫面與「裝置設定與測試」畫面的麥克風／系統音訊測試指示條，兩邊行為一致

### 2026-06-10
- 修正：`winget install Python` 加入 `--override "/quiet PrependPath=1 Include_pip=1"`，確保靜默安裝後 Python 自動加進 PATH
- 修正：`launcher.ps1` 加入全域 `trap`，攔截未處理例外，防止執行失敗時視窗直接閃退

### 2026-05-28 — Bug 修正 + 文件清理

- 修復：`both` 模式錄音中途麥克風斷線 30 秒無法重連時，錯誤訊息等級誤判為 `error`（觸發 UI 強制重置），應為 `warning`（保留系統音訊繼續錄）
  - 根因：`_save_mode` 初始值為 `"system"`，`_stop_recording` 呼叫前背景執行緒讀到的是舊值
  - 解法：`_start_recording` 開始時即設定 `_save_mode = mode`
- 清理：`ARCHITECTURE.md` 關鍵設定變數表移除指向舊函式內位置的四個條目（已重構為模組頂部常數）

### 2026-05-13 — 穩定性強化（裝置斷線、混音截短、麥克風警示）

- 修復：系統音訊裝置不可恢復（30 秒重試失敗）時，改為自動觸發儲存流程，保留已錄資料；原本 UI 直接重置，資料無法存取
- 修復：`_mix_pcm` 改用較長音軌為基準補零，避免 mic 中途斷線時 merged MP3 被截短
- 新增：麥克風 OSError 改為最多重試 30 秒（與系統音訊一致）；斷線期間 VU meter 旁顯示紅字「⚠ 麥克風斷線」，重連或停止後自動清除

---

### 2026-05-13 — 存檔進度顯示、裝置中斷處理、MME 完善

- 新增：存檔時在 log 即時顯示 MP3 編碼進度百分比（覆寫同一行，不堆疊）
  - `both/both` 模式最多 3 段：`⏳ 編碼 system (1/3) 45%` → `✓ system 編碼完成` → ...
  - `_encode_to_mp3` 改為 20 段 chunked 編碼，輸出與原本 bit-perfect 一致
- 修復：系統音訊 OSError（裝置中斷）改為最多重試 30 秒，失敗才通知使用者；VU 立即歸 0 告知狀況
- 修復：麥克風 OSError 原本靜默結束，現在送 warning 並將 VU 歸 0
- 改善：`_find_mme_mic_device` 加 `wasapi_idx` 參數，名稱比對失敗時送 warning 告知實際錄音裝置可能與所選不符
- 改善：裝置測試對話框的麥克風測試也改用 MME，確保測試結果與實際錄音一致

### 2026-05-13 — 麥克風改用 MME 繞過 Discord 音訊增強干擾

- 修復：與 Discord（及 Teams、Zoom）同時錄音時，麥克風音訊失真（AGC pumping、類機器人聲）
  - 根因：Discord 在 WASAPI 層對麥克風裝置啟用 Windows 音訊增強（AGC、降噪、回音消除），影響所有 WASAPI shared mode 客戶端
  - 解法：新增 `_find_mme_mic_device()` 將麥克風切換為 MME host API 開啟，繞過 WASAPI 層增強；含 WASAPI→MME 名稱比對、三層 fallback
- 新增：PITFALLS.md Pitfall 2 — Discord 音訊增強干擾麥克風錄音

### 2026-05-11 — 進階設定重構、主畫面輸出格式、config 常數

- 新增：主畫面「系統＋麥克風」模式下動態顯示輸出格式選項（合併 / 獨立 / 兩個都要）
- 新增：MP3 位元率可在進階設定選擇（128 / 192 / 320 kbps）
- 改善：進階設定重組為「音質」+「混音設定」兩區，移除輸出方式區塊
- 改善：等化子選項（增益上限、靜音過濾）在未開啟自動等化時 gray out
- 重構：硬編碼參數提取為模組頂部 config 常數（方便進階使用者直接修改）

### 2026-05-11 — 暫停錄音、停止不儲存、閃退修復

- 新增：暫停 / 繼續錄音 — 錄音中可暫停（保留資料），之後繼續錄音，計時器累計兩段時間
- 新增：停止不儲存 — 錄音中可丟棄本次錄音（彈確認框防誤觸），記錄顯示「✗ 錄音已捨棄（未儲存）」
- 修復：WASAPI `stream.read()` 無限阻塞導致閃退（無 traceback）
  - 根因：`join(timeout)` 超時後直接 `pa.terminate()`，Thread 仍持有 stream handle，C 層 crash
  - 解法：新增 `_force_stop_streams()` 在 join 前先 `stop_stream()`；加 `is_alive()` 防線避免 unsafe terminate

### 2026-05-05 — VU Meter、獨立音軌輸出、自動等化
- 新增：VU Meter — 主視窗顯示系統音訊與麥克風即時音量條，錄音中動態更新，停止後歸零
- 新增：進階設定彈窗（⚙ 按鈕）— 含輸出方式、自動等化、增益上限、靜音過濾設定
- 新增：獨立音軌輸出 — 「系統+麥克風」模式可選合併一軌 / 獨立兩軌（`_system.mp3` + `_mic.mp3`）/ 兩個都要
- 新增：自動等化音量 — 混音前自動拉齊兩軌響度，增益上限預設 4x，可選排除靜音段計算 RMS
- 重構：抽取 `_encode_to_mp3`、`_save_file` 輔助方法，`saved` 訊息格式改為 list 支援多檔回報
- 新增：`tests/test_audio_processing.py` — 13 個單元測試覆蓋等化與增益計算函式

### 2026-03-24
- 新增：專案初始建立，完整錄音功能上線
- 修改：主程式改為 tkinter GUI 視窗介面（參考 SnapTranscript 架構）
- 新增：自訂檔案名稱欄位，錄音前可先填好名稱
- 修正：launcher.ps1 加入 ARM64 架構偵測，強制安裝 x64 Python 確保 pyaudiowpatch 相容性
- 新增：三種錄音模式（電腦聲音 / 麥克風 / 兩者混音）
- 新增：Mode "both" 混音邏輯（loopback stereo + mic mono → upmix → 0.6 權重混音）
- 修正：靜音偵測依模式切換偵測對象（loopback 或麥克風）

### 2026-03-23 — 程式碼全面審查與 Bug 修正
- 修正：Mode "both" 麥克風無資料時改存純系統音訊並顯示警告，避免產生無聲 MP3
- 修正：Mode "both" 麥克風取樣率與 loopback 不一致時顯示警告（混音仍繼續）
- 修正：loopback 裝置持續不可用時加入 sleep(1) 避免 CPU busy-wait
- 修正：tkinter StringVar race condition — record_mode 改在主執行緒停止時鎖定為 _save_mode
- 修正：Mode "both" 麥克風啟動失敗改為 warning 而非 error，不中止儲存流程
- 改善：新增 warning 訊息類型至 _poll_queue，顯示在 log 但不中斷流程
- 改善：全程式補齊關鍵邏輯的中文註解（混音權重、靜音閾值、stream closure 設計等）
- 改善：_compute_rms() 提取為獨立函式，消除兩個 worker 的重複靜音偵測程式碼
