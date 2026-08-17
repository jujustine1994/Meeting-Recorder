# PITFALLS — Meeting Recorder

遇到問題再累積，格式：問題／原因／解法／禁止。

---

## Pitfall 1：WASAPI `stream.read()` 無限阻塞 → 閃退，無任何 traceback

**問題**  
按下「停止並儲存」後約 3～5 秒程式直接消失，無 Python 錯誤訊息、無 Windows 錯誤對話框。

**原因**  
`stream.read(512)` 是 blocking call，正常約 12ms 返回。但在以下情況 WASAPI 裝置停止派發 chunk，導致 read() **無限掛住**：
- 插拔耳機 / 切換播放裝置
- Teams、Zoom 等會議軟體搶佔音訊工作階段
- Windows 11 電源管理（螢幕關閉、省電模式）

原本流程：停止 → `is_recording = False` → `join(timeout=3)` 等 3 秒 → 超時，thread 仍阻塞在 read() → 直接呼叫 `pa.terminate()` → PortAudio C 層釋放資源，但 thread 仍持有 stream handle → **segfault，閃退**。

**解法**  
- 新增 `self._record_stream` / `self._mic_stream` 儲存活躍 stream 參照
- 在 join 之前呼叫 `_force_stop_streams()`：對活躍 stream 呼叫 `stop_stream()`，強制喚醒阻塞的 read()，讓 thread 收到 OSError 後自行退出
- 加 `is_alive()` 防線：join 後確認 thread 已結束才呼叫 `pa.terminate()`；若 thread 仍存活則 log warning 且**跳過** terminate

**禁止**  
- 不可在 `join(timeout)` 超時後直接呼叫 `pa.terminate()`  
- 不可移除 `_force_stop_streams()` 呼叫（即使「正常情況」執行時用不到）  
- 不可把 join timeout 設得過短（目前 5s，已考慮裝置切換的 0.5s + 1s sleep）

---

## Pitfall 2：Discord（及 Teams/Zoom）透過 WASAPI 啟用音訊增強 → 麥克風錄音破音

**問題**  
與 Discord 同時錄音時，`_mic.mp3` 聲音失真（AGC pumping、類機器人聲），但通話本身聲音正常，且用 Audacity 單獨測試麥克風也正常。

**原因**  
Discord 在 WASAPI 層對麥克風裝置啟用 Windows 音訊增強（AGC 自動增益、降噪、回音消除）。這些增強套用於整個 WASAPI shared mode session，所有透過 WASAPI 存取同一麥克風的應用程式（包括本程式）都會收到已處理的失真訊號，而非原始硬體音訊。

**解法**  
麥克風改用 **MME host API** 開啟（`_find_mme_mic_device()`），MME 直接存取硬體，完全繞過 WASAPI 層的音訊增強。`_record_mic_worker` 已改為呼叫此 helper 取得 MME 裝置 index。

**禁止**  
- 不可把麥克風 stream 改回用 `self.selected_input_idx`（WASAPI index）直接開啟  
- `_find_mme_mic_device()` 的 fallback 鏈（找不到對應 MME 裝置 → MME 預設 → 原 WASAPI index）必須保留，否則找不到 MME 時會 crash

---

## Pitfall 3：區域變數 `t` 遮蔽翻譯函式 → `_stop_recording` 直接 UnboundLocalError

**問題**
導入多語言後，按「停止並儲存」直接拋 `UnboundLocalError`，錄音檔存不出來。

**原因**
`from i18n import t` 是模組層級的匯入，但 `_stop_recording` 裡有一個叫 `t` 的區域指派。
Python 只要在函式內有 `t = ...`，整個函式的 `t` 就被當成區域變數，
**指派之前的每一次 `t(...)` 呼叫都會 UnboundLocalError**——即使模組層級明明有 `t`。

**解法**
把遮蔽的名字全部改掉，並加 `test_nothing_shadows_the_translation_function`
（用 AST 掃函式定義、迴圈變數、指派、**函式參數**四種形狀）永久釘住。

**禁止**
- 不可用 `t` 當任何區域變數、迴圈變數、函式參數的名字
- 不可只用 `grep "\bt\b"` 檢查，雜訊太多會漏；一定要用 AST
- 導入新的翻譯字串前，先跑那條測試

**備註**
這個坑在 10 個專案的多語言遷移裡有 6 個中招，Snap GIF Creator 一個就 15 處。
跨專案的完整說明見 `C:\Users\CTH\Documents\Code\_i18n_migration\i18n_lessons.md`。
