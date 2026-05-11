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
