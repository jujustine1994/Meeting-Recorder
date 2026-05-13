# Spec：存檔進度顯示

**日期：** 2026-05-13  
**背景：** 錄音超過一小時後，停止並儲存時 MP3 編碼耗時可能超過 30 秒，期間 UI 無任何進度反饋，使用者不確定程式是否正常運作。

---

## 目標

在現有 log 視窗中即時顯示 MP3 編碼進度（百分比），覆寫同一行，不堆疊。

---

## 不在範圍內

- join PCM frames 的進度（`b"".join()`，記憶體操作，快到不需要顯示）
- 混音步驟的進度（`_mix_pcm`，純 Python 運算，通常 < 1 秒）
- 寫檔進度（`_save_file`，I/O 操作，通常 < 1 秒）

---

## 改動清單

### 1. `_encode_to_mp3` 加 `progress_cb` 參數

```python
def _encode_to_mp3(self, pcm_data, channels, sample_rate,
                   bit_rate=_DEFAULT_BIT_RATE, progress_cb=None):
```

- 將 `pcm_data` 切成 20 段（每段對齊 frame 邊界：`channels × 2` bytes）
- 每段 encode 完後呼叫 `progress_cb(pct: int)`，pct 範圍 1–100
- 無 `progress_cb` 時行為與現在完全相同（向後相容）

### 2. App 新增 `_progress_line_active: bool` 旗標

初始為 `False`。

### 3. `msg_queue` 新增 `progress` 訊息類型

`_poll_queue` 收到 `("progress", text)` 時：

- 若 `_progress_line_active == True`：刪除 log 最後一行，插入新文字
- 若 `_progress_line_active == False`：正常 append，然後設旗標為 `True`

收到其他任何訊息類型時：清除 `_progress_line_active = False`，然後正常處理。

### 4. `_save_after_stop` 提供 callback

依模式計算 encode 總次數，為每次 encode 產生對應 callback：

| 模式 | encode 次數 | log 顯示 |
|------|-------------|---------|
| system | 1 | `⏳ 編碼音訊 (1/1) X%` |
| mic | 1 | `⏳ 編碼音訊 (1/1) X%` |
| both / separate | 2 | `⏳ 編碼 system (1/2) X%`、`⏳ 編碼 mic (2/2) X%` |
| both / merge | 1 | `⏳ 編碼合併音訊 (1/1) X%` |
| both / both | 3 | `⏳ 編碼 system (1/3) X%`、`⏳ 編碼 mic (2/3) X%`、`⏳ 編碼合併音訊 (3/3) X%` |

每次 encode 完成後，送 `("warning", "✓ XXX 完成")` 讓那行永久保留（旗標清除）。

---

## Log 視覺效果（both/both 模式）

```
儲存中...
⏳ 編碼 system (1/3) 45%      ← 同一行覆寫更新
⏳ 編碼 system (1/3) 100%
✓ system 編碼完成             ← 普通 append
⏳ 編碼 mic (2/3) 30%         ← 覆寫更新
...
✓ mic 編碼完成
混音中...                      ← status label（現有行為）
⏳ 編碼合併音訊 (3/3) 88%
✓ 合併音訊編碼完成
✓ meeting_2026-05-13.mp3      ← 現有 saved 訊息
```

---

## 額外負擔評估

- 20 次 `encoder.encode()` vs 1 次：Python call overhead 可忽略，C 層計算量完全相同
- MP3 輸出與現在 bit-perfect 一致（lameenc 設計支援分批餵入）
- `msg_queue` 多 20 條 progress 訊息：可忽略
