# Save Progress Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 log 視窗中即時覆寫顯示 MP3 編碼進度百分比，讓使用者知道長時間錄音儲存時程式正常運作。

**Architecture:** `_encode_to_mp3` 改為 20 段 chunked 編碼並透過 callback 回報進度；`_poll_queue` 處理新的 `progress`/`progress_done` 訊息類型，以覆寫方式更新 log 最後一行；`_save_after_stop` 為每次 encode 呼叫提供對應 callback。

**Tech Stack:** Python tkinter（Text widget 覆寫）、lameenc（支援分批 encode）、queue.Queue（現有 msg_queue）

---

## 檔案異動

- Modify: `main.py` — 所有改動都在這一個檔案

---

## Task 1：基礎設施 — 旗標、log 覆寫、callback 工廠

**Files:**
- Modify: `main.py:115`（`__init__` 加旗標）
- Modify: `main.py:1399`（`_poll_queue` 加 progress 處理）
- Modify: `main.py:1389`（`_log` 附近加兩個 helper method）

- [ ] **Step 1：在 `__init__` 加 `_progress_line_active` 旗標**

在 `main.py` 第 115 行 `self.msg_queue` 那行後面加：

```python
self._progress_line_active: bool = False  # log 最後一行是否為可覆寫的進度行
```

- [ ] **Step 2：加 `_log_progress` 與 `_make_progress_cb` 兩個 helper method**

在 `_log` method（約 1389 行）後面加入：

```python
def _log_progress(self, text: str, done: bool = False):
    """覆寫 log 最後一行（進度更新），done=True 時清除旗標讓該行永久保留。"""
    self.log_text.config(state="normal")
    if self._progress_line_active:
        self.log_text.delete("end-1c linestart", "end-1c")
        self.log_text.insert("end-1c", text)
    else:
        self.log_text.insert("end", text + "\n")
    self.log_text.see("end")
    self.log_text.config(state="disabled")
    self._progress_line_active = not done

def _make_progress_cb(self, label: str, file_idx: int, total: int):
    """回傳一個 callback，每次呼叫時把進度推進 msg_queue。"""
    def cb(pct: int):
        self.msg_queue.put(("progress",
            f"⏳ 編碼 {label} ({file_idx}/{total}) {pct}%"))
    return cb
```

- [ ] **Step 3：在 `_poll_queue` 加 `progress` 與 `progress_done` 處理**

在 `_poll_queue`（約 1399 行）的 `while True:` 迴圈最頂端（`if msg_type == "saved":` 之前）加：

```python
if msg_type == "progress":
    self._log_progress(data, done=False)
    continue

if msg_type == "progress_done":
    self._log_progress(data, done=True)
    continue
```

同時，在 `while True:` 迴圈第一行（`msg_type, data = self.msg_queue.get_nowait()` 之後，第一個 `if` 之前）加一行清旗標：

```python
if msg_type not in ("progress", "progress_done"):
    self._progress_line_active = False
```

> **注意：** 這行確保任何非進度訊息（如 warning、saved）都會讓下一次進度從新行開始。

- [ ] **Step 4：手動驗證旗標邏輯**

用紙筆或腦中模擬以下序列，確認行為：

```
1. put("progress", "⏳ sys 10%")   → flag=False → append "⏳ sys 10%\n", flag=True
2. put("progress", "⏳ sys 50%")   → flag=True  → 覆寫為 "⏳ sys 50%", flag=True
3. put("progress_done", "✓ sys")   → flag=True  → 覆寫為 "✓ sys", flag=False
4. put("warning", "something")     → flag=False → 不改旗標，正常 append
5. put("progress", "⏳ mic 10%")   → flag=False → 新行 "⏳ mic 10%\n", flag=True
```

確認步驟 3 後 "✓ sys" 會永久保留，步驟 5 不會覆寫 "✓ sys"。

- [ ] **Step 5：Commit**

```bash
git add main.py
git commit -m "feat: progress log 基礎設施（覆寫旗標、_log_progress、_make_progress_cb）"
```

---

## Task 2：`_encode_to_mp3` 改為 chunked 編碼

**Files:**
- Modify: `main.py:1222-1229`

- [ ] **Step 1：替換 `_encode_to_mp3` 實作**

將現有 `_encode_to_mp3`（1222–1229 行）整個替換：

```python
def _encode_to_mp3(self, pcm_data: bytes, channels: int, sample_rate: int,
                   bit_rate: int = _DEFAULT_BIT_RATE,
                   progress_cb=None) -> bytes:
    encoder = lameenc.Encoder()
    encoder.set_bit_rate(bit_rate)
    encoder.set_in_sample_rate(sample_rate)
    encoder.set_channels(channels)
    encoder.set_quality(2)

    if progress_cb is None or len(pcm_data) == 0:
        return encoder.encode(pcm_data) + encoder.flush()

    # 切 20 段，對齊 frame 邊界（channels × 2 bytes）
    frame_size = channels * 2
    total = len(pcm_data)
    chunk_size = max((total // 20 // frame_size) * frame_size, frame_size)

    result = b""
    offset = 0
    while offset < total:
        chunk = pcm_data[offset:offset + chunk_size]
        result += encoder.encode(chunk)
        offset += len(chunk)
        progress_cb(min(100, int(offset / total * 100)))

    return result + encoder.flush()
```

- [ ] **Step 2：確認無 callback 時行為不變**

舊 call site（還沒改的地方）不傳 `progress_cb`，走 `if progress_cb is None` 分支，行為與改動前完全一樣。

- [ ] **Step 3：Commit**

```bash
git add main.py
git commit -m "feat: _encode_to_mp3 支援 chunked 編碼與 progress callback"
```

---

## Task 3：`_save_after_stop` 接上 callback

**Files:**
- Modify: `main.py:1290-1377`

- [ ] **Step 1：system / mic 模式加 callback（1374–1375 行）**

將：
```python
mp3_data = self._encode_to_mp3(pcm_data, channels, sample_rate,
                               bit_rate=self._save_bit_rate)
filepath = self._save_file(mp3_data, base_name)
self.msg_queue.put(("saved", [filepath]))
```

改為：
```python
mp3_data = self._encode_to_mp3(pcm_data, channels, sample_rate,
                               bit_rate=self._save_bit_rate,
                               progress_cb=self._make_progress_cb("音訊", 1, 1))
self.msg_queue.put(("progress_done", "✓ 編碼完成"))
filepath = self._save_file(mp3_data, base_name)
self.msg_queue.put(("saved", [filepath]))
```

- [ ] **Step 2：both 模式 fallback（mic 無資料）加 callback（1314–1319 行）**

將：
```python
mp3_data = self._encode_to_mp3(
    b"".join(self.record_frames),
    self.record_channels, self.record_sample_rate,
    bit_rate=self._save_bit_rate)
filepath = self._save_file(mp3_data, base_name)
self.msg_queue.put(("saved", [filepath]))
```

改為：
```python
mp3_data = self._encode_to_mp3(
    b"".join(self.record_frames),
    self.record_channels, self.record_sample_rate,
    bit_rate=self._save_bit_rate,
    progress_cb=self._make_progress_cb("音訊", 1, 1))
self.msg_queue.put(("progress_done", "✓ 編碼完成"))
filepath = self._save_file(mp3_data, base_name)
self.msg_queue.put(("saved", [filepath]))
```

- [ ] **Step 3：both 模式正常路徑 — 計算 encode 總數並加 callback（1322 行後）**

在 `output_mode = self._save_output_mode` 後加：

```python
total_enc = (2 if output_mode in ("separate", "both") else 0) + \
            (1 if output_mode in ("merge", "both") else 0)
enc_idx = 0
```

- [ ] **Step 4：both 模式獨立音軌區塊加 callback（1328–1338 行）**

將：
```python
if output_mode in ("separate", "both"):
    sys_mp3 = self._encode_to_mp3(
        b"".join(sys_frames_snap),
        self.record_channels, self.record_sample_rate,
        bit_rate=self._save_bit_rate)
    mic_mp3 = self._encode_to_mp3(
        b"".join(mic_frames_snap),
        self.record_mic_channels, self.record_mic_rate,
        bit_rate=self._save_bit_rate)
    saved_paths.append(self._save_file(sys_mp3, base_name, "_system"))
    saved_paths.append(self._save_file(mic_mp3, base_name, "_mic"))
```

改為：
```python
if output_mode in ("separate", "both"):
    enc_idx += 1
    sys_mp3 = self._encode_to_mp3(
        b"".join(sys_frames_snap),
        self.record_channels, self.record_sample_rate,
        bit_rate=self._save_bit_rate,
        progress_cb=self._make_progress_cb("system", enc_idx, total_enc))
    self.msg_queue.put(("progress_done", "✓ system 編碼完成"))
    enc_idx += 1
    mic_mp3 = self._encode_to_mp3(
        b"".join(mic_frames_snap),
        self.record_mic_channels, self.record_mic_rate,
        bit_rate=self._save_bit_rate,
        progress_cb=self._make_progress_cb("mic", enc_idx, total_enc))
    self.msg_queue.put(("progress_done", "✓ mic 編碼完成"))
    saved_paths.append(self._save_file(sys_mp3, base_name, "_system"))
    saved_paths.append(self._save_file(mic_mp3, base_name, "_mic"))
```

- [ ] **Step 5：both 模式合併音軌區塊加 callback（1363–1365 行）**

將：
```python
merged_mp3 = self._encode_to_mp3(
    mixed_pcm, self.record_channels, self.record_sample_rate,
    bit_rate=self._save_bit_rate)
saved_paths.append(self._save_file(merged_mp3, base_name))
```

改為：
```python
enc_idx += 1
merged_mp3 = self._encode_to_mp3(
    mixed_pcm, self.record_channels, self.record_sample_rate,
    bit_rate=self._save_bit_rate,
    progress_cb=self._make_progress_cb("合併音訊", enc_idx, total_enc))
self.msg_queue.put(("progress_done", "✓ 合併音訊編碼完成"))
saved_paths.append(self._save_file(merged_mp3, base_name))
```

- [ ] **Step 6：手動測試**

分三種情境各錄 10 秒測試：

1. **system 模式** → 停止 → log 應出現 `⏳ 編碼 音訊 (1/1) X%` 覆寫更新，最後 `✓ 編碼完成`
2. **both / merge 模式** → 停止 → log 應出現混音中... + `⏳ 編碼 合併音訊 (1/1) X%`
3. **both / both 模式** → 停止 → log 應依序出現 system → mic → 合併音訊三段進度

確認：
- 進度行只佔 log 一行（不堆疊）
- 百分比從低到高更新
- 最後顯示 `✓ meeting_XXXX.mp3`（現有 saved 訊息）

- [ ] **Step 7：Commit**

```bash
git add main.py
git commit -m "feat: 存檔時在 log 顯示 MP3 編碼進度（both/both 最多 3 段）"
```
