# VU Meter、獨立音軌輸出、自動等化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在主視窗新增 VU Meter 即時音量顯示，支援「系統+麥克風」模式輸出獨立音軌，並加入自動等化音量功能。

**Architecture:** 全部修改集中在 `main.py`；錄音執行緒透過現有 `msg_queue` 送出 VU 資料，主執行緒在 `_poll_queue` 更新 UI。等化邏輯為 `@staticmethod` 純函式，可獨立測試。獨立音軌輸出在 `_save_after_stop` 的 "both" 分支實作。

**Tech Stack:** Python 3 / tkinter / pyaudiowpatch / lameenc / unittest（內建，無需新增相依）

---

## 檔案修改清單

| 檔案 | 變更類型 | 說明 |
|------|---------|------|
| `main.py` | 修改 | 所有功能變更 |
| `tests/test_audio_processing.py` | 新增 | 等化函式單元測試 |

---

## Task 1：新增實例變數

**Files:**
- Modify: `main.py`（`__init__` 方法，約 line 96–118）

- [ ] **Step 1：在 `__init__` 的「儲存設定」區塊之後加入以下變數**

找到這段（約 line 114）：
```python
        self.filename_var    = tk.StringVar()
```

在它之後插入：
```python
        # VU Meter
        self.vu_system_var = tk.DoubleVar(value=0)
        self.vu_mic_var    = tk.DoubleVar(value=0)

        # 進階設定
        self.output_mode       = tk.StringVar(value="merge")   # merge / separate / both
        self.equalize_enabled  = tk.BooleanVar(value=False)
        self.eq_gain_cap       = tk.IntVar(value=4)            # 倍數上限 1～16
        self.eq_filter_silence = tk.BooleanVar(value=True)
```

同時在「錄音狀態」區（約 line 108）補上停止時鎖定用的快照變數：
```python
        self._save_mode: str = "system"
        self._save_output_mode: str = "merge"      # ← 新增
        self._save_equalize: bool = False           # ← 新增
        self._save_gain_cap: float = 4.0            # ← 新增
        self._save_filter_silence: bool = True      # ← 新增
```

- [ ] **Step 2：Commit**
```
git add main.py
git commit -m "feat: 新增 VU Meter 與進階設定實例變數"
```

---

## Task 2：修改 "saved" 訊息格式為 list，支援多檔回報

`_save_after_stop` 之後會需要回報多個已儲存檔案。現在先把格式改好，確保向後相容。

**Files:**
- Modify: `main.py`（`_save_after_stop` 約 line 903、`_poll_queue` 約 line 938）

- [ ] **Step 1：修改 `_save_after_stop` 最後的 saved 訊息**

找到（約 line 901–903）：
```python
            with open(filepath, "wb") as f:
                f.write(mp3_data)

            self.msg_queue.put(("saved", filepath))
```

改為：
```python
            with open(filepath, "wb") as f:
                f.write(mp3_data)

            self.msg_queue.put(("saved", [filepath]))
```

- [ ] **Step 2：修改 `_poll_queue` 的 "saved" 處理**

找到（約 line 938–942）：
```python
                if msg_type == "saved":
                    filename = os.path.basename(data)
                    self._log(f"✓  {filename}")
                    self._reset_ui_after_stop()
                    self.status_label.config(text=f"已儲存：{filename}", foreground="green")
```

改為：
```python
                if msg_type == "saved":
                    for fp in data:
                        self._log(f"✓  {os.path.basename(fp)}")
                    self.status_label.config(
                        text=f"已儲存：{os.path.basename(data[-1])}", foreground="green")
                    self._reset_ui_after_stop()
```

- [ ] **Step 3：手動驗證（錄一段短音訊，確認仍正常儲存並顯示檔名）**

```
# 在終端機啟動：
venv\Scripts\python.exe main.py
```
操作：選系統模式 → 開始錄音 → 約 3 秒後停止 → 確認 log 顯示 ✓ 檔名，status 更新正確。

- [ ] **Step 4：Commit**
```
git add main.py
git commit -m "refactor: saved 訊息格式改為 list，支援多檔回報"
```

---

## Task 3：抽取 `_encode_to_mp3` 和 `_save_file` 輔助方法

**Files:**
- Modify: `main.py`（`_save_after_stop` 內的 MP3 編碼與存檔邏輯）

- [ ] **Step 1：在 `_save_after_stop` 之前（約 line 818）加入兩個新方法**

在 `def _save_after_stop(self):` 前面插入：

```python
    def _encode_to_mp3(self, pcm_data: bytes, channels: int, sample_rate: int) -> bytes:
        encoder = lameenc.Encoder()
        encoder.set_bit_rate(128)
        encoder.set_in_sample_rate(sample_rate)
        encoder.set_channels(channels)
        encoder.set_quality(2)
        return encoder.encode(pcm_data) + encoder.flush()

    def _save_file(self, mp3_data: bytes, base_name: str, suffix: str = "") -> str:
        name = base_name + suffix
        filepath = os.path.join(self.save_folder, f"{name}.mp3")
        counter = 2
        while os.path.exists(filepath):
            filepath = os.path.join(self.save_folder, f"{name} ({counter}).mp3")
            counter += 1
        with open(filepath, "wb") as f:
            f.write(mp3_data)
        return filepath
```

- [ ] **Step 2：將 `base_name` 計算移到 if/elif/else 之前，並改寫結尾的編碼存檔段落**

Task 8 的 "both" 分支會提前 `return`，必須在 if 之前就能取得 `base_name`，否則 undefined。

找到 `_save_after_stop` 的起始位置，在 `mode = self._save_mode` 之後、`if mode == "system":` 之前加入：

```python
            custom_name = self.filename_var.get().strip()
            base_name   = custom_name if custom_name else (
                "meeting_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            )
```

接著找到結尾的整段舊 encoder 程式碼（約 line 880–903）：
```python
            encoder = lameenc.Encoder()
            encoder.set_bit_rate(128)
            encoder.set_in_sample_rate(sample_rate)
            encoder.set_channels(channels)
            encoder.set_quality(2)  # 2=高品質，7=快速低品質

            mp3_data = encoder.encode(pcm_data) + encoder.flush()

            custom_name = self.filename_var.get().strip()
            base_name   = custom_name if custom_name else (
                "meeting_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            )

            # 同名已存在時自動加流水號，避免覆蓋
            filepath = os.path.join(self.save_folder, f"{base_name}.mp3")
            counter = 2
            while os.path.exists(filepath):
                filepath = os.path.join(self.save_folder, f"{base_name} ({counter}).mp3")
                counter += 1

            with open(filepath, "wb") as f:
                f.write(mp3_data)

            self.msg_queue.put(("saved", [filepath]))
```

改為（`base_name` 已移走，不需再定義）：
```python
            mp3_data = self._encode_to_mp3(pcm_data, channels, sample_rate)
            filepath = self._save_file(mp3_data, base_name)
            self.msg_queue.put(("saved", [filepath]))
```

- [ ] **Step 3：手動驗證（同 Task 2 Step 3）**

- [ ] **Step 4：Commit**
```
git add main.py
git commit -m "refactor: 抽取 _encode_to_mp3 和 _save_file 輔助方法"
```

---

## Task 4：新增 VU Meter UI 元素

**Files:**
- Modify: `main.py`（`_build_ui`、`_poll_queue`、`_reset_ui_after_stop`）

- [ ] **Step 1：修改 `_build_ui` 的按鈕區**

找到（約 line 204–207）：
```python
        ttk.Button(
            frame_btn, text="🔧 裝置設定與測試",
            command=self._show_device_test,
        ).pack(pady=(12, 0))
```

改為（VU Meter + 雙按鈕列）：
```python
        # VU Meter
        vu_frame = tk.Frame(frame_btn)
        vu_frame.pack(pady=(12, 0))

        tk.Label(vu_frame, text="系統音訊", width=8, anchor="w",
                 font=("", 9)).grid(row=0, column=0, padx=(0, 6))
        ttk.Progressbar(vu_frame, variable=self.vu_system_var,
                        maximum=100, length=200).grid(row=0, column=1)
        self.vu_system_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                        foreground="gray", font=("Consolas", 9))
        self.vu_system_pct.grid(row=0, column=2, padx=(4, 0))

        tk.Label(vu_frame, text="麥克風", width=8, anchor="w",
                 font=("", 9)).grid(row=1, column=0, padx=(0, 6), pady=(4, 0))
        ttk.Progressbar(vu_frame, variable=self.vu_mic_var,
                        maximum=100, length=200).grid(row=1, column=1, pady=(4, 0))
        self.vu_mic_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                     foreground="gray", font=("Consolas", 9))
        self.vu_mic_pct.grid(row=1, column=2, padx=(4, 0), pady=(4, 0))

        # 裝置設定 + 進階設定 雙按鈕列
        btn_row = tk.Frame(frame_btn)
        btn_row.pack(pady=(12, 0))
        ttk.Button(btn_row, text="🔧 裝置設定與測試",
                   command=self._show_device_test).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="⚙ 進階設定",
                   command=self._show_advanced_settings).pack(side="left")
```

- [ ] **Step 2：在 `_poll_queue` 的 `elif msg_type == "status":` 之前加入 VU 處理**

找到（約 line 954）：
```python
                elif msg_type == "status":
                    self.status_label.config(text=data, foreground="gray")
```

在它前面插入：
```python
                elif msg_type == "vu_system":
                    self.vu_system_var.set(data)
                    self.vu_system_pct.config(text=f"{int(data):3d}%")

                elif msg_type == "vu_mic":
                    self.vu_mic_var.set(data)
                    self.vu_mic_pct.config(text=f"{int(data):3d}%")
```

- [ ] **Step 3：在 `_reset_ui_after_stop` 結尾加入 VU 重置**

找到：
```python
    def _reset_ui_after_stop(self):
        self.btn_record.config(state="normal", text="⏺  開始錄音")
        self.timer_label.config(text="00:00", foreground="gray")
        self.silence_banner.grid_remove()
        self._set_mode_radios_state("normal")
```

在最後加：
```python
        self.vu_system_var.set(0)
        self.vu_mic_var.set(0)
        self.vu_system_pct.config(text="  0%")
        self.vu_mic_pct.config(text="  0%")
```

- [ ] **Step 4：新增空的 `_show_advanced_settings` stub（先讓按鈕不報錯）**

在 `_show_device_test` 方法之後加入：
```python
    def _show_advanced_settings(self):
        pass  # 由 Task 8 實作
```

- [ ] **Step 5：啟動 app，確認 VU Meter 兩條 Progressbar 出現，兩個按鈕並排顯示**

```
venv\Scripts\python.exe main.py
```

- [ ] **Step 6：Commit**
```
git add main.py
git commit -m "feat: 新增 VU Meter UI 與進階設定按鈕（stub）"
```

---

## Task 5：錄音執行緒送出 VU 訊息

**Files:**
- Modify: `main.py`（`_record_worker`、`_record_mic_worker`）

- [ ] **Step 1：在 `_record_worker` 的靜音偵測之前加入 vu_system 訊息**

找到（約 line 651）：
```python
                    # ---- 靜音偵測 ----
                    rms = _compute_rms(data)
```

在 `rms = _compute_rms(data)` 之後加一行：
```python
                    self.msg_queue.put(("vu_system", min(100.0, rms / 327.67)))
```

- [ ] **Step 2：改寫 `_record_mic_worker` 使 RMS 計算不受 check_silence 控制**

找到（約 line 744–761）：
```python
            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.mic_frames.append(data)

                    if check_silence:
                        rms = _compute_rms(data)
                        if rms < SILENCE_RMS_THRESHOLD:
```

改為（提前計算 rms，VU 與靜音偵測共用）：
```python
            while self.is_recording:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    self.mic_frames.append(data)

                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_mic", min(100.0, rms / 327.67)))

                    if check_silence:
                        if rms < SILENCE_RMS_THRESHOLD:
```

注意：原本 `if check_silence:` 內的第一行 `rms = _compute_rms(data)` 要刪掉（已移到外面）。

- [ ] **Step 3：啟動 app，錄音時確認兩條 Progressbar 隨音量動態變化，停止後歸零**

- [ ] **Step 4：Commit**
```
git add main.py
git commit -m "feat: 錄音執行緒送出 VU Meter 訊息"
```

---

## Task 6：等化函式（TDD）

**Files:**
- Create: `tests/test_audio_processing.py`
- Modify: `main.py`（新增 `_compute_equalize_gain`、`_apply_gain_to_pcm`）

- [ ] **Step 1：建立 `tests/` 目錄**
```
mkdir tests
```

- [ ] **Step 2：寫測試檔 `tests/test_audio_processing.py`**

```python
import sys
import os
import struct
import array
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from main import MeetingRecorderApp, _compute_rms


def make_pcm(value: int, n_samples: int = 512) -> bytes:
    """產生 n_samples 個相同振幅的 PCM Int16 bytes"""
    return struct.pack(f"{n_samples}h", *([value] * n_samples))


class TestComputeEqualizeGain(unittest.TestCase):

    def test_mic_quieter_gets_boosted(self):
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(250)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertAlmostEqual(mic_gain, 4.0, places=1)

    def test_sys_quieter_gets_boosted(self):
        sys_frames = [make_pcm(250)] * 10
        mic_frames = [make_pcm(1000)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertAlmostEqual(sys_gain, 4.0, places=1)
        self.assertEqual(mic_gain, 1.0)

    def test_gain_capped_at_gain_cap(self):
        sys_frames = [make_pcm(10000)] * 10
        mic_frames = [make_pcm(100)] * 10   # 比值 = 100，超過 cap
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False, gain_cap=4.0)
        self.assertEqual(mic_gain, 4.0)

    def test_both_silent_returns_no_gain(self):
        sys_frames = [make_pcm(0)] * 10
        mic_frames = [make_pcm(0)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_one_track_silent_returns_no_gain(self):
        sys_frames = [make_pcm(0)] * 10
        mic_frames = [make_pcm(1000)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_equal_rms_returns_no_gain(self):
        frames = [make_pcm(500)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            frames, frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_filter_silence_excludes_quiet_chunks(self):
        # sys: 全部 loud（RMS≈1000）
        # mic: 一半 silent（RMS≈10, 低於閾值100）、一半 loud（RMS≈1000）
        # filter_silence=True 後，mic active_rms ≈ 1000 → 與 sys 等響 → gain ≈ 1.0
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(10)] * 5 + [make_pcm(1000)] * 5
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=True, gain_cap=10.0)
        self.assertAlmostEqual(mic_gain, 1.0, places=1)

    def test_filter_silence_false_includes_quiet_chunks(self):
        # 不過濾靜音時，平均 RMS 被靜音段壓低，gain > 1
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(10)] * 5 + [make_pcm(1000)] * 5
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False, gain_cap=10.0)
        self.assertGreater(mic_gain, 1.5)


class TestApplyGainToPcm(unittest.TestCase):

    def test_gain_1_unchanged(self):
        data = make_pcm(1000)
        self.assertEqual(MeetingRecorderApp._apply_gain_to_pcm(data, 1.0), data)

    def test_gain_doubles_amplitude(self):
        data = make_pcm(1000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertEqual(list(arr), [2000, 2000, 2000, 2000])

    def test_positive_clamp_prevents_overflow(self):
        data = make_pcm(20000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertTrue(all(s == 32767 for s in arr))

    def test_negative_clamp_prevents_overflow(self):
        data = struct.pack("4h", -20000, -20000, -20000, -20000)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertTrue(all(s == -32768 for s in arr))

    def test_zero_gain_returns_silence(self):
        data = make_pcm(1000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 0.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertEqual(list(arr), [0, 0, 0, 0])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3：執行測試，確認全部 FAIL（函式尚未實作）**
```
venv\Scripts\python.exe -m unittest tests.test_audio_processing -v
```
預期：`AttributeError: type object 'MeetingRecorderApp' has no attribute '_compute_equalize_gain'`

- [ ] **Step 4：在 `main.py` 的 `_mix_pcm` 方法之後加入兩個 staticmethod**

找到 `_mix_pcm` 結尾（約 line 816），在它之後插入：

```python
    @staticmethod
    def _compute_equalize_gain(
        sys_frames: list,
        mic_frames: list,
        filter_silence: bool = True,
        gain_cap: float = 4.0,
    ) -> tuple:
        """
        回傳 (sys_gain, mic_gain)。只有較小的那軌 gain > 1.0。
        filter_silence=True 時排除 RMS < 100 的 chunk，避免靜音段影響平均。
        任一軌全靜音時回傳 (1.0, 1.0) 不等化。
        """
        THRESHOLD = 100

        def active_rms(frames):
            values = [_compute_rms(c) for c in frames]
            if filter_silence:
                values = [v for v in values if v >= THRESHOLD]
            return (sum(values) / len(values)) if values else 0.0

        sys_rms = active_rms(sys_frames)
        mic_rms = active_rms(mic_frames)

        if sys_rms == 0 or mic_rms == 0:
            return 1.0, 1.0

        if sys_rms >= mic_rms:
            return 1.0, min(sys_rms / mic_rms, gain_cap)
        else:
            return min(mic_rms / sys_rms, gain_cap), 1.0

    @staticmethod
    def _apply_gain_to_pcm(data: bytes, gain: float) -> bytes:
        """對 Int16 PCM 套用增益倍數，結果 clamp 至 [-32768, 32767]。"""
        if gain == 1.0:
            return data
        arr = array.array('h')
        arr.frombytes(data)
        scaled = array.array('h', [
            max(-32768, min(32767, int(s * gain))) for s in arr
        ])
        return scaled.tobytes()
```

- [ ] **Step 5：執行測試，確認全部通過**
```
venv\Scripts\python.exe -m unittest tests.test_audio_processing -v
```
預期：13 個 tests，全部 OK。

- [ ] **Step 6：Commit**
```
git add main.py tests/test_audio_processing.py
git commit -m "feat: 實作 _compute_equalize_gain 和 _apply_gain_to_pcm，含單元測試"
```

---

## Task 7：修改 `_stop_recording` 鎖定進階設定快照

背景執行緒不能讀取 tkinter 變數，需在主執行緒停止時鎖定快照。

**Files:**
- Modify: `main.py`（`_stop_recording` 方法，約 line 596–608）

- [ ] **Step 1：在 `_stop_recording` 中補鎖定邏輯**

找到（約 line 600）：
```python
        # 在主執行緒鎖定模式，避免背景 _save_after_stop 從 tkinter StringVar 讀取
        self._save_mode = self.record_mode.get()
```

在它之後加：
```python
        self._save_output_mode    = self.output_mode.get()
        self._save_equalize       = self.equalize_enabled.get()
        self._save_gain_cap       = float(self.eq_gain_cap.get())
        self._save_filter_silence = self.eq_filter_silence.get()
```

- [ ] **Step 2：Commit**
```
git add main.py
git commit -m "feat: _stop_recording 鎖定進階設定快照供背景執行緒使用"
```

---

## Task 8：改寫 `_save_after_stop` 的 "both" 分支

**Files:**
- Modify: `main.py`（`_save_after_stop` 的 "both" 段，約 line 854–878）

- [ ] **Step 1：找到並替換整個 `else:  # "both"` 分支**

找到（約 line 854–878）：
```python
            else:  # "both"
                if not self.record_frames:
                    self.msg_queue.put(("error", "沒有錄到任何系統音訊"))
                    return

                if not self.mic_frames:
                    # 麥克風無資料（開啟失敗或立即斷線），退回純系統音訊並警告
                    self.msg_queue.put(("warning", "麥克風無資料，改以「電腦聲音」模式儲存"))
                    pcm_data    = b"".join(self.record_frames)
                    channels    = self.record_channels
                    sample_rate = self.record_sample_rate
                else:
                    if self.record_mic_rate != self.record_sample_rate:
                        # 取樣率不一致，混音仍繼續但聲速會有輕微偏差
                        self.msg_queue.put(("warning",
                            f"麥克風取樣率（{self.record_mic_rate} Hz）與系統音訊"
                            f"（{self.record_sample_rate} Hz）不一致，麥克風聲音可能略有偏差"))

                    self.msg_queue.put(("status", "混音中..."))
                    pcm_data = self._mix_pcm(
                        b"".join(self.record_frames), self.record_channels,
                        b"".join(self.mic_frames),    self.record_mic_channels,
                    )
                    channels    = self.record_channels
                    sample_rate = self.record_sample_rate
```

替換為：
```python
            else:  # "both"
                if not self.record_frames:
                    self.msg_queue.put(("error", "沒有錄到任何系統音訊"))
                    return

                if not self.mic_frames:
                    # 麥克風無資料，退回純系統音訊
                    self.msg_queue.put(("warning", "麥克風無資料，改以「電腦聲音」模式儲存"))
                    mp3_data = self._encode_to_mp3(
                        b"".join(self.record_frames),
                        self.record_channels, self.record_sample_rate)
                    filepath = self._save_file(mp3_data, base_name)
                    self.msg_queue.put(("saved", [filepath]))
                    return

                output_mode = self._save_output_mode
                saved_paths = []

                # ---- 獨立音軌 ----
                if output_mode in ("separate", "both"):
                    sys_mp3 = self._encode_to_mp3(
                        b"".join(self.record_frames),
                        self.record_channels, self.record_sample_rate)
                    mic_mp3 = self._encode_to_mp3(
                        b"".join(self.mic_frames),
                        self.record_mic_channels, self.record_mic_rate)
                    saved_paths.append(self._save_file(sys_mp3, base_name, "_system"))
                    saved_paths.append(self._save_file(mic_mp3, base_name, "_mic"))

                # ---- 合併音軌 ----
                if output_mode in ("merge", "both"):
                    if self.record_mic_rate != self.record_sample_rate:
                        self.msg_queue.put(("warning",
                            f"麥克風取樣率（{self.record_mic_rate} Hz）與系統音訊"
                            f"（{self.record_sample_rate} Hz）不一致，麥克風聲音可能略有偏差"))

                    self.msg_queue.put(("status", "混音中..."))
                    sys_pcm = b"".join(self.record_frames)
                    mic_pcm = b"".join(self.mic_frames)

                    if self._save_equalize:
                        sys_gain, mic_gain = self._compute_equalize_gain(
                            self.record_frames, self.mic_frames,
                            filter_silence=self._save_filter_silence,
                            gain_cap=self._save_gain_cap,
                        )
                        sys_pcm = self._apply_gain_to_pcm(sys_pcm, sys_gain)
                        mic_pcm = self._apply_gain_to_pcm(mic_pcm, mic_gain)

                    mixed_pcm = self._mix_pcm(
                        sys_pcm, self.record_channels,
                        mic_pcm, self.record_mic_channels)
                    merged_mp3 = self._encode_to_mp3(
                        mixed_pcm, self.record_channels, self.record_sample_rate)
                    saved_paths.append(self._save_file(merged_mp3, base_name))

                self.msg_queue.put(("saved", saved_paths))
                return
```

注意：加了 `return` 讓 "both" 分支提前結束，後面原本的 encoder/save 程式碼不執行。

- [ ] **Step 2：刪除 "both" 分支之後多餘的舊 encoder 程式碼**

此時 `_save_after_stop` 的 "both" 分支已完整處理自己的存檔與回報，後面的舊 `encoder = lameenc.Encoder()...` 段落只由 "system" / "mic" 分支（以及舊 "both" fallback）使用。確認 "both" 分支已加 `return`，舊段落不需改動。

- [ ] **Step 3：手動驗證三種輸出模式**

```
venv\Scripts\python.exe main.py
```
測試流程（錄約 5 秒）：
1. 模式選「系統 + 麥克風」，進階設定維持預設「合併一軌」→ 確認存出一個 `.mp3`
2. 進階設定改「獨立兩軌」→ 確認存出 `_system.mp3` 與 `_mic.mp3`
3. 進階設定改「兩個都要」→ 確認存出三個檔案，log 顯示三行 ✓

- [ ] **Step 4：Commit**
```
git add main.py
git commit -m "feat: 支援獨立音軌輸出（separate/both）與等化串接"
```

---

## Task 9：實作 `_show_advanced_settings` 彈窗

**Files:**
- Modify: `main.py`（取代 Task 4 留下的 stub）

- [ ] **Step 1：找到 stub 並替換**

找到：
```python
    def _show_advanced_settings(self):
        pass  # 由 Task 8 實作
```

替換為：
```python
    def _show_advanced_settings(self):
        win = tk.Toplevel(self.root)
        win.title("進階設定")
        win.resizable(False, False)
        win.grab_set()
        pad = {"padx": 14, "pady": 6}

        # 本地變數：pre-populate 目前設定，Cancel 不影響 app 狀態
        output_var  = tk.StringVar(value=self.output_mode.get())
        eq_var      = tk.BooleanVar(value=self.equalize_enabled.get())
        cap_var     = tk.IntVar(value=self.eq_gain_cap.get())
        filter_var  = tk.BooleanVar(value=self.eq_filter_silence.get())

        # ---- 輸出方式 ----
        frame_output = ttk.LabelFrame(
            win, text=" 輸出方式（僅「系統＋麥克風」模式有效） ", padding=10)
        frame_output.grid(row=0, column=0, sticky="ew", **pad)
        for text, val in [("合併一軌", "merge"),
                          ("獨立兩軌", "separate"),
                          ("兩個都要", "both")]:
            ttk.Radiobutton(frame_output, text=text,
                            variable=output_var, value=val).pack(anchor="w")

        # ---- 自動等化 ----
        frame_eq = ttk.LabelFrame(win, text=" 自動等化音量 ", padding=10)
        frame_eq.grid(row=1, column=0, sticky="ew", **pad)
        ttk.Checkbutton(
            frame_eq,
            text="自動等化音量（讓麥克風與系統音訊響度接近）",
            variable=eq_var,
        ).pack(anchor="w")
        ttk.Label(
            frame_eq,
            text="僅影響含混音的輸出；獨立音軌存原始音量",
            foreground="gray", font=("", 8),
        ).pack(anchor="w", pady=(2, 0))

        # ---- 等化進階設定 ----
        frame_adv = ttk.LabelFrame(win, text=" 等化進階設定 ", padding=10)
        frame_adv.grid(row=2, column=0, sticky="ew", **pad)
        frame_adv.columnconfigure(2, weight=1)

        tk.Label(frame_adv, text="增益上限：").grid(row=0, column=0, sticky="w")
        tk.Spinbox(frame_adv, from_=1, to=16,
                   textvariable=cap_var, width=5).grid(row=0, column=1, sticky="w", padx=(4, 0))
        tk.Label(frame_adv, text="x  (1～16)",
                 foreground="gray").grid(row=0, column=2, sticky="w", padx=(4, 0))

        ttk.Checkbutton(
            frame_adv,
            text="靜音過濾：排除靜音段再計算 RMS",
            variable=filter_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # ---- 確認 / 取消 ----
        frame_btns = tk.Frame(win)
        frame_btns.grid(row=3, column=0, pady=12)

        def confirm():
            self.output_mode.set(output_var.get())
            self.equalize_enabled.set(eq_var.get())
            self.eq_gain_cap.set(cap_var.get())
            self.eq_filter_silence.set(filter_var.get())
            win.destroy()

        ttk.Button(frame_btns, text="確認", command=confirm).pack(side="left", padx=8)
        ttk.Button(frame_btns, text="取消", command=win.destroy).pack(side="left")
        win.columnconfigure(0, weight=1)
```

- [ ] **Step 2：手動驗證彈窗**

```
venv\Scripts\python.exe main.py
```
確認項目：
1. 按「⚙ 進階設定」彈出對話框，三個 radio、checkbox、spinbox 都顯示正確
2. 改為「獨立兩軌」→ 確認 → 再開彈窗，確認值已儲存
3. 按「取消」不改動之前設定

- [ ] **Step 3：執行單元測試，確認等化函式仍通過**
```
venv\Scripts\python.exe -m unittest tests.test_audio_processing -v
```

- [ ] **Step 4：Commit**
```
git add main.py
git commit -m "feat: 實作進階設定彈窗（輸出方式、等化開關、進階參數）"
```

---

## 完成驗收清單

- [ ] 主視窗顯示兩條 VU Meter，錄音中動態變化，停止後歸零
- [ ] 「⚙ 進階設定」按鈕與「裝置設定與測試」並排
- [ ] 進階設定彈窗：三種輸出方式可選，確認後生效
- [ ] 「系統+麥克風」+ 獨立兩軌：存出 `_system.mp3` 與 `_mic.mp3`
- [ ] 「系統+麥克風」+ 兩個都要：存出三個檔案，log 各顯示一行 ✓
- [ ] 自動等化開啟時，混音後兩軌響度明顯接近
- [ ] 獨立音軌不受等化影響（原始音量）
- [ ] 13 個單元測試全部通過
