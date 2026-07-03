# VU Meter 靈敏度調整 + 爆音警示 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 VU meter 的 RMS→百分比換算改成更靈敏的公式（正常說話落在 20~50%），並在偵測到接近爆音的音訊峰值時，讓對應指示條短暫變紅警示。

**Architecture:** 沿用現有 `main.py` 單檔 tkinter 架構與既有的 `msg_queue` 背景執行緒→主執行緒通訊模式。新增兩個模組層級純函式（`_rms_to_display_pct`、`_compute_peak`／`_is_clipping`）取代目前寫死在四個位置的 `rms / 327.67` 算式；正式錄音畫面透過既有 `msg_queue` 新增兩種訊息類型並在 `_poll_queue` 做「保持顯示 1.5 秒」的衰減判斷；裝置測試畫面（無 `msg_queue`）沿用現有的 `win.after(0, ...)` + list-cell 閉包模式直接切換 `ttk.Progressbar` 樣式。

**Tech Stack:** Python 3.13（專案 venv）、tkinter/ttk、pytest（既有 `tests/test_audio_processing.py`）。

## Global Constraints

- 除數常數 `_VU_DISPLAY_DIVISOR = 80`（原 `327.67`）— 來自 spec `docs/superpowers/specs/2026-07-03-vu-meter-sensitivity-and-clip-warning-design.md`
- 爆音峰值閾值 `_CLIP_PEAK_THRESHOLD = int(32767 * 0.9)`（約 29491）
- 爆音警示保持時間 `_CLIP_WARNING_HOLD_SECS = 1.5`
- 新常數命名沿用既有模組慣例（底線前綴，定義於 `_SILENCE_RMS_THRESHOLD` 旁）
- 錄音存檔的原始 PCM 不可變動，本次只改顯示端
- 靜音偵測（10 秒警告橫幅）邏輯不可受影響
- 正式錄音畫面與「裝置設定與測試」畫面的換算公式與爆音警示行為必須一致
- 遵循既有程式碼慣例：背景執行緒不可直接操作非執行緒安全的 UI（`Progressbar.config(style=...)` 這類呼叫一律透過 `msg_queue`／`win.after(0, ...)` 轉回主執行緒，`tk.DoubleVar.set()` 例外，因為既有註解已說明其為 CPython 執行緒安全）
- 每個任務完成後執行 `venv/Scripts/python.exe -m pytest tests/ -v`，確認既有 13 個測試仍全數通過

---

### Task 1: 新增 VU 顯示換算與爆音峰值偵測的純函式

**Files:**
- Modify: `main.py:26-31`（常數區）、`main.py:84-94`（`_compute_rms` 旁新增函式）
- Test: `tests/test_audio_processing.py`

**Interfaces:**
- Produces:
  - `_VU_DISPLAY_DIVISOR: float`（模組常數，值 80）
  - `_CLIP_PEAK_THRESHOLD: int`（模組常數，值 `int(32767 * 0.9)` = 29491）
  - `_CLIP_WARNING_HOLD_SECS: float`（模組常數，值 1.5）
  - `_rms_to_display_pct(rms: float) -> float`：回傳 0.0~100.0
  - `_compute_peak(data: bytes) -> int`：回傳 0~32767
  - `_is_clipping(peak: int) -> bool`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_audio_processing.py` 的 import 行加入新符號，並新增三個測試類別。

先修改 import：
```python
from main import (
    MeetingRecorderApp,
    _compute_rms,
    _rms_to_display_pct,
    _compute_peak,
    _is_clipping,
    _CLIP_PEAK_THRESHOLD,
)
```

在檔案最後 `if __name__ == "__main__":` 之前加入：
```python
class TestRmsToDisplayPct(unittest.TestCase):

    def test_zero_rms_is_zero_pct(self):
        self.assertEqual(_rms_to_display_pct(0.0), 0.0)

    def test_typical_speech_rms_lands_in_20_to_50_pct(self):
        self.assertAlmostEqual(_rms_to_display_pct(1600.0), 20.0, places=1)
        self.assertAlmostEqual(_rms_to_display_pct(4000.0), 50.0, places=1)

    def test_full_scale_rms_clamped_to_100(self):
        self.assertEqual(_rms_to_display_pct(32767.0), 100.0)

    def test_rms_exceeding_divisor_range_clamped_to_100(self):
        self.assertEqual(_rms_to_display_pct(100000.0), 100.0)


class TestComputePeak(unittest.TestCase):

    def test_silence_peak_is_zero(self):
        self.assertEqual(_compute_peak(make_pcm(0)), 0)

    def test_positive_peak(self):
        data = struct.pack("4h", 100, 500, -300, 200)
        self.assertEqual(_compute_peak(data), 500)

    def test_negative_peak_uses_absolute_value(self):
        data = struct.pack("4h", 100, -32000, 300, 200)
        self.assertEqual(_compute_peak(data), 32000)

    def test_empty_data_returns_zero(self):
        self.assertEqual(_compute_peak(b""), 0)


class TestIsClipping(unittest.TestCase):

    def test_below_threshold_not_clipping(self):
        self.assertFalse(_is_clipping(29000))

    def test_at_threshold_is_clipping(self):
        self.assertTrue(_is_clipping(_CLIP_PEAK_THRESHOLD))

    def test_full_scale_is_clipping(self):
        self.assertTrue(_is_clipping(32767))
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `venv/Scripts/python.exe -m pytest tests/test_audio_processing.py -v`
Expected: `ImportError`（`_rms_to_display_pct` 等符號在 `main.py` 中尚不存在）

- [ ] **Step 3: 在 `main.py` 新增常數與函式**

在 `main.py:26-31` 的常數區塊（`_SILENCE_WARNING_SECS` 之後）加入：
```python
_VU_DISPLAY_DIVISOR     = 80     # VU meter 顯示除數：正常說話 RMS 換算後約落在 20~50%
_CLIP_PEAK_THRESHOLD    = int(32767 * 0.9)   # 峰值達此值視為接近爆音（約 29491）
_CLIP_WARNING_HOLD_SECS = 1.5    # 爆音警示保持顯示的秒數，避免瞬間峰值一閃即逝
```

在 `main.py:94`（`_compute_rms` 函式結尾）之後加入：
```python
def _rms_to_display_pct(rms: float) -> float:
    """把 RMS（0~32767）換算成 VU meter 顯示用的 0~100 百分比。"""
    return min(100.0, rms / _VU_DISPLAY_DIVISOR)


def _compute_peak(data: bytes) -> int:
    """
    計算 PCM Int16 音訊資料的峰值（樣本最大絕對值）。
    回傳範圍 0 ~ 32767，用於偵測爆音，與 RMS（平均音量）分開判斷。
    """
    num_samples = len(data) // 2
    if num_samples == 0:
        return 0
    samples = struct.unpack(f"{num_samples}h", data)
    return max(abs(s) for s in samples)


def _is_clipping(peak: int) -> bool:
    """峰值是否達到爆音警示閾值。"""
    return peak >= _CLIP_PEAK_THRESHOLD
```

- [ ] **Step 4: 執行測試確認通過**

Run: `venv/Scripts/python.exe -m pytest tests/test_audio_processing.py -v`
Expected: 新增的 11 個測試全數 PASS，原有 13 個測試不受影響

- [ ] **Step 5: Commit**

```bash
git add main.py tests/test_audio_processing.py
git commit -m "feat: 新增 VU 顯示換算與爆音峰值偵測函式"
```

---

### Task 2: 主畫面 UI 準備 — 爆音樣式與 Progressbar 參照

**Files:**
- Modify: `main.py:272-292`（`_build_ui` 內 VU Meter 區塊）

**Interfaces:**
- Consumes: 無（純 UI 準備）
- Produces:
  - `self.vu_system_pb: ttk.Progressbar`（供 Task 5 切換樣式用）
  - `self.vu_mic_pb: ttk.Progressbar`
  - `self._sys_clip_until: float`、`self._mic_clip_until: float`（epoch 秒數，供 Task 5 判斷 hold 是否過期）
  - `self._sys_clip_active: bool`、`self._mic_clip_active: bool`（目前是否正在顯示紅色警示，避免每 100ms 重複呼叫 `.config()`）
  - ttk 樣式 `"Clip.Horizontal.TProgressbar"`（全域註冊，`_show_device_test` 對話框亦可直接引用同名樣式）

- [ ] **Step 1: 修改 VU Meter 區塊，捕捉 Progressbar 參照並註冊爆音樣式**

把 `main.py:272-292`：
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
        self.mic_offline_label = ttk.Label(vu_frame, text="", foreground="red", font=("", 9))
        self.mic_offline_label.grid(row=1, column=3, padx=(6, 0), pady=(4, 0))
```

改為：
```python
        # VU Meter
        vu_frame = tk.Frame(frame_btn)
        vu_frame.pack(pady=(12, 0))

        ttk.Style().configure("Clip.Horizontal.TProgressbar", background="red")

        tk.Label(vu_frame, text="系統音訊", width=8, anchor="w",
                 font=("", 9)).grid(row=0, column=0, padx=(0, 6))
        self.vu_system_pb = ttk.Progressbar(vu_frame, variable=self.vu_system_var,
                        maximum=100, length=200)
        self.vu_system_pb.grid(row=0, column=1)
        self.vu_system_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                        foreground="gray", font=("Consolas", 9))
        self.vu_system_pct.grid(row=0, column=2, padx=(4, 0))

        tk.Label(vu_frame, text="麥克風", width=8, anchor="w",
                 font=("", 9)).grid(row=1, column=0, padx=(0, 6), pady=(4, 0))
        self.vu_mic_pb = ttk.Progressbar(vu_frame, variable=self.vu_mic_var,
                        maximum=100, length=200)
        self.vu_mic_pb.grid(row=1, column=1, pady=(4, 0))
        self.vu_mic_pct = ttk.Label(vu_frame, text="  0%", width=5,
                                     foreground="gray", font=("Consolas", 9))
        self.vu_mic_pct.grid(row=1, column=2, padx=(4, 0), pady=(4, 0))
        self.mic_offline_label = ttk.Label(vu_frame, text="", foreground="red", font=("", 9))
        self.mic_offline_label.grid(row=1, column=3, padx=(6, 0), pady=(4, 0))

        self._sys_clip_until  = 0.0
        self._mic_clip_until  = 0.0
        self._sys_clip_active = False
        self._mic_clip_active = False
```

- [ ] **Step 2: 執行既有測試確認沒有壞掉**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 24 個測試（13 舊 + 11 新）全數 PASS（此任務不新增測試，因為是純 UI 佈局變更，既有測試套件不涵蓋 tkinter widget 建立，符合現行慣例）

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: 主畫面 VU meter 捕捉 Progressbar 參照並註冊爆音警示樣式"
```

---

### Task 3: 系統音訊錄音迴圈套用新換算公式與爆音偵測

**Files:**
- Modify: `main.py:934-937`（`_record_worker` 內）

**Interfaces:**
- Consumes: `_rms_to_display_pct`、`_compute_peak`、`_is_clipping`（Task 1）
- Produces: 新增 `msg_queue` 訊息類型 `("vu_system_clip", True)`（供 Task 5 的 `_poll_queue` 消費）

- [ ] **Step 1: 修改 `_record_worker` 的 RMS/VU 區塊**

把 `main.py:934-937`：
```python
                    # ---- 靜音偵測 ----
                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_system", min(100.0, rms / 327.67)))
                    if rms < SILENCE_RMS_THRESHOLD:
```

改為：
```python
                    # ---- 靜音偵測 ----
                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_system", _rms_to_display_pct(rms)))
                    if _is_clipping(_compute_peak(data)):
                        self.msg_queue.put(("vu_system_clip", True))
                    if rms < SILENCE_RMS_THRESHOLD:
```

- [ ] **Step 2: 執行既有測試確認沒有壞掉**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 24 個測試全數 PASS（`_record_worker` 依賴實際音訊裝置，無法在無硬體環境下自動化測試，人工驗證見 Task 7）

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: 系統音訊錄音迴圈套用新 VU 換算公式與爆音峰值偵測"
```

---

### Task 4: 麥克風錄音迴圈套用新換算公式與爆音偵測

**Files:**
- Modify: `main.py:1106-1109`（`_record_mic_worker` 內）

**Interfaces:**
- Consumes: `_rms_to_display_pct`、`_compute_peak`、`_is_clipping`（Task 1）
- Produces: 新增 `msg_queue` 訊息類型 `("vu_mic_clip", True)`（供 Task 5 的 `_poll_queue` 消費）

- [ ] **Step 1: 修改 `_record_mic_worker` 的 RMS/VU 區塊**

把 `main.py:1106-1109`：
```python
                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_mic", min(100.0, rms / 327.67)))

                    if check_silence:
```

改為：
```python
                    rms = _compute_rms(data)
                    self.msg_queue.put(("vu_mic", _rms_to_display_pct(rms)))
                    if _is_clipping(_compute_peak(data)):
                        self.msg_queue.put(("vu_mic_clip", True))

                    if check_silence:
```

- [ ] **Step 2: 執行既有測試確認沒有壞掉**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 24 個測試全數 PASS

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: 麥克風錄音迴圈套用新 VU 換算公式與爆音峰值偵測"
```

---

### Task 5: `_poll_queue` 處理爆音訊息並做 1.5 秒 hold 顯示，`_reset_ui_after_stop` 還原樣式

**Files:**
- Modify: `main.py:1494-1509`（`_reset_ui_after_stop`）、`main.py:1568-1605`（`_poll_queue`）

**Interfaces:**
- Consumes: `self.vu_system_pb`、`self.vu_mic_pb`、`self._sys_clip_until`、`self._mic_clip_until`、`self._sys_clip_active`、`self._mic_clip_active`（Task 2）；`"vu_system_clip"` / `"vu_mic_clip"` 訊息（Task 3、4）；`_CLIP_WARNING_HOLD_SECS`（Task 1）
- Produces: `self._update_clip_style(progressbar, until_attr, active_attr)` 方法（純 UI 副作用，不對外提供回傳值）

- [ ] **Step 1: 在 `_reset_ui_after_stop` 加入樣式還原**

把 `main.py:1494-1509`：
```python
    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self.is_paused = False
        self._elapsed_before_pause = 0.0
        self.btn_record.config(state="normal", text="⏺  開始錄音")
        self.btn_pause.config(text="⏸  暫停", command=self._pause_recording, state="normal")
        self.btn_discard.config(state="normal")
        self._secondary_row.pack_forget()
        self.timer_label.config(text="00:00", foreground="gray")
        self.silence_banner.grid_remove()
        self._set_mode_radios_state("normal")
        self.vu_system_var.set(0)
        self.vu_mic_var.set(0)
        self.vu_system_pct.config(text=f"{0:3d}%")
        self.vu_mic_pct.config(text=f"{0:3d}%")
        self.mic_offline_label.config(text="")
```

改為（新增最後 4 行）：
```python
    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self.is_paused = False
        self._elapsed_before_pause = 0.0
        self.btn_record.config(state="normal", text="⏺  開始錄音")
        self.btn_pause.config(text="⏸  暫停", command=self._pause_recording, state="normal")
        self.btn_discard.config(state="normal")
        self._secondary_row.pack_forget()
        self.timer_label.config(text="00:00", foreground="gray")
        self.silence_banner.grid_remove()
        self._set_mode_radios_state("normal")
        self.vu_system_var.set(0)
        self.vu_mic_var.set(0)
        self.vu_system_pct.config(text=f"{0:3d}%")
        self.vu_mic_pct.config(text=f"{0:3d}%")
        self.mic_offline_label.config(text="")
        self.vu_system_pb.config(style="Horizontal.TProgressbar")
        self.vu_mic_pb.config(style="Horizontal.TProgressbar")
        self._sys_clip_active = False
        self._mic_clip_active = False
```

- [ ] **Step 2: 新增 `_update_clip_style` 方法**

在 `_reset_ui_after_stop` 方法之後、`_poll_queue` 方法之前插入：
```python
    def _update_clip_style(self, progressbar, pct_label, until_attr, active_attr):
        """
        依 hold 截止時間決定爆音警示是否該顯示，只有狀態改變時才呼叫 .config()，
        避免每 100ms tick 都重繪 widget。
        """
        active = time.time() < getattr(self, until_attr)
        if active == getattr(self, active_attr):
            return
        setattr(self, active_attr, active)
        style = "Clip.Horizontal.TProgressbar" if active else "Horizontal.TProgressbar"
        progressbar.config(style=style)
        pct_label.config(foreground="red" if active else "gray")
```

- [ ] **Step 3: 在 `_poll_queue` 新增訊息處理與收尾時的 hold 檢查**

把 `main.py` 中 `vu_mic` 的處理區塊：
```python
                elif msg_type == "vu_mic":
                    val = max(0.0, min(100.0, float(data)))
                    self.vu_mic_var.set(val)
                    self.vu_mic_pct.config(text=f"{int(val):3d}%")
```

改為（新增 clip 訊息分支）：
```python
                elif msg_type == "vu_mic":
                    val = max(0.0, min(100.0, float(data)))
                    self.vu_mic_var.set(val)
                    self.vu_mic_pct.config(text=f"{int(val):3d}%")

                elif msg_type == "vu_system_clip":
                    self._sys_clip_until = time.time() + _CLIP_WARNING_HOLD_SECS

                elif msg_type == "vu_mic_clip":
                    self._mic_clip_until = time.time() + _CLIP_WARNING_HOLD_SECS
```

把函式收尾的：
```python
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)
```

改為：
```python
        except queue.Empty:
            pass

        self._update_clip_style(self.vu_system_pb, self.vu_system_pct,
                                 "_sys_clip_until", "_sys_clip_active")
        self._update_clip_style(self.vu_mic_pb, self.vu_mic_pct,
                                 "_mic_clip_until", "_mic_clip_active")

        self.root.after(100, self._poll_queue)
```

也把 `_poll_queue` 開頭 docstring 的訊息類型清單（約 `main.py:1516-1529`）補上兩行：
```
          vu_system_clip  — 系統音訊接近爆音，data = True（觸發 1.5 秒紅色警示）
          vu_mic_clip     — 麥克風接近爆音，data = True（觸發 1.5 秒紅色警示）
```

- [ ] **Step 4: 執行既有測試確認沒有壞掉**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 24 個測試全數 PASS

- [ ] **Step 5: Commit**

```bash
git add main.py
git commit -m "feat: _poll_queue 處理爆音警示訊息並以 1.5 秒 hold 切換指示條樣式"
```

---

### Task 6: 「裝置設定與測試」畫面套用相同換算公式與爆音警示

**Files:**
- Modify: `main.py:439-517`（麥克風測試區塊）、`main.py:532-604`（系統音訊測試區塊）

**Interfaces:**
- Consumes: `_rms_to_display_pct`、`_compute_peak`、`_is_clipping`、`_CLIP_WARNING_HOLD_SECS`（Task 1）；ttk 樣式 `"Clip.Horizontal.TProgressbar"`（Task 2 已全域註冊）
- Produces: 無對外介面（純 UI 行為，維持既有 `mic_running` / `sys_running` list-cell 模式）

- [ ] **Step 1: 麥克風測試區塊 — 捕捉 Progressbar 參照並新增 clip 狀態 cell**

把 `main.py:439-445`：
```python
        mic_level = tk.DoubleVar(value=0)
        ttk.Progressbar(frame_mic, variable=mic_level,
                        maximum=100, length=340).grid(row=1, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 4))
        mic_status = ttk.Label(frame_mic, text="請對著麥克風說話，確認音量指示條有所反應",
                               foreground="gray")
```

改為：
```python
        mic_level = tk.DoubleVar(value=0)
        mic_pb = ttk.Progressbar(frame_mic, variable=mic_level,
                        maximum=100, length=340)
        mic_pb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        mic_clip_until  = [0.0]
        mic_clip_active = [False]
        mic_status = ttk.Label(frame_mic, text="請對著麥克風說話，確認音量指示條有所反應",
                               foreground="gray")
```

- [ ] **Step 2: 麥克風測試迴圈 — 套用新換算公式與爆音偵測**

把 `main.py:475-481`：
```python
                while mic_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        mic_level.set(min(100, rms / 327.67))
                    except Exception:
                        break  # 視窗已關閉，DoubleVar 失效，結束迴圈
```

改為：
```python
                while mic_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        mic_level.set(_rms_to_display_pct(rms))
                    except Exception:
                        break  # 視窗已關閉，DoubleVar 失效，結束迴圈

                    if _is_clipping(_compute_peak(data)):
                        mic_clip_until[0] = time.time() + _CLIP_WARNING_HOLD_SECS
                    clip_active = time.time() < mic_clip_until[0]
                    if clip_active != mic_clip_active[0]:
                        mic_clip_active[0] = clip_active
                        style = "Clip.Horizontal.TProgressbar" if clip_active else "Horizontal.TProgressbar"
                        try:
                            win.after(0, lambda s=style: mic_pb.config(style=s))
                        except Exception:
                            break  # 視窗已關閉，結束迴圈
```

- [ ] **Step 3: 麥克風測試 `finally` 區塊 — 還原樣式**

把 `main.py:490-500`：
```python
            finally:
                p.terminate()
                mic_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_mic.config(state="normal"),
                        btn_mic_stop.config(state="disabled"),
                        mic_level.set(0),
                    ))
                except Exception:
                    pass
```

改為：
```python
            finally:
                p.terminate()
                mic_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_mic.config(state="normal"),
                        btn_mic_stop.config(state="disabled"),
                        mic_level.set(0),
                        mic_pb.config(style="Horizontal.TProgressbar"),
                    ))
                except Exception:
                    pass
```

- [ ] **Step 4: 系統音訊測試區塊 — 捕捉 Progressbar 參照並新增 clip 狀態 cell**

把 `main.py:532-539`：
```python
        sys_level = tk.DoubleVar(value=0)
        ttk.Progressbar(frame_sys, variable=sys_level,
                        maximum=100, length=340).grid(row=1, column=0, columnspan=2,
                                                       sticky="ew", pady=(0, 4))
        sys_status = ttk.Label(frame_sys,
                               text="請先播放任意音訊（音樂、影片等），再按下開始測試",
                               foreground="gray")
```

改為：
```python
        sys_level = tk.DoubleVar(value=0)
        sys_pb = ttk.Progressbar(frame_sys, variable=sys_level,
                        maximum=100, length=340)
        sys_pb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        sys_clip_until  = [0.0]
        sys_clip_active = [False]
        sys_status = ttk.Label(frame_sys,
                               text="請先播放任意音訊（音樂、影片等），再按下開始測試",
                               foreground="gray")
```

- [ ] **Step 5: 系統音訊測試迴圈 — 套用新換算公式與爆音偵測**

把 `main.py:561-567`：
```python
                while sys_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        sys_level.set(min(100, rms / 327.67))
                    except Exception:
                        break  # 視窗已關閉，結束迴圈
```

改為：
```python
                while sys_running[0]:
                    data = stream.read(512, exception_on_overflow=False)
                    rms  = _compute_rms(data)
                    try:
                        sys_level.set(_rms_to_display_pct(rms))
                    except Exception:
                        break  # 視窗已關閉，結束迴圈

                    if _is_clipping(_compute_peak(data)):
                        sys_clip_until[0] = time.time() + _CLIP_WARNING_HOLD_SECS
                    clip_active = time.time() < sys_clip_until[0]
                    if clip_active != sys_clip_active[0]:
                        sys_clip_active[0] = clip_active
                        style = "Clip.Horizontal.TProgressbar" if clip_active else "Horizontal.TProgressbar"
                        try:
                            win.after(0, lambda s=style: sys_pb.config(style=s))
                        except Exception:
                            break  # 視窗已關閉，結束迴圈
```

- [ ] **Step 6: 系統音訊測試 `finally` 區塊 — 還原樣式**

把 `main.py:576-586`：
```python
            finally:
                p.terminate()
                sys_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_sys.config(state="normal"),
                        btn_sys_stop.config(state="disabled"),
                        sys_level.set(0),
                    ))
                except Exception:
                    pass
```

改為：
```python
            finally:
                p.terminate()
                sys_running[0] = False
                try:
                    win.after(0, lambda: (
                        btn_sys.config(state="normal"),
                        btn_sys_stop.config(state="disabled"),
                        sys_level.set(0),
                        sys_pb.config(style="Horizontal.TProgressbar"),
                    ))
                except Exception:
                    pass
```

- [ ] **Step 7: 執行既有測試確認沒有壞掉**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 24 個測試全數 PASS

- [ ] **Step 8: Commit**

```bash
git add main.py
git commit -m "feat: 裝置測試畫面套用新 VU 換算公式與爆音警示，與正式錄音畫面一致"
```

---

### Task 7: 人工驗證與文件更新

**Files:**
- Modify: `CHANGELOG.md`（新增更新記錄章節）

**Interfaces:**
- Consumes: 無
- Produces: 無（收尾任務）

- [ ] **Step 1: 雙擊 `meeting_recorder啟動器.bat` 啟動程式，人工驗證下列情境**

對照 spec `docs/superpowers/specs/2026-07-03-vu-meter-sensitivity-and-clip-warning-design.md` 第五節逐一驗證：

1. 選「電腦聲音」或「麥克風」模式開始錄音，正常說話／播放音樂時，指示條落在約 20~50% 區間跳動（而非原本的 1~10%）
2. 對著麥克風大聲說話製造爆音（或把系統音量開到最大播放尖峰音訊），確認指示條與旁邊百分比文字變紅，停止大聲後約 1.5 秒恢復原色
3. 進「🔧 裝置設定與測試」，麥克風測試與系統音訊測試的指示條行為（靈敏度、爆音變紅）與正式錄音畫面一致
4. 完全靜音錄音超過 10 秒，確認原有靜音警告橫幅仍正常出現，不受本次改動影響
5. 錄一段測試音檔並存檔，確認 MP3 播放音質與音量正常，沒有因為這次改動而受影響

- [ ] **Step 2: 若驗證中發現除數 80 或閾值 90% 需要微調**

直接修改 `main.py` 中的 `_VU_DISPLAY_DIVISOR` / `_CLIP_PEAK_THRESHOLD` 常數值，重新執行 Step 1 驗證，確認後另外建立一個小 commit（訊息例如 `tune: 調整 VU meter 除數為 <值>`）。

- [ ] **Step 3: 更新 `CHANGELOG.md`**

在 `CHANGELOG.md` 的「現狀總覽」已完成功能清單中，把：
```
- [x] VU Meter 即時音量顯示（系統音訊 + 麥克風各一條）
```
改為：
```
- [x] VU Meter 即時音量顯示（線性換算，正常音量約 20~50%；峰值接近滿刻度時短暫變紅警示爆音）
```

並在「更新記錄」區塊最上方（`## 更新記錄` 標題之後）新增：
```markdown
### 2026-07-03
- 調整：VU meter 顯示公式除數改小，正常說話音量從原本的 1~10% 提升到約 20~50%，更容易一眼判斷是否有在收音
- 新增：音訊峰值超過滿刻度 90% 時，對應指示條與百分比文字短暫變紅警示爆音（保持 1.5 秒）
- 套用範圍：正式錄音畫面與「裝置設定與測試」畫面的麥克風／系統音訊測試指示條，兩邊行為一致
```

- [ ] **Step 4: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs: CHANGELOG 記錄 VU meter 靈敏度調整與爆音警示"
```
