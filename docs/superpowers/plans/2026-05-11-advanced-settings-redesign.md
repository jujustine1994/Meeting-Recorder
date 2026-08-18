# Advanced Settings Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 將「輸出方式」搬到主畫面（動態顯示）、進階設定重組為「音質」+「混音」兩區（含 gray-out logic）、以及將硬編碼參數提取為模組級常數。

**Architecture:** 所有改動都在 `main.py` 這一支檔案。分四個獨立步驟：先提取常數，再加 bit_rate 變數，再改主畫面 UI，最後重寫進階設定彈窗。每步可以獨立 commit 且不破壞現有功能。

**Tech Stack:** Python 3, tkinter, lameenc, pyaudiowpatch

---

### Task 1：提取模組級 config 常數

**Files:**
- Modify: `main.py` — 第 24 行前（import 區之後）、`_mix_pcm`、`_compute_equalize_gain`、`_record_worker`、`_record_mic_worker`

- [ ] **Step 1：在 `main.py` 第 24 行（`# ---- CTH Banner` 之前）插入常數區塊**

```python
# ---- 可自行調整的進階參數（直接編輯此區段） ----
_DEFAULT_BIT_RATE       = 128    # kbps：MP3 位元率預設值，可在進階設定中更改
_MIX_WEIGHT_SYSTEM      = 0.6    # 系統音軌混音權重（兩軌混音時）
_MIX_WEIGHT_MIC         = 0.6    # 麥克風混音權重
_SILENCE_RMS_THRESHOLD  = 100    # 靜音判斷 RMS 閾值（Int16 0~32767）
_SILENCE_WARNING_SECS   = 10     # 連續靜音幾秒後顯示警告
```

- [ ] **Step 2：`_mix_pcm` 替換硬編碼權重（約 line 1049）**

舊：
```python
max(-32768, min(32767, int(sys_arr[i] * 0.6 + mic_arr[i] * 0.6)))
```
新：
```python
max(-32768, min(32767, int(sys_arr[i] * _MIX_WEIGHT_SYSTEM + mic_arr[i] * _MIX_WEIGHT_MIC)))
```

- [ ] **Step 3：`_compute_equalize_gain` 替換 THRESHOLD（約 line 1063）**

舊：
```python
        THRESHOLD = 100
```
新：
```python
        THRESHOLD = _SILENCE_RMS_THRESHOLD
```

- [ ] **Step 4：`_record_worker` 替換兩個本地常數（約 line 868-869）**

舊：
```python
            SILENCE_RMS_THRESHOLD = 100
            SILENCE_WARNING_SECS  = 10
```
新：
```python
            SILENCE_RMS_THRESHOLD = _SILENCE_RMS_THRESHOLD
            SILENCE_WARNING_SECS  = _SILENCE_WARNING_SECS
```

- [ ] **Step 5：`_record_mic_worker` 做同樣替換（約 line 966-967）**

舊：
```python
            SILENCE_RMS_THRESHOLD = 100
            SILENCE_WARNING_SECS  = 10
```
新：
```python
            SILENCE_RMS_THRESHOLD = _SILENCE_RMS_THRESHOLD
            SILENCE_WARNING_SECS  = _SILENCE_WARNING_SECS
```

- [ ] **Step 6：跑現有測試確認沒有回退**

```
cd "C:\Users\CTH\Documents\Code\meeting recorder"
venv\Scripts\python.exe -m pytest tests/ -v
```
預期：全部 PASS（13 個測試）

- [ ] **Step 7：Commit**

```
git add main.py
git commit -m "refactor: 提取 config 常數（混音權重、靜音閾值、警告秒數）"
```

---

### Task 2：可設定的 MP3 位元率

**Files:**
- Modify: `main.py` — `__init__`、`_stop_recording`、`_encode_to_mp3`、`_save_after_stop`（5 個呼叫點）

- [ ] **Step 1：`__init__` 加兩個變數（在 `self._save_filter_silence` 後面，約 line 117）**

```python
        self._save_bit_rate: int = _DEFAULT_BIT_RATE   # 停止時從主執行緒鎖定
```

以及在 tkinter 變數區（`self.output_mode` 附近，約 line 132）加：

```python
        self.bit_rate          = tk.IntVar(value=_DEFAULT_BIT_RATE)
```

- [ ] **Step 2：`_stop_recording` 在快照區加一行（現有 5 行快照的最後，約 line 827）**

舊（結尾）：
```python
        self._save_filter_silence = self.eq_filter_silence.get()
```
新：
```python
        self._save_filter_silence = self.eq_filter_silence.get()
        self._save_bit_rate       = self.bit_rate.get()
```

- [ ] **Step 3：`_encode_to_mp3` 加 `bit_rate` 參數（約 line 1099）**

舊：
```python
    def _encode_to_mp3(self, pcm_data: bytes, channels: int, sample_rate: int) -> bytes:
        encoder = lameenc.Encoder()
        encoder.set_bit_rate(128)
```
新：
```python
    def _encode_to_mp3(self, pcm_data: bytes, channels: int, sample_rate: int,
                       bit_rate: int = _DEFAULT_BIT_RATE) -> bytes:
        encoder = lameenc.Encoder()
        encoder.set_bit_rate(bit_rate)
```

- [ ] **Step 4：`_save_after_stop` — 更新 5 個 `_encode_to_mp3` 呼叫，全部加 `bit_rate=self._save_bit_rate`**

呼叫點 1（fallback，約 line 1189）：
```python
                    mp3_data = self._encode_to_mp3(
                        b"".join(self.record_frames),
                        self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate)
```

呼叫點 2（sys_mp3，約 line 1203）：
```python
                    sys_mp3 = self._encode_to_mp3(
                        b"".join(sys_frames_snap),
                        self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate)
```

呼叫點 3（mic_mp3，約 line 1206）：
```python
                    mic_mp3 = self._encode_to_mp3(
                        b"".join(mic_frames_snap),
                        self.record_mic_channels, self.record_mic_rate,
                        bit_rate=self._save_bit_rate)
```

呼叫點 4（merged_mp3，約 line 1235）：
```python
                    merged_mp3 = self._encode_to_mp3(
                        mixed_pcm, self.record_channels, self.record_sample_rate,
                        bit_rate=self._save_bit_rate)
```

呼叫點 5（system/mic 模式，約 line 1245）：
```python
            mp3_data = self._encode_to_mp3(pcm_data, channels, sample_rate,
                                           bit_rate=self._save_bit_rate)
```

- [ ] **Step 5：跑現有測試**

```
venv\Scripts\python.exe -m pytest tests/ -v
```
預期：全部 PASS

- [ ] **Step 6：Commit**

```
git add main.py
git commit -m "feat: 可設定 MP3 位元率（預設 128，進階設定可改 192/320）"
```

---

### Task 3：主畫面內嵌輸出格式選項

**Files:**
- Modify: `main.py` — `_build_ui` 的 frame_mode 區段（約 line 161-182）

- [ ] **Step 1：在 `_build_ui` 的 `frame_mode` 區段，`?` 按鈕之後加內嵌輸出格式列**

找到這段（約 line 179-182）：
```python
        ttk.Button(
            frame_mode, text=" ? ", width=3,
            command=self._show_mode_help,
        ).pack(side="right")
```

在其後（仍在 `# row=1 錄音模式` 的 block 內）加：

```python
        # 輸出格式：僅「系統 + 麥克風」時顯示
        self._frame_output_inline = tk.Frame(frame_mode)
        ttk.Label(
            self._frame_output_inline, text="輸出格式：",
            font=("", 9), foreground="#555555",
        ).pack(side="left", padx=(0, 6))
        self._output_radios = []
        for text, val in [("合併一軌", "merge"), ("獨立兩軌", "separate"), ("兩個都要", "both")]:
            rb = ttk.Radiobutton(
                self._frame_output_inline, text=text,
                variable=self.output_mode, value=val,
            )
            rb.pack(side="left", padx=(0, 12))
            self._output_radios.append(rb)

        def _on_mode_change(*_):
            if self.record_mode.get() == "both":
                self._frame_output_inline.pack(fill="x", pady=(8, 0))
            else:
                self._frame_output_inline.pack_forget()

        self.record_mode.trace_add("write", _on_mode_change)
```

- [ ] **Step 2：手動測試**

啟動程式，確認：
1. 預設「電腦聲音」→ 輸出格式列不出現
2. 切到「系統 + 麥克風」→ 「輸出格式：● 合併一軌 ○ 獨立兩軌 ○ 兩個都要」出現在 frame_mode 底部
3. 切回其他模式 → 消失
4. 在「系統 + 麥克風」模式選「獨立兩軌」後開始錄音，停止後確認存出 `_system.mp3` 和 `_mic.mp3`

- [ ] **Step 3：Commit**

```
git add main.py
git commit -m "feat: 主畫面內嵌輸出格式選項（系統＋麥克風模式下動態顯示）"
```

---

### Task 4：重寫進階設定彈窗

**Files:**
- Modify: `main.py` — `_show_advanced_settings` 完整替換（約 line 592-658）

- [ ] **Step 1：完整替換 `_show_advanced_settings`**

找到方法開頭：
```python
    def _show_advanced_settings(self):
        win = tk.Toplevel(self.root)
        win.title("進階設定")
```

整個方法替換為：

```python
    def _show_advanced_settings(self):
        win = tk.Toplevel(self.root)
        win.title("進階設定")
        win.resizable(False, False)
        win.grab_set()
        pad = {"padx": 14, "pady": 6}

        # 本地變數：pre-populate 目前設定，Cancel 不影響 app 狀態
        br_var     = tk.IntVar(value=self.bit_rate.get())
        eq_var     = tk.BooleanVar(value=self.equalize_enabled.get())
        cap_var    = tk.IntVar(value=self.eq_gain_cap.get())
        filter_var = tk.BooleanVar(value=self.eq_filter_silence.get())

        # ---- 音質 ----
        frame_quality = ttk.LabelFrame(win, text=" 音質 ", padding=10)
        frame_quality.grid(row=0, column=0, sticky="ew", **pad)
        for label, val in [("128 kbps（標準，一般會議適用）", 128),
                            ("192 kbps（較好音質）",           192),
                            ("320 kbps（最高，後製 / Podcast）", 320)]:
            ttk.Radiobutton(frame_quality, text=label,
                            variable=br_var, value=val).pack(anchor="w")

        # ---- 混音設定 ----
        frame_mix = ttk.LabelFrame(
            win, text=" 混音設定（僅「系統＋麥克風」有效） ", padding=10)
        frame_mix.grid(row=1, column=0, sticky="ew", **pad)

        ttk.Checkbutton(
            frame_mix,
            text="自動等化音量（讓兩軌響度接近）",
            variable=eq_var,
        ).pack(anchor="w")
        ttk.Label(
            frame_mix,
            text="僅影響合併輸出；獨立音軌保留原始音量",
            foreground="gray", font=("", 8),
        ).pack(anchor="w", pady=(2, 8))

        # 等化子設定（eq 關閉時 gray out）
        frame_eq_sub = tk.Frame(frame_mix)
        frame_eq_sub.pack(fill="x")
        frame_eq_sub.columnconfigure(2, weight=1)

        lbl_cap = ttk.Label(frame_eq_sub, text="增益上限：")
        lbl_cap.grid(row=0, column=0, sticky="w")
        spn_cap = ttk.Spinbox(frame_eq_sub, from_=1, to=16,
                              textvariable=cap_var, width=5)
        spn_cap.grid(row=0, column=1, sticky="w", padx=(4, 0))
        lbl_x = ttk.Label(frame_eq_sub, text="x  (1～16)", foreground="gray")
        lbl_x.grid(row=0, column=2, sticky="w", padx=(4, 0))
        chk_filter = ttk.Checkbutton(
            frame_eq_sub,
            text="靜音過濾：排除靜音段再計算 RMS",
            variable=filter_var,
        )
        chk_filter.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        _eq_sub_widgets = [lbl_cap, spn_cap, lbl_x, chk_filter]

        def _update_eq_state(*_):
            state = "normal" if eq_var.get() else "disabled"
            for w in _eq_sub_widgets:
                w.config(state=state)

        eq_var.trace_add("write", _update_eq_state)
        _update_eq_state()  # 設定初始 gray-out 狀態

        # ---- 確認 / 取消 ----
        frame_btns = tk.Frame(win)
        frame_btns.grid(row=2, column=0, pady=12)

        def confirm():
            self.bit_rate.set(br_var.get())
            self.equalize_enabled.set(eq_var.get())
            self.eq_gain_cap.set(max(1, min(16, cap_var.get())))
            self.eq_filter_silence.set(filter_var.get())
            win.destroy()

        ttk.Button(frame_btns, text="確認", command=confirm).pack(side="left", padx=8)
        ttk.Button(frame_btns, text="取消", command=win.destroy).pack(side="left")
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.columnconfigure(0, weight=1)
```

- [ ] **Step 2：手動測試進階設定**

1. 開啟進階設定 → 看到「音質」和「混音設定」兩個區塊，沒有「輸出方式」
2. 「自動等化音量」未勾選 → 增益上限、靜音過濾 gray out 無法點擊
3. 勾選「自動等化音量」→ 子選項恢復可用
4. 改位元率到 320 → 確認 → 錄一小段 → 停止 → 確認存出的 mp3 是 320kbps（用 MediaInfo 或類似工具驗證）

- [ ] **Step 3：跑現有測試確認沒有回退**

```
venv\Scripts\python.exe -m pytest tests/ -v
```
預期：全部 PASS

- [ ] **Step 4：Commit**

```
git add main.py
git commit -m "feat: 重寫進階設定彈窗（音質 + 混音兩區，等化子選項 gray-out）"
```

---

### Task 5：更新文件並推上 remote

**Files:**
- Modify: `CHANGELOG.md`、`ARCHITECTURE.md`

- [ ] **Step 1：更新 `CHANGELOG.md`**

在現狀總覽的「已完成功能」補：
```
- [x] 主畫面內嵌輸出格式（系統＋麥克風模式動態顯示）
- [x] MP3 位元率可設定（128 / 192 / 320 kbps）
- [x] 模組級 config 常數（混音權重、靜音閾值等）
```

在更新記錄最上方加：
```
### 2026-05-11 — 進階設定重構、主畫面輸出格式、config 常數

- 新增：主畫面「系統＋麥克風」模式下動態顯示輸出格式選項（合併 / 獨立 / 兩個都要）
- 新增：MP3 位元率可在進階設定選擇（128 / 192 / 320 kbps）
- 改善：進階設定重組為「音質」+「混音設定」兩區，移除輸出方式區塊
- 改善：等化子選項（增益上限、靜音過濾）在未開啟自動等化時 gray out
- 重構：硬編碼參數提取為模組頂部 config 常數（方便進階使用者直接修改）
```

- [ ] **Step 2：更新 `ARCHITECTURE.md` 的「關鍵設定變數」表**

在最後補：
```
| `_DEFAULT_BIT_RATE` | 模組頂部 | MP3 位元率預設值，可在進階設定中更改 |
| `_MIX_WEIGHT_SYSTEM / _MIX_WEIGHT_MIC` | 模組頂部 | 混音權重，進階使用者直接改此常數 |
| `_SILENCE_RMS_THRESHOLD` | 模組頂部 | 靜音閾值，影響警告橫幅與等化計算 |
| `_SILENCE_WARNING_SECS` | 模組頂部 | 靜音警告倒計時 |
```

- [ ] **Step 3：Commit 並推上 remote**

```
git add CHANGELOG.md ARCHITECTURE.md
git commit -m "docs: 更新 CHANGELOG / ARCHITECTURE — 進階設定重構"
git push
```
