# 設計文件：VU Meter + 獨立音軌輸出 + 自動等化

**日期**：2026-05-05  
**狀態**：已確認，待實作

---

## 功能概要

新增三個互相關聯的功能：

1. **VU Meter**：錄音時即時顯示系統音訊與麥克風的音量指示條
2. **獨立音軌輸出**：「系統 + 麥克風」模式可選擇合併、獨立或兩者都存
3. **自動等化音量**：混音前自動拉齊兩軌響度，附進階參數調整

---

## 一、VU Meter

### UI 位置
插入主視窗按鈕區，位於 `status_label` 下方（`frame_btn` 內）：

```
[⏺ 開始錄音]
  00:00
  等待開始錄音...

  系統音訊  ████████████░░░░░░░░  58%
  麥克風    ████░░░░░░░░░░░░░░░░  22%

[🔧 裝置設定與測試]  [⚙ 進階設定]
```

### 行為
- 兩條 `ttk.Progressbar` 永遠可見，不錄音時數值為 0
- 錄音中每個 chunk 計算 RMS，透過 `msg_queue` 送到主執行緒更新
- 新增兩種訊息類型：`("vu_system", float)` / `("vu_mic", float)`
- `_poll_queue` 收到後更新對應 `tk.DoubleVar`（最大值 100）
- RMS 換算沿用現有公式：`min(100, rms / 327.67)`

### 變數新增
```python
self.vu_system_var = tk.DoubleVar(value=0)
self.vu_mic_var    = tk.DoubleVar(value=0)
```

---

## 二、⚙ 進階設定

### 按鈕位置
「裝置設定與測試」按鈕右側並排，新增「⚙ 進階設定」按鈕。

### 彈窗結構

**一般區**：
```
輸出方式（僅「系統+麥克風」模式有效）
  ○ 合併一軌（預設）
  ○ 獨立兩軌
  ○ 兩個都要

☐ 自動等化音量
   讓麥克風與系統音訊響度接近
   （僅影響含混音的輸出；獨立音軌存原始音量）
```

**進階調整區（LabelFrame）**：
```
┌ 等化進階設定 ─────────────────┐
│ 增益上限：[4] x  (1～16)      │
│ 靜音過濾：☑ 排除靜音段再計算   │
└───────────────────────────────┘
```

### 實例變數新增
```python
self.output_mode      = tk.StringVar(value="merge")   # merge / separate / both
self.equalize_enabled = tk.BooleanVar(value=False)
self.eq_gain_cap      = tk.IntVar(value=4)            # 倍數上限，1～16
self.eq_filter_silence = tk.BooleanVar(value=True)    # 過濾靜音段
```

設定存於 app 實例，本版不做持久化。

---

## 三、獨立音軌輸出

### 觸發條件
`self._save_mode == "both"` 且 `self.output_mode.get() != "merge"`

### 檔名規則
| 輸出方式 | 存檔 |
|---------|------|
| `merge`    | `{name}.mp3` |
| `separate` | `{name}_system.mp3` + `{name}_mic.mp3` |
| `both`     | `{name}.mp3` + `{name}_system.mp3` + `{name}_mic.mp3` |

- 獨立音軌永遠存**原始未處理**的 PCM，等化不套用
- 同名防衝突邏輯（流水號）沿用現有實作，每個檔案獨立判斷

### `_save_after_stop` 修改點
"both" 分支：
1. 依 `output_mode` 決定要存哪些檔案
2. 需要 merged 時：先做等化（若開啟），再呼叫 `_mix_pcm`，存 `{name}.mp3`
3. 需要 separate 時：直接從 `record_frames` / `mic_frames` 各自編碼，存兩個獨立 MP3

---

## 四、自動等化音量

### 演算法（`_compute_equalize_gain` 新增函式）

```python
def _compute_equalize_gain(sys_frames, mic_frames,
                            filter_silence=True, gain_cap=4):
    """
    回傳 (sys_gain, mic_gain)，只有較小的那軌會 > 1.0。
    filter_silence=True 時排除 RMS < SILENCE_RMS_THRESHOLD 的 chunk。
    """
    def active_rms(frames):
        values = [_compute_rms(c) for c in frames]
        if filter_silence:
            values = [v for v in values if v >= SILENCE_RMS_THRESHOLD]
        return (sum(values) / len(values)) if values else 0.0

    sys_rms = active_rms(sys_frames)
    mic_rms = active_rms(mic_frames)

    if sys_rms == 0 or mic_rms == 0:
        return 1.0, 1.0   # 其中一軌全靜音，不等化

    if sys_rms >= mic_rms:
        return 1.0, min(sys_rms / mic_rms, gain_cap)
    else:
        return min(mic_rms / sys_rms, gain_cap), 1.0
```

### 套用時機
- 僅在 `equalize_enabled == True` 且輸出含 merged 時執行
- 對原始 bytes 按增益係數縮放（Int16 clamp 至 ±32767），得到等化後的 bytes，再傳入 `_mix_pcm`
- 獨立音軌（separate）不套用等化，永遠存原始音量

### 限制（已知，可接受）
- 採用全段靜音過濾後平均 RMS，對非常不規律的講話節奏可能仍有誤差
- 增益上限（預設 4x / +12 dB）防止靜音軌被過度放大
- 若兩軌差距超過上限，混音後仍可能有響度差異

---

## 五、不在此次範圍

- 設定持久化（記住進階設定）
- 實時 AGC（錄音中動態調整增益）
- 重採樣（取樣率不一致的根本解法）
- WAV / FLAC 輸出格式

---

## 六、修改影響範圍

| 項目 | 變更 |
|------|------|
| `_build_ui` | 新增 VU Meter 兩條 progressbar + ⚙ 按鈕 |
| `_record_worker` | 新增 `vu_system` 訊息 |
| `_record_mic_worker` | 新增 `vu_mic` 訊息 |
| `_poll_queue` | 新增 `vu_system` / `vu_mic` 處理 |
| `_save_after_stop` | 拆分 "both" 分支，支援 separate / both 輸出 |
| `_mix_pcm` | 不改介面，等化在呼叫前處理 |
| 新增函式 | `_compute_equalize_gain` / `_apply_gain_to_pcm` / `_show_advanced_settings` |
| `__init__` | 新增 4 個實例變數 |
