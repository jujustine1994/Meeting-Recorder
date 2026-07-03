# 設計文件：VU Meter 靈敏度調整 + 爆音警示

**日期**：2026-07-03
**狀態**：已確認，待實作
**前情**：延續 `2026-05-05-vu-meter-and-track-output-design.md` 已實作的 VU Meter，本次只調整換算公式與新增爆音警示，不改動 UI 位置與既有訊息機制。

---

## 背景與問題

現行 VU meter 換算公式為 `min(100, rms / 327.67)`，把 RMS 線性對應到 Int16 滿刻度（32767）的百分比。實測正常說話音量只落在 1~10%，使用者難以從畫面判斷「有沒有在收音」，也沒有任何機制提示音量過大或爆音。

---

## 一、換算公式調整

### 除數變更
`_compute_rms` 保持不變，僅調整顯示端換算：

```python
VU_DISPLAY_DIVISOR = 80   # 原 327.67
level = min(100, rms / VU_DISPLAY_DIVISOR)
```

- 效果：安靜背景／氣音約 5~15%，正常說話音量落在 20~50%，只有異常大聲才接近頂端
- 常數集中定義（模組層級，緊鄰 `_SILENCE_RMS_THRESHOLD`），方便日後微調

### 套用範圍
以下四處全部改用 `VU_DISPLAY_DIVISOR`，不再各自寫死 `327.67`：

| 位置 | 現行程式碼（main.py 行號，實作前應重新確認） |
|------|------|
| 正式錄音 - 系統音訊 | 935 行附近 `vu_system` 訊息 |
| 正式錄音 - 麥克風 | 1106 行附近 `vu_mic` 訊息 |
| 裝置測試 - 麥克風測試 | 477 行附近 |
| 裝置測試 - 系統音訊測試 | 563 行附近 |

四處統一，確保測試畫面看到的百分比跟正式錄音時一致，不會讓使用者混淆「測試時正常、錄音時偏低」。

---

## 二、爆音警示（新增）

### 峰值計算
新增輔助函式，回傳區塊內樣本最大絕對值（與 RMS 分開，不影響現有靜音偵測邏輯）：

```python
def _compute_peak(data: bytes) -> int:
    """回傳 PCM Int16 資料的峰值（樣本最大絕對值），範圍 0~32767。"""
    num_samples = len(data) // 2
    if num_samples == 0:
        return 0
    samples = struct.unpack(f"{num_samples}h", data)
    return max(abs(s) for s in samples)
```

### 判定與訊息
```python
CLIP_PEAK_THRESHOLD = int(32767 * 0.9)   # 約 29491
CLIP_WARNING_HOLD_SECS = 1.5
```

- 系統音訊與麥克風的錄音 worker（`_record_worker` / `_record_mic_worker`）在既有的 RMS 計算旁，額外計算 peak
- peak 超過閾值時，透過 `msg_queue` 送出新訊息：`("vu_system_clip", True)` / `("vu_mic_clip", True)`
- 未超過閾值時不用每次都送 False；由主執行緒用「最後一次收到 True 的時間」搭配 `CLIP_WARNING_HOLD_SECS` 做衰減，超過保持時間才恢復原色，避免瞬間峰值一閃即逝看不到
- 裝置測試畫面（麥克風測試、系統音訊測試）的兩個背景執行緒比照辦理，各自送出對應 clip 訊息

### UI 呈現
- `ttk.Progressbar` 預設樣式不變；新增一個警示樣式（例如 `Clip.Horizontal.TProgressbar`，前景色改紅），用 `ttk.Style().configure(...)` 定義一次
- 主執行緒的 `_poll_queue`（或裝置測試視窗對應的 UI 更新邏輯）收到 clip 訊息或判斷仍在 hold 期間時，把對應 Progressbar 的 `style` 換成警示樣式，並將旁邊的百分比 `Label` 文字顏色一併改紅；hold 到期後兩者都還原
- 系統音訊與麥克風分開判斷、分開顯示，互不影響

### 不影響範圍
- 靜音偵測（10 秒警告橫幅）沿用現有 RMS 閾值機制，不受本次調整影響
- 錄音存檔的原始 PCM 資料不受影響，這次只改「畫面顯示」的換算方式與新增視覺警示

---

## 三、不在此次範圍

- 音量顯示改成 dB 對數刻度（本次維持線性，只調整除數）
- 自動降低錄音增益（AGC）以避免爆音，本次僅做警示，不介入實際訊號
- 除數、峰值閾值、hold 秒數的使用者可調整介面（本次先寫死常數，之後如有需要再開放到「進階設定」）

---

## 四、修改影響範圍

| 項目 | 變更 |
|------|------|
| 模組層級常數 | 新增 `VU_DISPLAY_DIVISOR`、`CLIP_PEAK_THRESHOLD`、`CLIP_WARNING_HOLD_SECS` |
| 新增函式 | `_compute_peak` |
| `_record_worker` | RMS 換算改用新除數；新增 peak 計算與 `vu_system_clip` 訊息 |
| `_record_mic_worker` | RMS 換算改用新除數；新增 peak 計算與 `vu_mic_clip` 訊息 |
| 裝置測試 - 麥克風測試背景執行緒 | RMS 換算改用新除數；新增 peak 計算與 clip 訊息 |
| 裝置測試 - 系統音訊測試背景執行緒 | RMS 換算改用新除數；新增 peak 計算與 clip 訊息 |
| `_poll_queue` | 新增 `vu_system_clip` / `vu_mic_clip` 處理，含 hold 衰減邏輯與樣式切換 |
| 裝置測試視窗的訊息處理邏輯 | 同上，新增 clip 處理 |
| `ttk.Style` | 新增一個紅色警示用的 Progressbar 樣式 |

---

## 五、測試方式

- 正常說話音量錄音，確認指示條落在約 20~50% 區間跳動
- 對著麥克風大聲說話或製造爆音，確認指示條與百分比文字變紅，停止大聲後約 1.5 秒恢復原色
- 裝置測試畫面比照上述兩項，確認跟正式錄音畫面數值/行為一致
- 完全靜音時，確認仍維持原有的 10 秒靜音警告橫幅邏輯不受影響
