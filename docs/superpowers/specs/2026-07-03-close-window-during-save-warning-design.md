# 設計文件：儲存處理中關閉視窗跳出警告確認

**日期**：2026-07-03
**狀態**：已確認，待實作

---

## 背景與問題

錄音停止或暫停後，`_stop_recording()` 會啟動背景執行緒 `_save_after_stop` 進行 MP3 編碼與存檔。這段期間主視窗完全沒有攔截 `WM_DELETE_WINDOW`，使用者若在此時按下視窗右上角的關閉鈕，視窗會直接關閉、程式結束，可能導致正在寫入的 MP3 檔案損毀或遺失。

範圍限定：本次只處理「停止/暫停後正在編碼存檔」這段期間。錄音進行中（尚未按下停止）時關閉視窗的行為維持現狀不變，不在本次範圍內。

---

## 一、狀態追蹤

新增明確旗標 `self._is_saving: bool`，取代用按鈕文字判斷是否在處理中：

- `_stop_recording()` 方法開頭設為 `True`
- `_reset_ui_after_stop()` 方法（無論存檔成功 `saved`、失敗 `error`、或捨棄 `discarded` 都會呼叫到這裡）設回 `False`

`__init__` 中初始化 `self._is_saving = False`。

---

## 二、視窗關閉攔截

### 綁定
在 `_build_ui` 建立主視窗元件完成後，加入：
```python
self.root.protocol("WM_DELETE_WINDOW", self._on_close_request)
```

### 新增方法 `_on_close_request`
```python
def _on_close_request(self):
    if self._is_saving:
        proceed = messagebox.askyesno(
            "檔案處理中",
            "檔案處理中，若現在關閉將遺失處理進度，確定要關閉嗎？"
        )
        if not proceed:
            return
    self.root.destroy()
```

- `self._is_saving == True` 時：跳出確認對話框（是／否）
  - 選「是」→ 使用者接受風險，執行 `self.root.destroy()` 強制關閉（存檔可能中斷，這是使用者主動選擇的結果，不额外攔阻）
  - 選「否」→ 直接 `return`，不做任何事，視窗保持開啟
- `self._is_saving == False` 時：跳過確認，直接 `self.root.destroy()`，與現行行為一致

---

## 三、不在此次範圍

- 錄音進行中（`is_recording == True` 且尚未按下停止）時關閉視窗的攔截，本次不處理
- 不提供「等待存檔完成後自動關閉」之類的排隊關閉機制
- 不記錄或持久化使用者選擇（每次關閉都重新詢問）

---

## 四、修改影響範圍

| 項目 | 變更 |
|------|------|
| `__init__` | 新增 `self._is_saving = False` |
| `_stop_recording` | 開頭加入 `self._is_saving = True` |
| `_reset_ui_after_stop` | 加入 `self._is_saving = False` |
| `_build_ui` | 加入 `self.root.protocol("WM_DELETE_WINDOW", self._on_close_request)` |
| 新增方法 | `_on_close_request` |

---

## 五、測試方式

- 開始錄音、按下停止，趁「轉換為 MP3 中...」狀態時按視窗關閉鈕，確認跳出確認對話框
- 對話框選「否」，確認視窗維持開啟、程式繼續處理
- 對話框選「是」，確認視窗直接關閉、程式結束
- 存檔完成（狀態變回「已儲存：...」）後再按關閉鈕，確認不跳警告、直接關閉
- 尚未開始錄音、或錄音進行中尚未按停止時按關閉鈕，確認行為與改動前一致（直接關閉，不受影響）
