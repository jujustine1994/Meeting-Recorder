# 儲存處理中關閉視窗警告確認 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 停止/暫停後正在編碼存檔期間，若使用者按下主視窗關閉鈕，跳出確認對話框（是否接受遺失處理進度而強制關閉），避免無提示中斷正在寫入的 MP3 檔案。

**Architecture:** 新增一個明確的旗標 `self._is_saving`，在既有的 `_stop_recording()` / `_reset_ui_after_stop()` 生命週期方法中標記進出「處理中」狀態；在 `__init__` 綁定 `WM_DELETE_WINDOW` 到一個新方法 `_on_close_request`，由它讀取旗標決定是否跳出 `messagebox.askyesno` 確認對話框。

**Tech Stack:** Python 3.13（專案 venv）、tkinter/`tkinter.messagebox`、pytest（既有 `tests/test_audio_processing.py`）。

## Global Constraints

- 觸發範圍僅限「停止/暫停後正在編碼存檔」期間；錄音進行中（尚未按停止）關閉視窗的行為本次不變
- 對話框文字固定為標題「檔案處理中」、內容「檔案處理中，若現在關閉將遺失處理進度，確定要關閉嗎？」
- 使用 `messagebox.askyesno`（是／否兩個按鈕），選「是」→ `self.root.destroy()`；選「否」→ 不做任何事，視窗保持開啟
- 非處理中關閉視窗，行為必須與改動前完全一致（直接 `destroy()`，不跳窗）
- 不做「等待存檔完成後自動關閉」之類的排隊機制（spec 明確排除）
- 此功能是 tkinter 對話框互動，既有 pytest 套件只測純函式、不測 tkinter 元件（既有慣例，見 `tests/test_audio_processing.py`），本次不新增自動化測試；每個 commit 前執行 `venv/Scripts/python.exe -m pytest tests/ -v` 確認既有 25 個測試仍全數通過，並用 `ast.parse` 做語法檢查

---

### Task 1: 新增 `_is_saving` 旗標與視窗關閉攔截

**Files:**
- Modify: `main.py:130`（`__init__` 錄音狀態初始化區）
- Modify: `main.py:177`（`__init__` 尾端，`_build_ui()` 呼叫之後）
- Modify: `main.py:937-954`（`_stop_recording`）
- Modify: `main.py:1562-1583`（`_reset_ui_after_stop`）

**Interfaces:**
- Consumes: 無（不依賴其他任務）
- Produces: `self._is_saving: bool`（其他程式碼可讀取此旗標判斷是否正在存檔）、`self._on_close_request(self) -> None`（`WM_DELETE_WINDOW` callback，無回傳值）

- [ ] **Step 1: 在 `__init__` 新增 `_is_saving` 初始狀態**

把 `main.py:130`：
```python
        self.is_recording = False
```

改為：
```python
        self.is_recording = False
        self._is_saving = False  # True＝停止/暫停後正在編碼存檔，尚未完成
```

- [ ] **Step 2: `_stop_recording` 進入處理中時標記旗標**

把 `main.py:937-940`：
```python
    def _stop_recording(self):
        self.is_recording = False

        # 在主執行緒鎖定模式，避免背景 _save_after_stop 從 tkinter StringVar 讀取
```

改為：
```python
    def _stop_recording(self):
        self.is_recording = False
        self._is_saving = True

        # 在主執行緒鎖定模式，避免背景 _save_after_stop 從 tkinter StringVar 讀取
```

- [ ] **Step 3: `_reset_ui_after_stop` 處理完成時清除旗標**

把 `main.py:1562-1564`：
```python
    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self.is_paused = False
```

改為：
```python
    def _reset_ui_after_stop(self):
        """錄音與儲存流程完全結束後，還原所有 UI 元件狀態"""
        self._is_saving = False
        self.is_paused = False
```

- [ ] **Step 4: 新增 `_on_close_request` 方法**

在 `_reset_ui_after_stop` 方法結尾（`main.py:1583` 的 `self._mic_clip_until = 0.0` 之後，`_update_clip_style` 方法定義之前）插入：
```python

    def _on_close_request(self):
        """
        WM_DELETE_WINDOW callback。存檔處理中關閉視窗會跳出確認，
        避免無提示中斷正在寫入的 MP3 檔案；非處理中則直接關閉，行為不變。
        """
        if self._is_saving:
            proceed = messagebox.askyesno(
                "檔案處理中",
                "檔案處理中，若現在關閉將遺失處理進度，確定要關閉嗎？"
            )
            if not proceed:
                return
        self.root.destroy()
```

- [ ] **Step 5: 在 `__init__` 綁定 `WM_DELETE_WINDOW`**

把 `main.py:177`：
```python
        self._build_ui()
        self._poll_queue()
```

改為：
```python
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_request)
        self._poll_queue()
```

- [ ] **Step 6: 確認 `messagebox` 已匯入**

`main.py` 頂部的 import 已經是：
```python
from tkinter import ttk, filedialog, scrolledtext, messagebox
```
（見 `main.py:20`）`messagebox` 已包含在內，這一步不需要修改，只需確認沒有被移除。

- [ ] **Step 7: 執行既有測試與語法檢查**

Run: `venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 25 個測試全數 PASS（此任務為 tkinter 對話框邏輯，既有測試套件不會涵蓋，符合現行慣例；這裡只是確認沒有把其他東西改壞）

Run: `venv/Scripts/python.exe -c "import ast; ast.parse(open('main.py', encoding='utf-8').read())"`
Expected: 無輸出、無錯誤（語法正確）

- [ ] **Step 8: Commit**

```bash
git add main.py
git commit -m "feat: 存檔處理中關閉主視窗跳出警告確認，避免中斷檔案寫入"
```

- [ ] **Step 9: 人工驗證**

雙擊 `meeting_recorder啟動器.bat` 啟動程式，依序驗證（對照 spec 第五節）：

1. 開始錄音、按下停止，趁狀態列顯示「轉換為 MP3 中...」時按視窗右上角關閉鈕，確認跳出標題「檔案處理中」、內容「檔案處理中，若現在關閉將遺失處理進度，確定要關閉嗎？」的對話框
2. 對話框選「否」，確認視窗維持開啟、狀態列繼續顯示處理中，程式沒有被關閉
3. 再次按關閉鈕，這次選「是」，確認視窗直接關閉、程式結束
4. 重新開啟程式，錄音停止後等到狀態列變成「已儲存：...」（處理完成）再按關閉鈕，確認**不**跳警告、直接關閉
5. 完全還沒開始錄音時按關閉鈕，確認直接關閉，不跳警告
6. 錄音進行中（尚未按停止）按關閉鈕，確認行為跟改動前一樣（直接關閉，不受本次改動影響）

若驗證中發現任何一步不符預期，記錄下實際看到的行為，回報後再修正。
