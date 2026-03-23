# CHANGELOG — Meeting Recorder

## 現狀總覽（2026-03-23）

### 已完成功能
- [x] WASAPI Loopback 系統音訊錄製
- [x] lameenc MP3 編碼（無需 ffmpeg）
- [x] tkinter 資料夾選擇視窗
- [x] 錄音計時器顯示
- [x] 時間戳記自動命名
- [x] 連續錄音（不重啟）
- [x] 自動安裝啟動器（launcher.ps1）
- [x] 三種錄音模式（電腦聲音 / 麥克風 / 兩者混音）
- [x] 靜音警告橫幅（連續 10 秒 RMS < 100 時顯示）
- [x] 同名自動流水號避免覆蓋

### 未完成 / 待規劃
- [ ] 錄音中顯示音量波形（vu meter）
- [ ] 支援選擇錄音裝置（多音效卡環境）
- [ ] 錄音完成後自動開啟資料夾

---

## 更新記錄

### 2026-03-24
- 新增：專案初始建立，完整錄音功能上線
- 修改：主程式改為 tkinter GUI 視窗介面（參考 SnapTranscript 架構）
- 新增：自訂檔案名稱欄位，錄音前可先填好名稱
- 修正：launcher.ps1 加入 ARM64 架構偵測，強制安裝 x64 Python 確保 pyaudiowpatch 相容性
- 新增：三種錄音模式（電腦聲音 / 麥克風 / 兩者混音）
- 新增：Mode "both" 混音邏輯（loopback stereo + mic mono → upmix → 0.6 權重混音）
- 修正：靜音偵測依模式切換偵測對象（loopback 或麥克風）

### 2026-03-23 — 程式碼全面審查與 Bug 修正
- 修正：Mode "both" 麥克風無資料時改存純系統音訊並顯示警告，避免產生無聲 MP3
- 修正：Mode "both" 麥克風取樣率與 loopback 不一致時顯示警告（混音仍繼續）
- 修正：loopback 裝置持續不可用時加入 sleep(1) 避免 CPU busy-wait
- 修正：tkinter StringVar race condition — record_mode 改在主執行緒停止時鎖定為 _save_mode
- 修正：Mode "both" 麥克風啟動失敗改為 warning 而非 error，不中止儲存流程
- 改善：新增 warning 訊息類型至 _poll_queue，顯示在 log 但不中斷流程
- 改善：全程式補齊關鍵邏輯的中文註解（混音權重、靜音閾值、stream closure 設計等）
- 改善：_compute_rms() 提取為獨立函式，消除兩個 worker 的重複靜音偵測程式碼
