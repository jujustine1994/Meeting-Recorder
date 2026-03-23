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
