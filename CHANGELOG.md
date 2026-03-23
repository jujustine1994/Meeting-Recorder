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

### 2026-03-23
- 新增：專案初始建立，完整錄音功能上線
- 修改：主程式改為 tkinter GUI 視窗介面（參考 SnapTranscript 架構）
