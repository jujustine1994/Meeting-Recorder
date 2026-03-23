# TODO — Meeting Recorder

## 待辦

- [ ] 測試首次安裝流程（乾淨環境）
- [ ] 測試各種採樣率（44100 / 48000 Hz）
- [ ] 測試長時間錄音（60 分鐘以上）
- [ ] 確認 ARM64 電腦相容性（pyaudiowpatch 支援情況）

## 未來功能（有需要再做）

- [ ] **MacBook Windows VM 相容性**：WASAPI Loopback 在 VM 環境（Parallels / VMware / VirtualBox）無法使用，因虛擬音效卡驅動不支援。解法方向：偵測無 Loopback 裝置時顯示友善錯誤，引導使用者安裝 VB-Cable 虛擬音訊線作為替代方案。

- [ ] 錄音中顯示音量 vu meter
- [ ] 支援選擇特定音效裝置
- [ ] 錄完自動開啟 Explorer 到儲存位置
- [ ] 設定檔（記住上次的儲存位置）
