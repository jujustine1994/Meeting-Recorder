# TODO — Meeting Recorder

## 文件維護
- [ ] 校正專案 MD（依新模板：ARCHITECTURE 補現狀，CHANGELOG 拿掉現狀段）

## 待測試（2026-05-13 本次改動）

- [ ] **MME 麥克風錄音**：Discord 開著的狀態下錄 30 秒，聽 `_mic.mp3` 聲音是否正常（不再有機器人聲或 AGC pumping）
- [ ] **裝置測試一致性**：Discord 開著，進「裝置設定與測試」→ 麥克風測試，VU meter 應正常動，不失真
- [ ] **存檔進度條**：錄 5 分鐘以上，用 `both/both` 模式停止，觀察 log 是否即時顯示 `⏳ 編碼 system (1/3) X%` 並更新，最後出現 `✓ meeting_XXXX.mp3`
- [ ] **裝置中斷重試**：錄音中拔插耳機，log 應出現「系統音訊裝置中斷，嘗試重新連線...」然後「已重新連線，錄音繼續」（或 30 秒後 error）

## 待測試（2026-07-23 VU meter 除數再調整）

- [ ] **VU meter 除數再確認**：實測回報 650 過鬆——小聲趨近 0%、大聲僅約 10%，回推 RMS ≈ 65。已依比例調回 `_VU_DISPLAY_DIVISOR = 120`（目標大聲落在 50~60%），但這也是估算值，尚未實測驗證，之後測過有問題再回報
- [ ] **爆音警示閾值一併確認**：`_CLIP_PEAK_THRESHOLD` 目前是 96%（原本 90% 太容易誤判），跟上面除數一起測，確認正常講話不會誤觸發紅色警示

## 已知風險（非緊急）

- 超過 1 小時的錄音，停止時記憶體峰值約 1.3GB（PCM join 產生暫存副本）。電腦記憶體夠用就沒問題，極端情況可能 crash。

## i18n 遷移紀錄（2026-08-15，`feat/i18n` 分支）

> 判斷理由與待校對項留在這裡，日後改東西時對照用。現況：43 條測試全綠、四語各 108 條 key。

### 判成「資料」不翻的（已有測試釘住，這裡只記理由）

- **錄音模式代號 `system` / `mic` / `both`**（`main.py` 第 974、1056、1364、1380、1567、1575 行）：
  全都是 `mode == "mic"` 這種相等比對在決定走哪條 code path，翻了整個錄音流程會走錯分支。
  同時也是 `LOG_TEXT["record_start"]` 的 `{mode}` 值（第 1029–1030 行）。
- **輸出方式代號 `merge` / `separate` / `both`**（`main.py` 第 203、224、1601–1611、1630 行）：
  同上，`output_mode in ("separate", "both")` 決定編幾軌、存幾個檔。
- **檔名後綴 `_system` / `_mic`**（`main.py` 第 1626–1627 行）與副檔名 `.mp3`
  （第 311、1502、1505 行）、同名流水號樣板 `{name} ({counter}).mp3`。
  檔名跟著介面語言變，同一批錄音會在不同語言下產出不同名字。
- **`logs/app.log` 全部內容**：集中在 `logtext.py`，**固定繁體中文，不跟介面語言走**。
  理由寫在該檔 docstring：log 是給維護者除錯用的，跟著使用者語言變等於自己看不懂自己的 log。
  錯誤行刻意只記 `type(e).__name__`（`main.py` 第 1195、1381、1675 行），
  不記例外訊息本體——pyaudio / lameenc 的例外會挾帶裝置路徑。

### 偏離 `pattern_i18n.py` 的地方

- **`tests/test_i18n.py` 的 `ALLOWLIST`（第 31–34 行）只放 `i18n.py` 與 `logtext.py`，
  `main.py` 刻意不在裡面**。本工具的 GUI 全部住在 `main.py`（8 萬多字元），
  把它豁免掉等於「不得寫死中日文」那條測試整條失效。
  log 字串因此才被抽到 `logtext.py`——那是唯一有正當理由的豁免。
  **日後不管 `main.py` 多難處理，都不要把它加進 ALLOWLIST。**
- **`i18n.ui_font()` 建了但全專案沒有呼叫端**（`i18n.py` 第 96 行）。
  跟另外三個同批專案一致：一接字型，繁中介面外觀就跟遷移前不一樣，
  無法用「畫面長得一模一樣」驗證遷移沒改壞東西。留著給日後用。
- **沒有 `_log()` 的 `log_msg` / `file_msg` 雙訊息參數**。
  `_log()`（第 1679 行）維持原本的 `(msg, level, to_file)` 簽名，
  落檔一律另外直接呼叫 `_write_log(LOG_TEXT[...])`（第 1029、1182、1195… 行），
  也就是「UI 訊息」與「log 訊息」是兩個獨立呼叫點，不是同一個呼叫的兩個參數。
  ⚠ 這跟 Video-Combiner（加了 `log_msg` 參數）和 TW Earning Slides Downloader
  （加了 `file_msg` 參數）都不一樣，三個專案三種做法，日後要統一時記得這支的形狀不同。

### 既有測試被改動：`tests/test_audio_processing.py`（11 行）

這是跟著 `be0fe05`（VU meter 除數 650→120）走的，**與 i18n 無關**，改了兩條：

1. `test_rms_scales_with_divisor`：原本寫死
   `_rms_to_display_pct(13000.0) == 20.0` / `(32500.0) == 50.0`，
   改成綁常數 `_rms_to_display_pct(_VU_DISPLAY_DIVISOR * 20.0) == 20.0`（50 同理）。
   **理由**：除數是會被實測回報反覆調整的參數（歷史上 80→130→650→120），
   寫死就是每調一次紅兩條測試，而那兩條紅燈其實沒抓到任何 bug。
2. `test_theoretical_max_rms_does_not_clamp_at_current_divisor`
   **改名為** `test_theoretical_max_rms_clamps_to_100`，斷言也反過來。
   **理由**：除數 650 時 Int16 理論最大 RMS（32767）換算約 50%、不會撞 clamp；
   改成 120 之後遠超 100%，一定會 clamp。測試名稱與斷言都必須跟著翻面，
   否則是在斷言一件已經不成立的事。
   ⚠ 這條測試的名字綁著當下的除數值，**除數再調時要再檢查一次這條**。

### 譯文待校對

`locales/zh_cn.py`、`locales/en.py`、`locales/ja.py` 的譯文是 AI 產出的，**沒有母語者校對過**。
改譯文**不影響任何邏輯**（程式一律用 key 比對），改錯最壞只是畫面顯示怪。
校對時**只改 value、不要動 key**，`{minutes}` / `{count}` / `{mode}` 這類具名 placeholder
必須保留（`tests/test_i18n.py` 的 `test_placeholders_match_across_languages` 會擋）。

比較沒把握的幾條：

| key | zh_tw | zh_cn | en | ja |
|---|---|---|---|---|
| `gui.mode.system` | `電腦聲音` | `电脑声音` | `Computer audio` | `PC の音声` |
| `gui.mode.both` | `系統 + 麥克風` | `系统 + 麦克风` | `System + Mic` | `システム + マイク` |
| `gui.output.merge` | `合併一軌` | `合并一轨` | `One merged track` | `ミックス 1 トラック` |
| `gui.output.separate` | `獨立兩軌` | `独立两轨` | `Two separate tracks` | `個別 2 トラック` |
| `gui.btn.discard` | `🗑  停止不儲存` | `🗑  停止不保存` | `🗑  Stop, discard` | `🗑  停止して破棄` |
| `gui.lbl.vu_system` | `系統音訊` | `系统音频` | `System` | `システム` |

重點：
- **「軌」的譯法不一致**：英文用 `track`、日文一個用「トラック」一個也用「トラック」但前綴
  「ミックス」vs「個別」，中日文都可以，但英日的 `One merged track` / `ミックス 1 トラック`
  是不是錄音軟體圈的慣用說法沒有把握。
- **`gui.lbl.vu_system` 英文只寫 `System`**（不是 `System audio`），
  因為那是 VU meter 旁的短標籤，欄位很窄；繁中是「系統音訊」四個字。長度不一致是刻意的。
- **`gui.mode.system` 日文用「PC の音声」**——中英都說「電腦聲音 / Computer audio」，
  日文直譯「コンピューターの音声」太長，縮成 `PC の`，不確定自然度。
- **emoji（`⏺` `⏸` `▶` `🗑` `⏹`）四語都保留**，只翻後面的文字，維持按鈕寬度一致。

## 待規劃功能

- [ ] 支援選擇錄音裝置（多音效卡環境）
- [ ] 錄音完成後自動開啟資料夾
- [ ] 存檔處理中（暫停/停止後正在編碼 MP3）若按下關閉視窗，應跳出警告阻擋，不可直接關閉；目前主視窗沒有攔截 `WM_DELETE_WINDOW`，關閉會直接中斷處理中的檔案
