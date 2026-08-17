# logtext.py
"""logs/app.log 的訊息字串——**永遠繁體中文，不跟使用者介面語言走**。

log 是給維護者除錯用的：跟著使用者語言變，等於自己看不懂自己的 log。
所以這些字串刻意不進 i18n，也刻意集中在這個檔——main.py 才能被
tests/test_i18n.py 的「不得寫死中日文」那條測試完整涵蓋（本工具的 GUI
全部住在 main.py，把 main.py 丟進豁免清單等於直接把那條測試關掉），
本檔則是清單裡少數幾個有理由的豁免。

用法：
    from logtext import LOG_TEXT
    _write_log_header(LOG_TEXT["record_start"].format(mode="both", ...))

格式一律具名 placeholder（`{mode}` 不是 `{0}`）。
本檔只放常數，不放邏輯。
"""

from __future__ import annotations

LOG_TEXT: dict[str, str] = {
    # 任務起始（=== 行，唯一有完整日期的行）
    "record_start":       "錄音 模式:{mode} | 輸出:{output} | {bitrate}kbps",

    # 任務結果
    "result_ok":          "成功，耗時 {minutes}分{seconds}秒 | 存檔 {count} 個",
    "result_fail":        "失敗，耗時 {minutes}分{seconds}秒",
    "result_discarded":   "使用者捨棄，耗時 {minutes}分{seconds}秒",

    # 錯誤行：只記 exception 類型 / 逾時 / 重試次數，不記例外訊息本體
    # （pyaudio、lameenc 的例外訊息可能挾帶裝置路徑等資訊）
    "sys_device_timeout": "系統音訊裝置中斷 -> 30秒逾時未恢復 | 重試 30/30",
    "mic_device_timeout": "麥克風裝置中斷 -> 30秒逾時未恢復 | 重試 30/30",
    "thread_error_sys":   "錄音執行緒(system) -> {exc}",
    "thread_error_mic":   "錄音執行緒(mic) -> {exc}",
    "save_error":         "儲存處理 -> {exc}",
}
