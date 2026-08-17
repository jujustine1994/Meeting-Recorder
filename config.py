# config.py
"""config.py — 讀寫 config.json。

讀出來的內容一律與 DEFAULT_CONFIG 合併，缺的 key 補上預設值，
呼叫端不必到處寫 `.get(..., 預設)`。

config.json 只存**機器值**（語言代號 zh_tw / en / ...），不存任何顯示文字。
"""

from __future__ import annotations

import copy
import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULT_CONFIG: dict = {
    # 介面語言。代號清單見 i18n.LANGUAGES。
    #
    # 預設是**空字串而不是 "zh_tw"**：空字串代表「使用者還沒選過」，
    # main._pick_language_on_first_run() 靠它決定首次啟動要不要問。填了
    # "zh_tw" 就分不出「他選了繁中」和「他沒選過」，只能再加一個布林值，
    # 而兩個欄位描述同一件事遲早會不同步。
    #
    # 空字串餵給 i18n.set_lang() 會退回預設語言，所以就算問語言那步被跳過，
    # 程式照樣跑得動。
    "language": "",
}


def load_config(path: str | None = None) -> dict:
    """讀 config.json，缺的 key 用預設值補齊。檔案壞掉時整份退回預設。"""
    if path is None:
        path = CONFIG_PATH
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError, ValueError):
            return cfg
        if isinstance(data, dict):
            for key, default_val in DEFAULT_CONFIG.items():
                if key in data and not isinstance(default_val, dict):
                    cfg[key] = data[key]
    return cfg


def save_config(cfg: dict, path: str | None = None) -> None:
    """寫回 config.json（UTF-8）。寫不進去不讓主程式掛掉，只是設定沒存成。"""
    if path is None:
        path = CONFIG_PATH
    try:
        parent = os.path.dirname(path)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass
