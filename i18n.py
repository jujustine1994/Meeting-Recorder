# i18n.py
"""i18n.py — 介面顯示文字的多語言查表。

用法：
    from i18n import t
    ttk.Button(text=t("gui.btn.record_start"))

啟動時由 main.py 呼叫一次 `set_lang(cfg["language"])`，之後全程式共用。

## 設計約束（改這個檔前先讀）

1. **t() 永不 raise、永不回空字串。** 查找順序是
   `目標語言 → 繁體中文 → key 本身`。最壞情況畫面顯示 `gui.btn.record_start`
   這串 key，一眼看得出哪裡漏翻；回空字串會變成看不見的按鈕，那才是災難。

2. **機器鍵不進這裡。** 會被寫進檔案、或拿去跟檔案裡的值比對的字串是
   **資料**不是介面文字（例如錄音模式代號 system / mic / both、輸出模式
   merge / separate / both、msg_queue 的訊息類型、存檔用的 _system / _mic
   檔名後綴）。這些一律維持英文字面，不進本表。

3. **logs/app.log 不吃這裡的翻譯**，永遠繁中。log 是給維護者除錯用的，
   跟著使用者語言變等於自廢。log 的字串住在 `logtext.py`。

## 新增語言只要兩步

1. 複製一份 `locales/en.py` 改譯文
2. 在下面 LANGUAGES 加一行

語言下拉選單、首次啟動選單、tests/test_i18n.py 的漏 key 檢查全部自動涵蓋。
"""

from __future__ import annotations

import importlib

# (代號, 下拉選單顯示名, 字型)
#
# 代號     存進 config.json 的值，也是 locales/<代號>.py 的檔名
# 顯示名   用各語言自己的說法，任何語言下使用者都認得出哪個是哪個
# 字型     微軟正黑體缺日文假名字形，日文要 Yu Gothic；兩者皆 Windows 內建
LANGUAGES: list[tuple[str, str, str]] = [
    ("zh_tw", "繁體中文", "微軟正黑體"),
    ("zh_cn", "简体中文", "Microsoft YaHei"),
    ("en",    "English",  "Calibri"),
    ("ja",    "日本語",   "Yu Gothic"),
]

DEFAULT_LANG = "zh_tw"

# 找不到 key 時的最終退路。母語言是其他語言的翻譯來源，所以它一定最完整。
FALLBACK_LANG = "zh_tw"

_LANG_CODES = [code for code, _, _ in LANGUAGES]
_current_lang: str = DEFAULT_LANG
_cache: dict[str, dict[str, str]] = {}


def _strings(lang: str) -> dict[str, str]:
    """載入某語言的字串表。載入失敗回空 dict，讓 t() 自己退回 fallback。"""
    if lang in _cache:
        return _cache[lang]
    try:
        mod = importlib.import_module(f"locales.{lang}")
        table = getattr(mod, "STRINGS", {})
    except (ImportError, AttributeError):
        # 語言檔缺失或壞掉不能讓整個程式起不來
        table = {}
    _cache[lang] = table
    return table


def available_languages() -> list[tuple[str, str]]:
    """給選單用：[(代號, 顯示名), ...]，順序即選單順序。"""
    return [(code, name) for code, name, _ in LANGUAGES]


def is_supported(lang: str) -> bool:
    return lang in _LANG_CODES


def set_lang(lang: str | None) -> str:
    """設定目前語言。不認得的代號（含 None、舊 config 的怪值）退回預設。

    回傳實際採用的代號——要顯示「現在是什麼語言」時用回傳值，不要用傳進去
    的參數，兩者在退回時不同。
    """
    global _current_lang
    _current_lang = lang if is_supported(lang or "") else DEFAULT_LANG
    return _current_lang


def get_lang() -> str:
    return _current_lang


def ui_font(lang: str | None = None) -> str:
    """該語言的字型名稱。

    本工具目前**不指定字型**（Windows 的 tkinter 預設 Segoe UI 四種語言都顯示
    得出來；指定字型會改變繁中的既有外觀）。保留這個函式是給日後真的量到日文
    出現豆腐、或要輸出報表時用——需要時直接呼叫，不要在模組層級求值。
    """
    target = lang if is_supported(lang or "") else _current_lang
    for code, _, font in LANGUAGES:
        if code == target:
            return font
    return LANGUAGES[0][2]


def t(key: str, **fmt) -> str:
    """查表。目標語言 → 母語言 → key 本身。

    **fmt 走 str.format，給帶變數的訊息用：
        t("gui.status.saved", name="meeting.mp3")

    格式化失敗（譯文的 placeholder 打錯）不 raise，回未格式化的原字串——
    畫面上看到 {name} 殘留，比整個程式當掉好處理。
    """
    s = _strings(_current_lang).get(key)
    if s is None:
        s = _strings(FALLBACK_LANG).get(key)
    if s is None:
        return key
    if not fmt:
        return s
    try:
        return s.format(**fmt)
    except (KeyError, IndexError, ValueError):
        return s
