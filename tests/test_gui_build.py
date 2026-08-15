# tests/test_gui_build.py
"""四種語言各建置一次完整 GUI，確認沒有殘留的 key 字串。

畫面上出現 `gui.btn.record_start` 就是那條漏翻——t() 查不到時回 key 本身，
所以漏翻不會 crash，只會靜靜地顯示一串程式碼。這個測試就是那雙眼睛。

用 withdraw()，不進 mainloop。
"""

import re
import tkinter as tk
from tkinter import ttk

import pytest

import i18n
import main

LANGS = [code for code, _, _ in i18n.LANGUAGES]

# 殘留 key 長這樣：小寫英數，中間有點。gui.btn.pause / err.no_wasapi
KEY_LIKE = re.compile(r"^[a-z][a-z0-9_]*(\.[a-z0-9_]+)+$")


def _collect(widget, acc):
    for opt in ("text",):
        try:
            v = widget.cget(opt)
        except Exception:
            continue
        # ⚠ Entry / Spinbox 的 cget("text") 會回 PY_VAR0 這種變數名，
        #   是雜訊不是漏翻，要排除
        if isinstance(v, str) and v and not v.startswith("PY_VAR"):
            acc.append(v)
    # Combobox 的 values 與 Treeview 的 heading 用 cget("text") 拿不到
    try:
        if isinstance(widget, ttk.Combobox):
            acc.extend(v for v in widget.cget("values") if isinstance(v, str) and v)
    except Exception:
        pass
    try:
        if isinstance(widget, ttk.Treeview):
            for col in widget.cget("columns"):
                acc.append(widget.heading(col).get("text", ""))
    except Exception:
        pass
    for child in widget.winfo_children():
        _collect(child, acc)


def _build_everything(tk_root, lang):
    """在指定語言下建主視窗與三個 popup，回傳所有可見文字。"""
    i18n.set_lang(lang)
    win = tk.Toplevel(tk_root)
    win.withdraw()
    app = main.MeetingRecorderApp(win)

    texts = []
    _collect(win, texts)
    texts.append(win.title())

    for open_popup in (app._show_mode_help, app._show_device_test,
                       app._show_advanced_settings):
        before = {str(w) for w in win.winfo_children()}
        open_popup()
        for w in win.winfo_children():
            if str(w) in before:
                continue
            try:
                w.grab_release()
            except tk.TclError:
                pass
            texts.append(w.title())
            _collect(w, texts)
            w.destroy()

    win.destroy()
    return [s for s in texts if s]


@pytest.mark.parametrize("lang", LANGS)
def test_gui_builds_with_no_leftover_keys(tk_root, lang):
    texts = _build_everything(tk_root, lang)
    assert len(texts) > 30, f"{lang} 只收集到 {len(texts)} 條文字，走訪大概壞了"
    leftover = sorted({s for s in texts if KEY_LIKE.match(s)})
    assert not leftover, f"{lang} 畫面上有未翻譯的 key：{leftover}"


def test_all_languages_produce_the_same_widget_count(tk_root):
    """四語的 widget 文字條數應該一致——某語言少一條通常代表 t() 回了
    空字串或 widget 建置在該語言下走了不同分支。"""
    counts = {lang: len(_build_everything(tk_root, lang)) for lang in LANGS}
    assert len(set(counts.values())) == 1, f"各語言條數不一致：{counts}"


def teardown_module(module):
    i18n.set_lang(i18n.DEFAULT_LANG)
