# tests/test_first_run_language.py
"""首次啟動選語言的行為。

⚠ `_pick_language_on_first_run` 裡面是 `root.wait_window(dlg)`，會卡在
   自己的事件迴圈。測試不可以直接呼叫後等它回來——要先用 `after()` 排一個
   模擬點擊，wait_window 的迴圈會把它跑掉，函式才回得來。
"""

import json
import os
import tkinter as tk
from tkinter import ttk

import pytest

import config
import i18n
import main


def _click_first_language_button(dlg_parent, delay=60):
    """在事件迴圈裡找到語言視窗，按下第 2 個語言（简体中文）的按鈕。"""
    def go():
        for w in dlg_parent.winfo_children():
            if isinstance(w, tk.Toplevel) and w.title() == "Language":
                buttons = [c for c in w.winfo_children()
                           if isinstance(c, ttk.Button)]
                buttons[1].invoke()
                return
        # 視窗還沒建好就再等一輪
        dlg_parent.after(delay, go)
    dlg_parent.after(delay, go)


@pytest.fixture
def tmp_cfg(tmp_path, monkeypatch):
    path = str(tmp_path / "config.json")
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    monkeypatch.setattr(config, "CONFIG_PATH", path)
    return path


def test_first_run_shows_dialog_and_saves_choice(tk_root, tmp_cfg):
    assert not os.path.exists(tmp_cfg)

    win = tk.Toplevel(tk_root)
    win.withdraw()
    _click_first_language_button(win)
    main._pick_language_on_first_run(win)
    win.destroy()

    assert os.path.exists(tmp_cfg), "選完語言沒有寫出 config.json"
    with open(tmp_cfg, encoding="utf-8") as f:
        saved = json.load(f)
    assert saved["language"] == i18n.LANGUAGES[1][0]


def test_second_run_does_not_ask_again(tk_root, tmp_cfg):
    config.save_config({"language": "en"}, tmp_cfg)

    win = tk.Toplevel(tk_root)
    win.withdraw()
    before = len(win.winfo_children())
    # 沒有排任何模擬點擊：若它還是開了視窗，wait_window 會卡住直到逾時，
    # 這裡能直接返回就代表它認得「選過了」
    main._pick_language_on_first_run(win)
    assert len(win.winfo_children()) == before, "已選過語言卻又跳了選語言視窗"
    win.destroy()

    assert config.load_config(tmp_cfg)["language"] == "en"


def test_closing_the_dialog_still_saves_a_language(tk_root, tmp_cfg):
    """直接關掉視窗＝接受第一個選項並照樣存檔，否則下次開又跳一次。"""
    win = tk.Toplevel(tk_root)
    win.withdraw()

    def close_it():
        for w in win.winfo_children():
            if isinstance(w, tk.Toplevel) and w.title() == "Language":
                w.destroy()
                return
        win.after(60, close_it)
    win.after(60, close_it)

    main._pick_language_on_first_run(win)
    win.destroy()

    assert config.load_config(tmp_cfg)["language"] == i18n.LANGUAGES[0][0]


def test_unknown_language_code_falls_back_to_default():
    """舊 config 的怪值不能讓程式起不來。"""
    assert i18n.set_lang("klingon") == i18n.DEFAULT_LANG
    assert i18n.set_lang(None) == i18n.DEFAULT_LANG
    assert i18n.set_lang("ja") == "ja"
    i18n.set_lang(i18n.DEFAULT_LANG)


def test_t_returns_the_key_when_nothing_matches():
    assert i18n.t("no.such.key.at.all") == "no.such.key.at.all"
