# tests/conftest.py
"""共用 fixture。

兩件事：
1. 把專案根目錄放進 sys.path，並把 pyaudiowpatch / lameenc 換成 mock，
   讓測試不需要真的音效卡也跑得起來。
2. 提供 session 級的隱藏 tk.Tk root。
"""

import os
import sys
import unittest.mock

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# 音訊/編碼原生套件換成 mock。get_device_count 要回真的整數，
# 否則裝置測試視窗的 range() 會炸在 MagicMock 上。
_pa_mod = unittest.mock.MagicMock()
_pa_inst = _pa_mod.PyAudio.return_value
_pa_inst.get_device_count.return_value = 0
_pa_inst.get_host_api_info_by_type.return_value = {"index": 0, "defaultOutputDevice": 0}

unittest.mock.patch.dict("sys.modules", {
    "pyaudiowpatch": _pa_mod,
    "lameenc": unittest.mock.MagicMock(),
}).start()


@pytest.fixture(scope="session")
def tk_root():
    """整個測試 session 共用一個隱藏的 Tk root。

    ⚠ 不可以每個測試建一個 tk.Tk()：Microsoft Store 版 Python 在短時間內
    反覆建立／銷毀 Tcl 直譯器時，會間歇性地丟
    `TclError: Can't find a usable init.tcl ... No error`——測試看起來隨機
    紅綠，跟被測程式一點關係都沒有。要另一個視窗就開 Toplevel。
    """
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    yield root
    try:
        root.destroy()
    except tk.TclError:
        pass
