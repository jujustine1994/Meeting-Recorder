# tests/test_i18n.py
"""三道防止退化的測試。

第 3 條（不得寫死中日文）是**永久**的：它擋的不是這次遷移，是下一次。
新增功能時順手寫一個中文按鈕標籤最自然不過，沒有它三個月後就又回到
全部寫死的狀態。
"""

import ast
import os
import re

import pytest

import i18n

# ⚠ 本專案的 .py 在**根目錄**，不在 src/。
# 寫成 (ROOT / "src").rglob("*.py") 會回空 list，parametrize 收集到 0 個
# case——測試「通過」但其實什麼都沒檢查。下面的 test_scan_set_is_not_empty
# 就是釘住這件事的。
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKIP_DIRS = {"venv", "tests", "locales", "__pycache__", "docs", ".git", "logs"}

CJK = re.compile(r"[一-鿿぀-ヿ]")
PLACEHOLDER = re.compile(r"\{[a-zA-Z_][a-zA-Z0-9_]*\}")
LANGS = [code for code, _, _ in i18n.LANGUAGES]

# 豁免清單。每一條都要有理由——沒理由的豁免等於把第三條測試關掉。
# ⚠ main.py **不在**這裡：本工具的 GUI 全部住在 main.py，把它豁免掉
#   等於整條測試失效。log 字串因此抽到 logtext.py。
ALLOWLIST = {
    "i18n.py",      # 語言自稱（「繁體中文」「日本語」）本來就該用各語言自己的說法
    "logtext.py",   # logs/app.log 的內容永遠繁中，不跟使用者介面語言走
}


def _scannable():
    out = []
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for name in sorted(filenames):
            if name.endswith(".py") and name not in ALLOWLIST:
                out.append(os.path.join(dirpath, name))
    return sorted(out)


SCAN = _scannable()


def test_every_language_has_the_same_keys():
    """任一語言少一條就紅燈。新增語言時漏翻幾條是必然，靠人眼比對上百條
    不可能可靠——這條測試就是那個「不可能可靠」的替代品。"""
    base = set(i18n._strings(i18n.FALLBACK_LANG))
    assert base, "母表是空的，locale 載入壞了"
    for lang in LANGS:
        keys = set(i18n._strings(lang))
        assert not (base - keys), f"{lang} 少了：{sorted(base - keys)[:10]}"
        assert not (keys - base), f"{lang} 多了：{sorted(keys - base)[:10]}"


def test_placeholders_match_across_languages():
    """譯文的 {name} 打錯或漏掉，t() 會 format 失敗並吐出未格式化的原字串——
    畫面上看到 {name} 殘留，不會 crash 所以特別容易漏掉。"""
    base = i18n._strings(i18n.FALLBACK_LANG)
    for lang in LANGS:
        table = i18n._strings(lang)
        for key, src in base.items():
            assert set(PLACEHOLDER.findall(src)) == set(PLACEHOLDER.findall(table[key])), \
                f"{lang} / {key} 的 placeholder 不一致"


def _hardcoded_cjk(path):
    """docstring 與註解不算——那些是寫給人看的說明。"""
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    docs = set()
    for n in ast.walk(tree):
        if isinstance(n, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            b = n.body
            if b and isinstance(b[0], ast.Expr) and isinstance(b[0].value, ast.Constant) \
                    and isinstance(b[0].value.value, str):
                docs.add(id(b[0].value))
    return sorted((n.lineno, n.value) for n in ast.walk(tree)
                  if isinstance(n, ast.Constant) and isinstance(n.value, str)
                  and CJK.search(n.value) and id(n) not in docs)


def test_scan_set_is_not_empty():
    """反向測試：豁免清單一寫寬、或掃描路徑寫錯，下面那條就靜默零覆蓋。
    這條釘住掃描範圍確實含主程式。"""
    assert SCAN, "掃描到 0 個檔案，下面那條測試等於沒跑"
    names = {os.path.basename(p) for p in SCAN}
    assert "main.py" in names, "主程式（GUI 全在裡面）不在掃描範圍內"


@pytest.mark.parametrize("path", SCAN, ids=lambda p: os.path.basename(p))
def test_no_hardcoded_cjk(path):
    """介面文字一律走 t()。真的需要豁免就加進 ALLOWLIST，但要寫清楚理由。"""
    hits = _hardcoded_cjk(path)
    assert not hits, f"{os.path.basename(path)} 有 {len(hits)} 條寫死的中日文：" + \
        "; ".join(f"行 {ln}: {v[:40]!r}" for ln, v in hits[:5])
