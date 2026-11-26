"""標準ライブラリだけで、ログとWindowsコンソールの差異を吸収する。"""

import logging
import os
import sys
import traceback


# ── 1. ログと標準ストリーム ──────────────────────────
class LFログハンドラー(logging.FileHandler):
    """LFでログを書き込み、共有切断時の書込み失敗を呼び出し元へ伝える。"""

    書込み失敗 = False

    def _open(自身):
        return open(自身.baseFilename, 自身.mode, encoding=自身.encoding, errors=自身.errors, newline="\n")

    def handleError(自身, レコード):
        # 共有切断時も元の例外を隠さず、ログ内容の二次的なダンプを抑止する。
        自身.書込み失敗 = True


def 標準ストリームを補完():
    # 依存ライブラリはimport時にもflushする。画面を作らず有効なストリームを補う。
    """画面のないEXEでも依存ライブラリが標準ストリームを安全に利用できるようにする。"""
    for 対象属性名, モード in (("stdout", "w"), ("stderr", "w"), ("stdin", "r")):
        if getattr(sys, 対象属性名) is None:
            setattr(sys, 対象属性名, open(os.devnull, モード, encoding="utf-8"))


def 例外を記録(コード):
    """秘密値を含み得る例外メッセージや変数を避け、型とスタック位置を記録する。"""
    種類, 未使用, スタック = sys.exc_info()
    logging.getLogger("vba2src").debug(
        "%s exception=%s\n%s",
        コード,
        種類.__name__ if 種類 else "unknown",
        "".join(traceback.format_list(traceback.extract_tb(スタック))),
    )


def 安全に表示(メッセージ, 例外情報=False):
    """利用可能な標準出力へ通知する。出力先の消失による二次例外を抑止する。"""
    ストリーム = sys.stderr if 例外情報 else sys.stdout
    if ストリーム is not None:
        try:
            ストリーム.write(str(メッセージ) + "\n")
            ストリーム.flush()
        except (OSError, ValueError, UnicodeError):
            pass


# ── 2. Windowsコンソールの接続 ──────────────────────────
def コンソールを準備():
    """Windows配布EXEを親コンソールへ接続する。"""
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return
    from .windows_console import 標準入出力を接続

    標準入出力を接続()
