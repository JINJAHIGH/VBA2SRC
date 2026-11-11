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
        except OSError, ValueError, UnicodeError:
            pass


# ── 2. Windowsコンソールの接続 ──────────────────────────
def コンソールを準備(作成を許可=False):
    """親コンソールへ接続する。新規作成は引数なしの対話起動だけに許可する。"""
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return False
    import ctypes

    Windows関数群 = ctypes.WinDLL("kernel32", use_last_error=True)
    Windows関数群.GetStdHandle.restype = ctypes.c_void_p
    Windows関数群.GetFileType.argtypes = [ctypes.c_void_p]
    Windows関数群.GetConsoleWindow.restype = ctypes.c_void_p
    接続済み = Windows関数群.AttachConsole(ctypes.c_uint32(-1))
    入出力あり = all(
        Windows関数群.GetFileType(Windows関数群.GetStdHandle(ctypes.c_uint32(フラグ)))
        for フラグ in (-10, -11)
    )
    作成済み = bool(
        作成を許可
        and not 接続済み
        and not Windows関数群.GetConsoleWindow()
        and not 入出力あり
        and Windows関数群.AllocConsole()
    )
    if 作成済み:
        Windows関数群.SetConsoleTitleW("VBA2SRC")
    # 呼び出し元から受け取ったリダイレクト先を優先する。
    for 属性名, フラグ in (("stdout", -11), ("stderr", -12), ("stdin", -10)):
        if getattr(sys, 属性名) is not None:
            continue
        Windows関数群.GetStdHandle.restype = ctypes.c_void_p
        ハンドル = Windows関数群.GetStdHandle(ctypes.c_uint32(フラグ))
        if ハンドル not in (None, 0, ctypes.c_void_p(-1).value):
            try:
                import msvcrt

                # 呼び出し元のハンドルを閉じないよう、複製をストリームへ渡す。
                現在値 = Windows関数群.GetCurrentProcess
                現在値.restype = ctypes.c_void_p
                複製ハンドル = ctypes.c_void_p()
                Windows関数群.DuplicateHandle.argtypes = [
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.POINTER(ctypes.c_void_p),
                    ctypes.c_uint32,
                    ctypes.c_int,
                    ctypes.c_uint32,
                ]
                if Windows関数群.DuplicateHandle(
                    現在値(), ハンドル, 現在値(), ctypes.byref(複製ハンドル), 0, False, 2
                ):
                    読取り = 属性名 == "stdin"
                    ファイル記述子 = msvcrt.open_osfhandle(
                        複製ハンドル.value, os.O_RDONLY if 読取り else os.O_WRONLY
                    )
                    setattr(
                        sys,
                        属性名,
                        os.fdopen(
                            ファイル記述子,
                            "r" if 読取り else "w",
                            encoding="utf-8",
                            errors="strict" if 読取り else "backslashreplace",
                            buffering=1,
                        ),
                    )
            except OSError, ValueError:
                pass
    return 作成済み
