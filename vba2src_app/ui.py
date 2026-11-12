"""コンソールの案内と結果表示。端末以外には制御コードを出力しない。"""

import os
import sys
from importlib.resources import files
from pathlib import Path

from . import __version__
from .runtime import 安全に表示

# ── 1. 端末の色と表示値 ───────────────────────────────────
色コード = {"案内": "36", "成功": "32", "注意": "33", "失敗": "31", "補足": "90"}


def 端末色を使える(ストリーム):
    """TTYで色が使える場合だけ有効化する。リダイレクトとNO_COLORを優先する。"""
    if ストリーム is None or "NO_COLOR" in os.environ or os.environ.get("TERM") == "dumb":
        return False
    try:
        if not ストリーム.isatty():
            return False
        if os.name != "nt":
            return True
        import ctypes
        import msvcrt

        関数群 = ctypes.WinDLL("kernel32", use_last_error=True)
        ハンドル = ctypes.c_void_p(msvcrt.get_osfhandle(ストリーム.fileno()))
        モード = ctypes.c_uint32()
        return bool(
            関数群.GetConsoleMode(ハンドル, ctypes.byref(モード))
            and 関数群.SetConsoleMode(ハンドル, モード.value | 0x0004)
        )
    except (OSError, ValueError, AttributeError):
        return False


def 色付き表示(メッセージ, 種類="案内", 例外情報=False):
    """端末だけにANSI色を付け、ログやパイプでは通常のテキストにする。"""
    ストリーム = sys.stderr if 例外情報 else sys.stdout
    if 端末色を使える(ストリーム):
        メッセージ = f"\x1b[1;{色コード[種類]}m{メッセージ}\x1b[0m"
    安全に表示(メッセージ, 例外情報)


def 表示値(値):
    """パス等の制御文字をエスケープし、結果の行構造を保つ。"""
    return "".join(
        f"\\u{ord(c):04x}" if ord(c) < 32 or 127 <= ord(c) <= 159 or c in "\u2028\u2029" else c
        for c in str(値)
    )


# ── 2. 起動案内と入力の見出し ─────────────────────────────
def 起動案内を表示():
    """ビルドスクリプトと同じASCIIタイトル、対応形式、短い操作案内を表示する。"""
    安全に表示("")
    色付き表示(files("vba2src_app").joinpath("title.txt").read_text(encoding="ascii").rstrip("\n"))
    色付き表示(f"  VBA2SRC {__version__}", "補足")
    安全に表示("\n  Officeを起動せず、VBA・名前定義を抽出します。\n")
    色付き表示("  【対応ファイル】")
    安全に表示("  Excel             .xlsx .xlsm .xls .xla .xlsb .xltm .xlam .xltx")
    安全に表示("  Word              .docx .docm .dotx .dotm .doc .dot")
    安全に表示("  Access            .accdb .mdb")
    安全に表示("\n  パスを貼り付けてEnter。中止は q。")


def 入力見出しを表示(番号, 見出し, 補足=None):
    """質問ごとに余白を設け、入力対象と既定値を分けて表示する。"""
    色付き表示(f"\n  {番号:02d}  {見出し}")
    if 補足:
        色付き表示(f"      {補足}", "補足")


# ── 3. 終了状態の表示 ─────────────────────────────────────
def 通知(タイトル, メッセージ, 例外情報=False):
    """ポップアップを使わず、利用可能なコンソールへ結果と原因を表示する。"""
    種類 = "失敗" if 例外情報 else "案内"
    色付き表示(f"\n  ── {タイトル} ──", 種類, 例外情報)
    安全に表示("\n".join(f"  {行}" for 行 in メッセージ.splitlines()), 例外情報)
    安全に表示("", 例外情報)


def 抽出結果を表示(結果, フォルダ, 経過秒):
    """取得済み件数と確認先をまとめ、成功・一部成功・失敗を色と文言で区別する。"""
    失敗 = 結果.終了コード == 1
    種類 = {0: "成功", 3: "注意", 1: "失敗"}[結果.終了コード]
    見出し = {0: "抽出完了", 3: "一部の情報を取得できませんでした", 1: "抽出に失敗しました"}[結果.終了コード]
    色付き表示(f"\n  ── {見出し} ──", 種類, 失敗)
    安全に表示(f"  対象：{表示値(結果.ファイル名)}", 失敗)
    色付き表示(f"  結果：{結果.状態}  ／  終了コード {結果.終了コード}", 種類, 失敗)
    安全に表示(f"  VBA：{結果.VBA有無}", 失敗)
    安全に表示(f"  取得プロジェクト数：{len(結果.プロジェクト一覧)}", 失敗)
    安全に表示(f"  取得モジュール数：{sum(len(p.モジュール一覧) for p in 結果.プロジェクト一覧)}", 失敗)
    安全に表示(
        f"  取得した名前定義：{len(結果.名前定義情報['items'])}件（{結果.名前定義情報['status']}）", 失敗
    )
    色付き表示(f"  処理時間：{経過秒:.2f}秒", "補足", 失敗)
    色付き表示("\n  結果フォルダ", "案内", 失敗)
    安全に表示(f"  {表示値(フォルダ)}", 失敗)
    色付き表示("\n  最初に確認するファイル", "案内", 失敗)
    安全に表示(f"  {表示値(Path(フォルダ) / '抽出結果一覧.txt')}", 失敗)
    if 結果.終了コード:
        色付き表示("  詳細は抽出結果一覧の診断コードと実行ログを確認してください。", 種類, 失敗)
    安全に表示("", 失敗)
