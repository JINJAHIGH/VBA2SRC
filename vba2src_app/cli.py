"""引数と質問への回答を受け付け、抽出・保存・コンソール表示を統括する。"""

import argparse
import sys
from pathlib import Path

import vba2src_config as config

from . import __version__
from .paths import 既定親フォルダを取得
from .run import 抽出を実行
from .runtime import コンソールを準備, 安全に表示, 標準ストリームを補完
from .ui import 入力見出しを表示, 抽出結果を表示, 色付き表示, 起動案内を表示, 通知


# ── 1. 引数なしの対話入力 ──────────────────────────
def 回答を取得(質問):
    """対話入力の引用符と中止指示を扱う。候補パスワードは入力しない。"""
    値 = input(質問).strip()
    if 値.casefold() in ("q", "quit"):
        raise KeyboardInterrupt
    if len(値) >= 2 and 値[0] == 値[-1] and 値[0] in ('"', "'"):
        値 = 値[1:-1]
    return 値


def 対話モードを実行(終了待ち=False):
    """引数なし起動の質問を順に処理し、確定値を共通の実行経路へ渡す。"""
    コード = 130
    try:
        起動案内を表示()
        入力見出しを表示(1, "対象ファイル")
        while True:
            対象 = 回答を取得("  パス > ")
            try:
                if 対象 and Path(対象).is_file():
                    break
            except (OSError, ValueError):
                pass
            色付き表示("  ファイルが見つかりません。パス・接続・権限を確認してください。", "注意")
        既定値 = 既定親フォルダを取得(config.既定保存先親フォルダ)
        入力見出しを表示(2, "保存先の親フォルダ", "空欄なら次の場所へ保存します。")
        安全に表示(f"      {既定値}")
        while True:
            出力 = 回答を取得("  保存先 > ") or str(既定値)
            try:
                if not Path(出力).exists() or Path(出力).is_dir():
                    break
            except (OSError, ValueError):
                pass
            色付き表示("  フォルダとして使用できません。別の保存先を入力してください。", "注意")
        入力見出しを表示(3, "詳細ログ")
        while True:
            詳細ログ = 回答を取得("  保存しますか？ [y/N] > ").casefold()
            if 詳細ログ in ("", "n", "no", "y", "yes"):
                break
            色付き表示("  y または n を入力してください。", "注意")
        色付き表示("\n  抽出しています…")
        引数一覧 = [対象, "-o", 出力, "--no-ui"]
        if 詳細ログ in ("y", "yes"):
            引数一覧.append("--debug-log")
        コード = メイン(引数一覧)
    except (EOFError, KeyboardInterrupt):
        色付き表示("\n  中止しました。", "注意")
    if 終了待ち and sys.stdin.isatty():
        try:
            input("  Enterキーで終了 > ")
        except (EOFError, KeyboardInterrupt):
            pass
    return コード


# ── 2. 引数解析と通知 ──────────────────────────
class 引数エラー(Exception):
    """画面なし起動でも呼び出し元が扱える引数エラー。"""

    pass


class 引数解析器(argparse.ArgumentParser):
    """標準出力のないEXEでもヘルプ表示と引数エラーを安全に扱う。"""

    def error(自身, メッセージ):
        raise 引数エラー(メッセージ)

    def _print_message(自身, メッセージ, 対象ファイル=None):
        if メッセージ:
            安全に表示(メッセージ.rstrip(), 例外情報=対象ファイル is sys.stderr)


def 引数解析器を作成():
    """PythonとEXEで共通のオプションと日本語ヘルプを定義する。"""
    結果 = 引数解析器(description="Officeを起動せず、1ファイルのVBAを読み取り専用で抽出します。")
    結果.add_argument("対象", metavar="target", help="対象ファイル（1ファイルのみ）")
    結果.add_argument(
        "-o", "--output", dest="出力", metavar="output", help="抽出結果フォルダを格納する親フォルダ"
    )
    結果.add_argument(
        "--no-ui", dest="画面なし", action="store_true", help="新しい画面を作成せず、入力を待たない"
    )
    結果.add_argument(
        "--debug-log", dest="詳細ログ", action="store_true", help="詳細ログ・例外スタックを記録"
    )
    結果.add_argument(
        "--group-by-type",
        dest="種類別出力",
        action="store_true",
        help="VBAまとめフォルダにモジュールの種類別テキストを追加",
    )
    結果.add_argument("--version", action="version", version=f"VBA2SRC {__version__}")
    return 結果


# ── 3. 抽出処理の実行と終了 ──────────────────────────
def メイン(起動引数=None):
    """実行単位のログを開始し、入力・保存の失敗を固定の終了コードへ変換する。"""
    起動引数 = list(sys.argv[1:] if 起動引数 is None else 起動引数)
    コンソール作成済み = コンソールを準備(作成を許可=not 起動引数)
    標準ストリームを補完()
    if not 起動引数:
        return 対話モードを実行(終了待ち=コンソール作成済み)
    try:
        解析済み引数 = 引数解析器を作成().parse_args(起動引数)
    except 引数エラー as 例外情報:
        通知(
            "VBA2SRC 引数エラー",
            f"引数エラー：{例外情報}\n対象ファイルは1回につき1ファイルです。--helpで使い方を確認してください。",
            True,
        )
        return 2
    except SystemExit as 終了情報:
        return 終了情報.code
    実行 = 抽出を実行(
        解析済み引数.対象,
        解析済み引数.出力,
        詳細ログ=解析済み引数.詳細ログ,
        種類別出力=解析済み引数.種類別出力,
    )
    if 実行.メッセージ:
        保存先案内 = f"\n途中結果：{実行.フォルダ}" if 実行.フォルダ else ""
        通知("VBA2SRC 失敗", 実行.メッセージ + 保存先案内, True)
    else:
        抽出結果を表示(実行.結果, 実行.フォルダ, 実行.経過秒)
    return 実行.終了コード
