"""引数の有無でGUI起動とコマンドライン実行を切り替える。"""

import argparse
import sys

from . import __version__
from .console import 抽出結果を表示, 色付き表示, 通知
from .run import 抽出を実行
from .runtime import コンソールを準備, 安全に表示, 標準ストリームを補完


# ── 1. 引数解析と通知 ──────────────────────────
class 引数エラー(Exception):
    """画面なし起動でも呼び出し元が扱える引数エラー。"""


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


# ── 2. 抽出処理の実行と終了 ──────────────────────────
def メイン(起動引数=None):
    """実行単位のログを開始し、入力・保存の失敗を固定の終了コードへ変換する。"""
    起動引数 = list(sys.argv[1:] if 起動引数 is None else 起動引数)
    if not 起動引数:
        標準ストリームを補完()
        from .gui import GUIを起動

        return GUIを起動()
    コンソールを準備()
    標準ストリームを補完()
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
    if 実行.設定の注意:
        色付き表示(f"  注意：{実行.設定の注意}", "注意", 例外情報=True)
    return 実行.終了コード
