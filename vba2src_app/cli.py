"""引数と質問への回答を受け付け、抽出・保存・コンソール表示を統括する。"""

import argparse
import logging
import stat
import sys
import time
from importlib import metadata as メタデータ
from pathlib import Path

import vba2src_config as config

from . import __version__
from .paths import 保存先の親を決定, 既定親フォルダを取得, 結果フォルダを作成
from .runtime import LFログハンドラー, コンソールを準備, 例外を記録, 安全に表示, 標準ストリームを補完
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
            except OSError, ValueError:
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
            except OSError, ValueError:
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
    except EOFError, KeyboardInterrupt:
        色付き表示("\n  中止しました。", "注意")
    if 終了待ち and sys.stdin.isatty():
        try:
            input("  Enterキーで終了 > ")
        except EOFError, KeyboardInterrupt:
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
    フォルダ = None
    ログハンドラー = None
    ロガー = logging.getLogger("vba2src")
    開始時刻 = time.monotonic()
    処理段階 = "入力ファイルの確認"
    try:
        対象 = Path(解析済み引数.対象).resolve()
        try:
            対象ファイル状態 = 対象.stat()
        except FileNotFoundError:
            通知(
                "VBA2SRC 失敗",
                "INPUT_NOT_FOUND：対象ファイルが見つかりません。パスとネットワーク接続を確認してください。",
                True,
            )
            return 1
        except PermissionError:
            通知(
                "VBA2SRC 失敗",
                "INPUT_ACCESS_DENIED：対象ファイルへのアクセスが拒否されました。",
                True,
            )
            return 1
        if not stat.S_ISREG(対象ファイル状態.st_mode):
            通知(
                "VBA2SRC 失敗",
                "対象ファイルが存在しないか、通常ファイルではありません。",
                True,
            )
            return 1
        処理段階 = "保存先の決定"
        # ドラッグ起動を特別扱いせず、引数付き起動の保存規則を一か所で決定する。
        親 = 保存先の親を決定(対象, 解析済み引数.出力)
        処理段階 = "結果フォルダの作成"
        フォルダ = 結果フォルダを作成(親, 対象.name)
        ログフォルダ = フォルダ / "ログ"
        ログフォルダ.mkdir()
        ログハンドラー = LFログハンドラー(ログフォルダ / "実行ログ.txt", encoding="utf-8")
        ログハンドラー.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        ロガー.addHandler(ログハンドラー)
        ロガー.setLevel(logging.DEBUG if 解析済み引数.詳細ログ else logging.INFO)
        ロガー.propagate = False
        ロガー.info("開始 VBA2SRC %s input=%s output=%s", __version__, 対象, フォルダ)
        ロガー.info("Python=%s OS=%s", sys.version.split()[0], sys.platform)
        for 配布パッケージ名 in ("oletools", "olefile", "msoffcrypto-tool", "cryptography", "pyOpenVBA"):
            try:
                バージョン = メタデータ.version(配布パッケージ名)
            except メタデータ.PackageNotFoundError:
                バージョン = "不明（バンドル）"
            ロガー.info("library %s=%s", 配布パッケージ名, バージョン)
        from .extract import 抽出
        from .output import 抽出結果を保存

        処理段階 = "入力ファイルの解析"
        結果 = 抽出(対象, config.VBA保護パスワード候補)
        処理段階 = "抽出結果の書込み"
        抽出結果を保存(フォルダ, 結果)
        for コード in sorted(
            set(
                結果.失敗一覧
                + [
                    コード
                    for 対象プロジェクト in 結果.プロジェクト一覧
                    for コード in 対象プロジェクト.失敗一覧
                ]
            )
        ):
            ロガー.warning("%s", コード)
        ロガー.info(
            "終了 result=%s code=%d elapsed=%.3f", 結果.状態, 結果.終了コード, time.monotonic() - 開始時刻
        )
        # 成功通知の前にログを閉じ、最後の共有書込み失敗も終了コードへ反映する。
        ロガー.removeHandler(ログハンドラー)
        try:
            ログハンドラー.close()
        except OSError:
            ログハンドラー.書込み失敗 = True
        if ログハンドラー.書込み失敗:
            通知(
                "VBA2SRC 失敗",
                "LOG_WRITE_FAILED：実行ログを書き込めません。共有先の接続と権限を確認してください。",
                True,
            )
            return 1
        抽出結果を表示(結果, フォルダ, time.monotonic() - 開始時刻)
        return 結果.終了コード
    except KeyboardInterrupt:
        安全に表示("ユーザーにより中止しました。", True)
        return 130
    except Exception:
        if ログハンドラー:
            例外を記録("RUN_FAILED")
            ロガー.error("RUN_FAILED")
        import sys as system

        例外種類 = system.exc_info()[0].__name__
        通知(
            "VBA2SRC 失敗",
            f"{処理段階}に失敗しました（{例外種類}）。ネットワーク接続、権限、空き容量を確認してください。"
            + (f"\n途中結果：{フォルダ}" if フォルダ else ""),
            True,
        )
        return 1
    finally:
        if ログハンドラー:
            ロガー.removeHandler(ログハンドラー)
            try:
                ログハンドラー.close()
            except OSError:
                安全に表示("LOG_CLOSE_FAILED：実行ログの終了書込みに失敗しました。", True)
