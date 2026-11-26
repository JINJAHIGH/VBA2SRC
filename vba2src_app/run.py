"""CLIとGUIが共有する入力確認、抽出、保存、ログの実行処理。"""

import logging
import stat
import sys
import time
from dataclasses import dataclass
from importlib import metadata as メタデータ
from pathlib import Path
from threading import get_ident

from . import __version__
from .candidate_config import 候補設定エラー, 既定候補を読む
from .model import 抽出結果
from .paths import 保存先の親を決定, 結果フォルダを作成
from .preferences import 設定を読む
from .runtime import LFログハンドラー, 例外を記録

実行ログ = logging.getLogger("vba2src")
実行ログ.addHandler(logging.NullHandler())
実行ログ.propagate = False
実行ログ.setLevel(logging.DEBUG)


@dataclass
class 実行結果:
    終了コード: int
    フォルダ: Path | None = None
    結果: 抽出結果 | None = None
    メッセージ: str = ""
    経過秒: float = 0
    設定の注意: str = ""


def 抽出を実行(対象パス, 保存先=None, *, 詳細ログ=False, 種類別出力=False):
    """入力を読み取り専用で処理し、画面に依存しない実行結果を返す。"""
    フォルダ = None
    結果 = None
    ログハンドラー = None
    設定の注意 = ""
    開始時刻 = time.monotonic()
    処理段階 = "入力ファイルの確認"
    try:
        対象 = Path(対象パス).resolve()
        try:
            対象ファイル状態 = 対象.stat()
        except FileNotFoundError:
            return 実行結果(
                1, メッセージ="INPUT_NOT_FOUND：対象ファイルが見つかりません。パスと接続を確認してください。"
            )
        except PermissionError:
            return 実行結果(1, メッセージ="INPUT_ACCESS_DENIED：対象ファイルへのアクセスが拒否されました。")
        if not stat.S_ISREG(対象ファイル状態.st_mode):
            return 実行結果(
                1, メッセージ="対象は通常ファイルではありません。Officeファイルを指定してください。"
            )
        処理段階 = "VBA保護パスワード候補の読込み"
        利用者候補 = 設定を読む().VBA候補
        候補 = 利用者候補
        if 候補 is None:
            try:
                候補 = 既定候補を読む()
            except 候補設定エラー as 例外情報:
                設定の注意 = str(例外情報) + " VBA保護の候補なしで抽出します。"
                候補 = ()

        処理段階 = "保存先の決定"
        親 = 保存先の親を決定(対象, 保存先)
        処理段階 = "結果フォルダの作成"
        フォルダ = 結果フォルダを作成(親, 対象.name)
        if 詳細ログ:
            処理段階 = "詳細ログの準備"
            ログフォルダ = フォルダ / "ログ"
            ログフォルダ.mkdir()
            ログハンドラー = LFログハンドラー(ログフォルダ / "実行ログ.txt", encoding="utf-8")
            ログハンドラー.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
            # 同じプロセスから並行実行しても、対象ごとのログを混在させない。
            実行スレッド = get_ident()
            ログハンドラー.addFilter(lambda レコード: レコード.thread == 実行スレッド)
            実行ログ.addHandler(ログハンドラー)
            実行ログ.info("開始 VBA2SRC %s input=%s output=%s", __version__, 対象, フォルダ)
            実行ログ.info("Python=%s OS=%s", sys.version.split()[0], sys.platform)
            if 設定の注意:
                実行ログ.warning("%s", 設定の注意)
            for 名前 in ("oletools", "olefile", "msoffcrypto-tool", "cryptography", "pyOpenVBA"):
                try:
                    バージョン = メタデータ.version(名前)
                except メタデータ.PackageNotFoundError:
                    バージョン = "不明（バンドル）"
                実行ログ.info("library %s=%s", 名前, バージョン)

        処理段階 = "入力ファイルの解析"
        from .extract import 抽出
        from .output import 抽出結果を保存

        結果 = 抽出(対象, 候補)
        処理段階 = "抽出結果の書込み"
        抽出結果を保存(フォルダ, 結果, 種類別出力=種類別出力)
        if ログハンドラー:
            失敗コード = set(結果.失敗一覧)
            for 対象プロジェクト in 結果.プロジェクト一覧:
                失敗コード.update(対象プロジェクト.失敗一覧)
            for コード in sorted(失敗コード):
                実行ログ.warning("%s", コード)
            実行ログ.info(
                "終了 result=%s code=%d elapsed=%.3f", 結果.状態, 結果.終了コード, time.monotonic() - 開始時刻
            )
        終了コード, メッセージ = 結果.終了コード, ""
    except KeyboardInterrupt:
        終了コード, メッセージ = 130, "ユーザーにより中止しました。"
    except Exception as 例外情報:
        例外を記録("RUN_FAILED")
        終了コード = 1
        メッセージ = (
            f"{処理段階}に失敗しました（{type(例外情報).__name__}）。"
            "ネットワーク接続、権限、空き容量を確認してください。"
        )
    finally:
        if ログハンドラー:
            実行ログ.removeHandler(ログハンドラー)
            try:
                ログハンドラー.close()
            except OSError:
                ログハンドラー.書込み失敗 = True

    if ログハンドラー and ログハンドラー.書込み失敗 and not メッセージ:
        終了コード = 1
        メッセージ = "LOG_WRITE_FAILED：実行ログを書き込めません。保存先の接続と権限を確認してください。"
    return 実行結果(終了コード, フォルダ, 結果, メッセージ, time.monotonic() - 開始時刻, 設定の注意)
