"""配布時はEXEの格納フォルダの一つ上から、公開設定の候補リストだけを読む。"""

import ast
import sys
from pathlib import Path

候補設定名 = "VBA保護パスワード候補"


class 候補設定エラー(ValueError):
    """候補設定の読み取りに失敗したことを案内する。"""


def 既定候補を読む():
    """文字列のリストまたはタプルを読み取る。設定ファイルがない場合は候補なし。"""
    if getattr(sys, "frozen", False):
        ルート = Path(sys.executable).resolve().parent.parent
    else:
        ルート = Path(__file__).resolve().parent.parent
    パス = ルート / "vba2src_config.py"
    try:
        設定 = ast.parse(パス.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return ()
    except (OSError, UnicodeError, SyntaxError):
        raise 候補設定エラー(
            "vba2src_config.py を読めません。UTF-8で保存し、ファイルの構文を確認してください。"
        ) from None

    候補定義 = [
        文.value
        for 文 in 設定.body
        if isinstance(文, ast.Assign)
        and any(isinstance(名前, ast.Name) and 名前.id == 候補設定名 for 名前 in 文.targets)
    ]
    if len(候補定義) == 1 and isinstance(候補定義[0], (ast.List, ast.Tuple)):
        値一覧 = 候補定義[0].elts
        if all(isinstance(値, ast.Constant) and isinstance(値.value, str) for 値 in 値一覧):
            return tuple(値.value for 値 in 値一覧)
    raise 候補設定エラー(
        "vba2src_config.py の VBA保護パスワード候補 を文字列のリストまたはタプルで1件定義してください。"
    )
