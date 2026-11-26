"""ユーザーの一時フォルダに保存する利用者ごとの初期値。"""

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile, gettempdir

設定パス = Path(gettempdir()) / "VBA2SRC" / "settings.json"


@dataclass
class GUI設定:
    保存先: str = ""
    画面フォント: str = ""
    画面サイズ: int = 10
    コードフォント: str = ""
    コードサイズ: int = 10
    VBA候補: list[str] | None = None


def 設定を読む():
    try:
        値 = json.loads(設定パス.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return GUI設定()
    if not isinstance(値, dict):
        return GUI設定()
    既定値 = asdict(GUI設定())
    for 名前, 既定 in 既定値.items():
        設定値 = 値.get(名前)
        if type(設定値) is type(既定):
            既定値[名前] = 設定値
    if not 9 <= 既定値["画面サイズ"] <= 14:
        既定値["画面サイズ"] = 10
    if not 8 <= 既定値["コードサイズ"] <= 24:
        既定値["コードサイズ"] = 10
    候補 = 値.get("VBA候補")
    if isinstance(候補, list) and all(isinstance(項目, str) for 項目 in 候補):
        既定値["VBA候補"] = 候補
    return GUI設定(**既定値)


def 設定を保存(設定):
    """書き込み途中の内容が次回起動で読まれないよう同じフォルダで置き換える。"""
    設定パス.parent.mkdir(parents=True, exist_ok=True)
    一時パス = None
    try:
        with NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n", dir=設定パス.parent, delete=False
        ) as ファイル:
            一時パス = Path(ファイル.name)
            ファイル.write(json.dumps(asdict(設定), ensure_ascii=False, indent=2) + "\n")
        os.replace(一時パス, 設定パス)
        一時パス = None
    finally:
        if 一時パス is not None:
            一時パス.unlink(missing_ok=True)
