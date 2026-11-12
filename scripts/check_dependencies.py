"""ビルド前に固定バージョンと必要な実行ライブラリを確認する。自動インストールはしない。"""

import importlib
import platform
import re
import struct
import sys
from importlib import metadata
from pathlib import Path


def 要件を収集(定義パス, 読取済み=None):
    読取済み = set() if 読取済み is None else 読取済み
    定義パス = 定義パス.resolve()
    if 定義パス in 読取済み:
        return {}
    読取済み.add(定義パス)
    要件一覧 = {}
    for 行 in 定義パス.read_text(encoding="utf-8").splitlines():
        行 = 行.split("#", 1)[0].strip()
        if not 行:
            continue
        if 行.startswith("-r "):
            要件一覧.update(要件を収集(定義パス.parent / 行[3:].strip(), 読取済み))
            continue
        一致 = re.fullmatch(r"([A-Za-z0-9_.-]+)==([^;\s]+)", 行)
        if not 一致:
            raise ValueError("未対応の依存定義：" + 行)
        要件一覧[一致[1]] = 一致[2]
    return 要件一覧


def メイン():
    ルート = Path(__file__).resolve().parent.parent
    print(f"Python {platform.python_version()} / {struct.calcsize('P') * 8} bit / {sys.platform}")
    if sys.version_info < (3, 13):
        print("ビルドには Python 3.13 以上が必要です。", file=sys.stderr)
        return 1
    if sys.platform != "win32":
        print("Windows用EXEのビルドはWindowsで実行してください。", file=sys.stderr)
        return 1
    失敗一覧 = []
    for 名前, 必要版 in sorted(要件を収集(ルート / "requirements-build.txt").items()):
        try:
            現在版 = metadata.version(名前)
            if 現在版 != 必要版:
                失敗一覧.append(f"{名前}: 必要 {必要版} / 現在 {現在版}")
            else:
                print(f"OK {名前}=={現在版}")
        except metadata.PackageNotFoundError:
            失敗一覧.append(f"{名前}=={必要版}: 未導入")
    for モジュール名 in (
        "olefile",
        "oletools.olevba",
        "msoffcrypto",
        "pyopenvba.access.database",
        "cryptography.hazmat.bindings._rust",
    ):
        try:
            importlib.import_module(モジュール名)
            print("読込み成功：" + モジュール名)
        except Exception as 例外:
            失敗一覧.append(f"{モジュール名}: {type(例外).__name__}")
    if 失敗一覧:
        print("\n依存関係の確認に失敗しました。", file=sys.stderr)
        for 診断 in 失敗一覧:
            print(診断, file=sys.stderr)
        print("導入例：py -3.13 -m pip install -r requirements-build.txt", file=sys.stderr)
        return 1
    print("依存ライブラリのバージョンと読込みを確認しました。")
    return 0


if __name__ == "__main__":
    raise SystemExit(メイン())
