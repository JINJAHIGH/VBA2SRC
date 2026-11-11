"""配布対象ライブラリのライセンスを収集し、同梱用の通知を生成する。"""

import sys
from importlib import metadata as メタデータ
from pathlib import Path

from packaging.requirements import Requirement

リポジトリルート = Path(__file__).resolve().parent.parent


def メイン():
    未処理一覧 = ["oletools", "olefile", "msoffcrypto-tool", "pyOpenVBA", "PyInstaller"]
    処理済み = set()
    章一覧 = []
    while 未処理一覧:
        名前 = 未処理一覧.pop(0)
        配布情報 = メタデータ.distribution(名前)
        正規化名 = 配布情報.metadata["Name"].lower().replace("_", "-")
        if 正規化名 in 処理済み:
            continue
        処理済み.add(正規化名)
        章一覧.append(f"\n{'=' * 72}\n{配布情報.metadata['Name']} {配布情報.version}\n")
        検出済み = False
        for 対象ファイル in sorted(配布情報.files or [], key=str):
            if any(
                パーツ.lower().startswith(("license", "licence", "copying", "notice"))
                for パーツ in 対象ファイル.parts
            ):
                パス = Path(配布情報.locate_file(対象ファイル))
                if パス.is_file():
                    章一覧.append(f"\n{対象ファイル}\n" + パス.read_text(encoding="utf-8", errors="replace"))
                    検出済み = True
        if not 検出済み:
            章一覧.append(
                str(
                    配布情報.metadata.get("License-Expression")
                    or 配布情報.metadata.get("License")
                    or "See installed distribution metadata."
                )
            )
        for 元データ in 配布情報.requires or []:
            依存要件 = Requirement(元データ)
            if 依存要件.marker is None or 依存要件.marker.evaluate({"extra": ""}):
                未処理一覧.append(依存要件.name)
    for パス in [Path(sys.base_prefix) / "LICENSE.txt"]:
        if パス.is_file():
            章一覧.append(
                f"\n{'=' * 72}\nPython: {パス.name}\n" + パス.read_text(encoding="utf-8", errors="replace")
            )
    出力 = リポジトリルート / "build/THIRD_PARTY_NOTICES.txt"
    出力.parent.mkdir(parents=True, exist_ok=True)
    出力.write_text("VBA2SRC bundled software notices\n" + "\n".join(章一覧), encoding="utf-8", newline="\n")
    print(出力)


if __name__ == "__main__":
    メイン()
