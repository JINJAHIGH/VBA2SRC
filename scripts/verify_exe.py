"""生成したWindows EXEを、開発用PATHに依存しない条件で検証する。"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

リポジトリルート = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(リポジトリルート))
from tests.samples import (  # noqa: E402
    Accessデータベースを生成,
    Word文書を生成,
    ブックを生成,
    プロジェクトストリームを生成,
    複合ファイルを生成,
)


def 比較内容を取得(フォルダ):
    return {
        p.relative_to(フォルダ).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
        for p in フォルダ.rglob("*")
        if p.is_file() and "ログ" not in p.parts
    }


def メイン():
    if sys.platform != "win32":
        raise SystemExit("Windows-only verification")
    # 試験ごとに作業先を確保し、過去の確認結果を上書きしない。
    import tempfile

    親 = リポジトリルート / "build/verification-smoke"
    親.mkdir(parents=True, exist_ok=True)
    作業フォルダ = Path(tempfile.mkdtemp(prefix="run_", dir=親))
    実行ファイル = 作業フォルダ / "VBA2SRC.exe"
    shutil.copy2(リポジトリルート / "dist/VBA2SRC.exe", 実行ファイル)
    マクロファイル = ブックを生成(作業フォルダ / "確認 用 (1).xlsm")
    入力フォルダ = 作業フォルダ / "入力 ブック"
    入力フォルダ.mkdir()
    平文データ = ブックを生成(入力フォルダ / "VBAなし.xlsx", プロジェクト一覧=0)
    名前付きブック = ブックを生成(
        入力フォルダ / "名前付き.xlsx",
        プロジェクト一覧=0,
        名前定義XML='<definedNames><definedName name="売上範囲" localSheetId="0" hidden="1">\'売上表\'!$B$2:$B$20</definedName></definedNames>',
    )
    一部成功ファイル = ブックを生成(作業フォルダ / "一部.xlsm", モジュールを欠損=True)
    暗号化データ = 作業フォルダ / "暗号化.xlsm"
    暗号化データ.write_bytes(
        複合ファイルを生成({"EncryptionInfo": bytes(4096), "EncryptedPackage": bytes(4096)})
    )
    未対応ファイル = 作業フォルダ / "未対応.txt"
    未対応ファイル.write_bytes(b"plain text")
    破損ファイル = 作業フォルダ / "破損.xlsm"
    破損ファイル.write_bytes(b"PK\3\4broken")
    旧方式 = 作業フォルダ / "旧方式.bin"
    旧方式.write_bytes(複合ファイルを生成(プロジェクトストリームを生成(パスワード="test", 旧方式=True)))
    Wordファイル = Word文書を生成(作業フォルダ / "文書.docm", 条件付きコンパイル引数="Test = -1")
    Accessファイル = Accessデータベースを生成(
        作業フォルダ / "データ.accdb", 条件付きコンパイル引数="Test = -1"
    )
    Access平文 = Accessデータベースを生成(作業フォルダ / "VBAなし.accdb", VBAあり=False)
    入力控え = {
        p: p.read_bytes()
        for p in (
            マクロファイル,
            平文データ,
            名前付きブック,
            一部成功ファイル,
            暗号化データ,
            未対応ファイル,
            破損ファイル,
            旧方式,
            Wordファイル,
            Accessファイル,
            Access平文,
        )
    }
    環境変数 = os.environ.copy()
    Windowsフォルダ = os.environ["SystemRoot"]
    環境変数["PATH"] = os.pathsep.join([str(Path(Windowsフォルダ) / "System32"), Windowsフォルダ])
    for 名前 in ("PYTHONHOME", "PYTHONPATH", "TCL_LIBRARY", "TK_LIBRARY"):
        環境変数.pop(名前, None)
    結果行一覧 = []

    def 実行(表示ラベル, 引数一覧, 期待値, 入力テキスト=None):
        実行結果 = subprocess.run(
            [str(実行ファイル), *map(str, 引数一覧)],
            cwd=作業フォルダ,
            env=環境変数,
            capture_output=True,
            timeout=40,
            creationflags=subprocess.CREATE_NO_WINDOW,
            input=入力テキスト.encode("utf-8") if 入力テキスト is not None else None,
        )
        assert 実行結果.returncode == 期待値, (表示ラベル, 実行結果.returncode, 実行結果.stderr)
        結果行一覧.append(
            {
                "確認項目": 表示ラベル,
                "終了コード": 実行結果.returncode,
                "標準出力": 実行結果.stdout.decode("utf-8", errors="replace"),
                "標準エラー": 実行結果.stderr.decode("utf-8", errors="replace"),
            }
        )

    出力先 = 作業フォルダ / "結果"
    実行("macro", [マクロファイル, "-o", 出力先, "--no-ui"], 0)
    最初の結果 = next(出力先.iterdir())
    実行("macro_debug", [マクロファイル, "-o", 出力先, "--no-ui", "--debug-log"], 0)
    二番目の結果 = next(p for p in 出力先.iterdir() if p != 最初の結果)
    assert 比較内容を取得(最初の結果) == 比較内容を取得(二番目の結果)
    実行("beside_input_no_output", [平文データ], 0)
    assert len(list(入力フォルダ.glob("VBAなし.xlsx_*"))) == 1
    assert not list(作業フォルダ.glob("VBAなし.xlsx_*"))
    実行("no_vba", [平文データ, "-o", 出力先, "--no-ui"], 0)
    実行("defined_names_no_vba", [名前付きブック, "-o", 出力先, "--no-ui"], 0)
    名前情報 = json.loads(
        next(出力先.glob("名前付き.xlsx_*/設定情報/名前定義一覧.json")).read_text(encoding="utf-8")
    )
    assert 名前情報["取得状態"] == "設定あり"
    assert 名前情報["名前定義一覧"][0]["名前"] == "売上範囲"
    assert 名前情報["名前定義一覧"][0]["参照先の保存式"] == "'売上表'!$B$2:$B$20"
    assert 名前情報["名前定義一覧"][0]["非表示フラグ（保存値）"] is True
    実行("partial", [一部成功ファイル, "-o", 出力先, "--no-ui"], 3)
    実行("encrypted", [暗号化データ, "-o", 出力先, "--no-ui"], 1)
    実行("unsupported", [未対応ファイル, "-o", 出力先, "--no-ui"], 1)
    実行("corrupt", [破損ファイル, "-o", 出力先, "--no-ui", "--debug-log"], 1)
    実行("legacy_candidates_unset", [旧方式, "-o", 出力先, "--no-ui"], 0)
    for 表示ラベル, 対象 in (("word", Wordファイル), ("access", Accessファイル)):
        実行(表示ラベル, [対象, "-o", 出力先, "--no-ui"], 0)
        保存先 = next(出力先.glob(対象.name + "_*"))
        引数 = json.loads((保存先 / "設定情報/条件付きコンパイル引数.json").read_text(encoding="utf-8"))
        assert 引数["設定一覧"] == [{"名前": "Test", "設定値": -1}]
        参照 = json.loads((保存先 / "設定情報/参照設定.json").read_text(encoding="utf-8"))
        assert 参照["参照設定一覧"][0]["名前"] == "Alpha"
        assert list(保存先.rglob("*.bas")) and list(保存先.rglob("*.cls"))
    実行("access_no_vba", [Access平文, "-o", 出力先, "--no-ui"], 0)
    実行("multiple_targets", [マクロファイル, 平文データ, "--no-ui"], 2)
    実行("missing_target", ["--no-ui"], 2)
    実行("not_found", ["missing.xlsm", "--no-ui"], 1)
    実行("output_failure", [平文データ, "-o", 未対応ファイル / "child", "--no-ui"], 1)
    実行("help", ["--help"], 0)
    実行("version", ["--version"], 0)
    実行("interactive_questions", [], 0, f'"{マクロファイル}"\n"{出力先}"\nn\n')
    起動表示 = 結果行一覧[-1]["標準出力"]
    assert "対話モード" not in 起動表示
    assert "【対応ファイル】" in 起動表示 and ".xlsx .xlsm" in 起動表示
    assert "Word" in 起動表示 and "Access" in 起動表示
    assert "未対応" not in 起動表示 and "引用符付き" not in 起動表示
    assert "形式ごとの取得範囲" not in 起動表示
    assert "\x1b" not in 起動表示
    実行("interactive_cancel", [], 130, "q\n")
    # 標準ストリームのない分離起動でも、画面を作らず終了できることを確認する。
    分離実行結果 = subprocess.run(
        [str(実行ファイル), str(平文データ), "-o", str(出力先)],
        cwd=作業フォルダ,
        env=環境変数,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.DETACHED_PROCESS,
        timeout=40,
    )
    assert 分離実行結果.returncode == 0
    assert all(パス.read_bytes() == データ for パス, データ in 入力控え.items())
    レポート = {
        "確認結果一覧": 結果行一覧,
        "標準ストリームなしの分離起動": "成功",
        "入力内容の不変": True,
        "詳細ログ有無で比較用ファイル一致": True,
        "EXEのSHA256": hashlib.sha256(実行ファイル.read_bytes()).hexdigest(),
    }
    (作業フォルダ / "verification.json").write_text(
        json.dumps(レポート, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"{len(結果行一覧) + 1} EXE checks passed: {作業フォルダ}")


if __name__ == "__main__":
    メイン()
