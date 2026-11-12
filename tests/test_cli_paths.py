import io
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import pytest

from tests.samples import ブックを生成
from vba2src_app import cli, paths


def test_atomic_same_second_collision(一時フォルダ):
    現在時刻 = datetime(2026, 11, 8, 9, 42, 33)
    with ThreadPoolExecutor(max_workers=8) as 並列実行器:
        フォルダ一覧 = list(
            並列実行器.map(
                lambda 未使用: paths.結果フォルダを作成(一時フォルダ, "対象.xlsx", 現在時刻), range(20)
            )
        )
    assert len(set(フォルダ一覧)) == 20
    assert {p.name for p in フォルダ一覧} == {"対象.xlsx_20261108_094233"} | {
        f"対象.xlsx_20261108_094233_{n:03d}" for n in range(1, 20)
    }


def test_default_parent_frozen_and_cwd(一時フォルダ, 差替え):
    ツール = paths.ツールフォルダを取得()
    差替え.chdir(一時フォルダ)
    assert paths.既定親フォルダを取得(None) == ツール
    assert paths.既定親フォルダを取得("") == ツール
    assert paths.既定親フォルダを取得("does-not-exist") == ツール
    assert paths.既定親フォルダを取得(str(一時フォルダ)) == 一時フォルダ
    差替え.setattr(sys, "frozen", True, raising=False)
    差替え.setattr(sys, "executable", str(一時フォルダ / "VBA2SRC.exe"))
    assert paths.ツールフォルダを取得() == 一時フォルダ
    assert "_MEI" not in str(paths.ツールフォルダを取得())


def test_name_safety_and_case_collision():
    assert paths.安全な保存名("CON.txt") == "_CON.txt"
    assert paths.安全な保存名("a:b?. ") == "a_b_"
    割当器 = paths.保存名割当器()
    assert 割当器.割当("Foo", ".bas") == "Foo.bas"
    assert 割当器.割当("foo", ".bas") == "foo_001.bas"
    assert 割当器.割当("a/b", ".bas") == "a_b.bas"
    assert 割当器.割当("a:b", ".bas") == "a_b_001.bas"


@pytest.mark.parametrize("画面なし", [False, True])
def test_target_only_saves_beside_input_without_asking(一時フォルダ, 差替え, 画面なし):
    入力フォルダ = 一時フォルダ / "入力 ブック"
    入力フォルダ.mkdir()
    パス = ブックを生成(入力フォルダ / "対象 (1).xlsm", プロジェクト一覧=0)
    起動フォルダ = 一時フォルダ / "起動場所"
    起動フォルダ.mkdir()
    設定保存先 = 一時フォルダ / "設定保存先"
    設定保存先.mkdir()
    差替え.chdir(起動フォルダ)
    差替え.setattr(cli.config, "既定保存先親フォルダ", str(設定保存先))
    差替え.setattr("builtins.input", lambda *未使用: pytest.fail("保存先を尋ねてはいけない"))
    差替え.setitem(sys.modules, "tkinter", None)
    assert cli.メイン([str(パス), *(["--no-ui"] if 画面なし else [])]) == 0
    assert len(list(入力フォルダ.glob("対象 (1).xlsm_*"))) == 1
    assert not list(設定保存先.iterdir()) and not list(起動フォルダ.iterdir())


def test_explicit_output_overrides_input_parent(一時フォルダ, 差替え):
    パス = ブックを生成(一時フォルダ / "in.xlsm", プロジェクト一覧=0)
    差替え.setitem(sys.modules, "tkinter", None)
    assert cli.メイン([str(パス), "-o", str(一時フォルダ / "explicit")]) == 0
    assert len(list((一時フォルダ / "explicit").glob("in.xlsm_*"))) == 1
    assert not list(一時フォルダ.glob("in.xlsm_*"))


def test_no_ui_even_without_streams_or_tkinter(一時フォルダ, 差替え):
    パス = ブックを生成(一時フォルダ / "in.xlsx", プロジェクト一覧=0)
    差替え.setattr(sys, "stdout", None)
    差替え.setattr(sys, "stderr", None)
    差替え.setitem(sys.modules, "tkinter", None)
    assert cli.メイン([str(パス), "-o", str(一時フォルダ / "out"), "--no-ui"]) == 0
    assert cli.メイン(["missing", "--no-ui"]) == 1
    assert cli.メイン(["a", "b", "--no-ui"]) == 2
    assert cli.メイン(["--help"]) == 0
    assert cli.メイン(["--version"]) == 0


def test_output_failure_and_relative_paths(一時フォルダ, 差替え):
    パス = ブックを生成(一時フォルダ / "入力 ファイル (1).xlsx", プロジェクト一覧=0)
    差替え.chdir(一時フォルダ)
    assert cli.メイン([パス.name, "-o", "親/子", "--no-ui"]) == 0
    使用不可パス = 一時フォルダ / "blocked"
    使用不可パス.write_text("file", encoding="utf-8")
    assert cli.メイン([パス.name, "-o", str(使用不可パス / "child"), "--no-ui"]) == 1
    差替え.setattr(
        paths.Path, "mkdir", lambda *a, **追加キーワード: (未使用 for 未使用 in ()).throw(PermissionError())
    )
    assert cli.メイン([パス.name, "-o", "denied", "--no-ui"]) == 1


def test_argument_errors_and_stderr(一時フォルダ, 差替え):
    ストリーム = io.StringIO()
    差替え.setattr(sys, "stderr", ストリーム)
    assert cli.メイン(["--no-ui"]) == 2
    assert "引数エラー" in ストリーム.getvalue()
    assert cli.メイン(["a", "b", "--no-ui"]) == 2
    assert cli.メイン(["a", "--unknown", "--no-ui"]) == 2


def test_fresh_dependency_import_without_standard_streams(一時フォルダ):
    パス = ブックを生成(一時フォルダ / "fresh.xlsm")
    スクリプト = """
import sys
sys.stdout = sys.stderr = sys.stdin = None
from vba2src_app.cli import メイン
raise SystemExit(メイン(sys.argv[1:]))
"""
    実行結果 = subprocess.run(
        [sys.executable, "-c", スクリプト, str(パス), "-o", str(一時フォルダ / "result"), "--no-ui"],
        capture_output=True,
        timeout=30,
    )
    assert 実行結果.returncode == 0, 実行結果.stderr
    assert len(list((一時フォルダ / "result").rglob("*.bas"))) == 1


@pytest.mark.parametrize("失敗種別", ["write", "close"])
def test_log_write_failure_returns_error_without_secondary_exception(一時フォルダ, 差替え, 失敗種別):
    対象 = ブックを生成(一時フォルダ / "input.xlsx", プロジェクト一覧=0)

    class 切断ストリーム(io.StringIO):
        def write(自身, 値):
            if 失敗種別 == "write":
                raise OSError("simulated disconnected share")
            return super().write(値)

        def close(自身):
            super().close()
            if 失敗種別 == "close":
                raise OSError("simulated disconnected share at close")

    差替え.setattr(cli.LFログハンドラー, "_open", lambda 未使用: 切断ストリーム())
    assert cli.メイン([str(対象), "-o", str(一時フォルダ / "out"), "--no-ui"]) == 1
    assert len(list((一時フォルダ / "out").rglob("抽出結果一覧.txt"))) == 1


def test_input_permission_is_not_no_vba(一時フォルダ, 差替え):
    from vba2src_app.extract import 抽出

    パス = ブックを生成(一時フォルダ / "denied.xlsx", プロジェクト一覧=0)
    差替え.setattr(
        paths.Path, "read_bytes", lambda 未使用: (未使用 for 未使用 in ()).throw(PermissionError())
    )
    結果 = 抽出(パス)
    assert 結果.終了コード == 1 and 結果.VBA有無 == "判定不能"
    assert 結果.失敗一覧 == ["INPUT_ACCESS_DENIED"]
