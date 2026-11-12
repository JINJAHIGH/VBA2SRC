import io
import os
import subprocess
import sys
from pathlib import Path, PureWindowsPath

import pytest

from tests.samples import ブックを生成
from vba2src_app import cli


def 回答を設定(差替え, 値一覧):
    反復子 = iter(値一覧)
    差替え.setattr("builtins.input", lambda 質問: next(反復子))


def test_interactive_retries_quotes_default_and_debug(一時フォルダ, 差替え):
    対象 = ブックを生成(一時フォルダ / "確認 用 (1).xlsm")
    使用不可パス = 一時フォルダ / "not-a-folder"
    使用不可パス.write_text("file", encoding="utf-8")
    差替え.setattr(cli.config, "既定保存先親フォルダ", str(一時フォルダ))
    差替え.setitem(sys.modules, "tkinter", None)
    回答を設定(差替え, ["", "missing.xlsm", f'"{対象}"', str(使用不可パス), "", "maybe", "y"])
    assert cli.メイン([]) == 0
    フォルダ一覧 = list(一時フォルダ.glob("確認 用 (1).xlsm_*"))
    assert len(フォルダ一覧) == 1
    assert "DEBUG" in (フォルダ一覧[0] / "ログ/実行ログ.txt").read_text(encoding="utf-8")


def test_interactive_cancel_and_end_of_input(差替え):
    回答を設定(差替え, ["q"])
    assert cli.メイン([]) == 130
    差替え.setattr("builtins.input", lambda 未使用: (未使用 for 未使用 in ()).throw(EOFError()))
    assert cli.メイン([]) == 130


def test_no_ui_with_missing_target_never_prompts(差替え):
    差替え.setattr("builtins.input", lambda 未使用: pytest.fail("--no-ui must never ask questions"))
    assert cli.メイン(["--no-ui"]) == 2


def test_interactive_keeps_mapped_and_unc_paths(差替え):
    # OSに依存せず、UNCパスとドライブ表記の保持を検証する。
    for 元データ in (r'"Y:\部門 共通\対象 (1).xlsm"', r'"\\server\share\部門 共通\対象.xlsm"'):
        回答を設定(差替え, [元データ])
        値 = cli.回答を取得("path: ")
        assert 値 == 元データ[1:-1]
        assert PureWindowsPath(値).is_absolute()
    出力 = io.StringIO()
    差替え.setattr(sys, "stdout", 出力)
    cli.安全に表示(r"\\server\share\保存先")
    assert 出力.getvalue() == "\\\\server\\share\\保存先\n"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows drive alias test")
def test_mapped_drive_input_and_output(一時フォルダ):
    # SUBSTでドライブ経由の入出力を検証する。SMB通信は対象外。
    import ctypes

    ビットマスク = ctypes.windll.kernel32.GetLogicalDrives()
    ドライブ文字 = next((chr(n + 65) for n in range(25, 3, -1) if not ビットマスク & (1 << n)), None)
    if ドライブ文字 is None:
        pytest.skip("No unused drive letter")
    ドライブ割当コマンド = Path(os.environ["SystemRoot"]) / "System32/subst.exe"
    ブックを生成(一時フォルダ / "割当 ドライブ.xlsm")
    subprocess.run(
        [str(ドライブ割当コマンド), ドライブ文字 + ":", str(一時フォルダ)], check=True, capture_output=True
    )
    try:
        assert (
            cli.メイン([f"{ドライブ文字}:\\割当 ドライブ.xlsm", "-o", f"{ドライブ文字}:\\出力 親", "--no-ui"])
            == 0
        )
        assert len(list((一時フォルダ / "出力 親").rglob("*.bas"))) == 1
    finally:
        subprocess.run([str(ドライブ割当コマンド), ドライブ文字 + ":", "/D"], check=True, capture_output=True)
