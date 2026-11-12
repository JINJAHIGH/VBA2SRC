"""起動案内、色を使えない出力先、ポップアップなしの実行を検証する。"""

import io
import sys

from tests.samples import ブックを生成
from vba2src_app import cli, ui


def test_banner_has_shared_ascii_title_and_honest_format_support(差替え):
    出力 = io.StringIO()
    差替え.setattr(sys, "stdout", 出力)
    ui.起動案内を表示()
    表示 = 出力.getvalue()
    assert "対話モード" not in 表示
    assert ui.files("vba2src_app").joinpath("title.txt").read_text(encoding="ascii").strip() in 表示
    assert "Word" in 表示 and "Access" in 表示 and "PowerPoint" not in 表示
    assert ".docx .docm .dotx .dotm .doc .dot" in 表示
    assert ".accdb .mdb" in 表示
    assert "形式ごとの取得範囲" not in 表示
    assert ".xlsx .xlsm" in 表示
    assert "\x1b" not in 表示


def test_questions_only_show_necessary_steps(差替え):
    出力 = io.StringIO()
    差替え.setattr(sys, "stdout", 出力)
    差替え.setattr("builtins.input", lambda 質問: "q")
    assert cli.メイン([]) == 130
    assert "引用符付き" not in 出力.getvalue()
    assert "01  対象ファイル" in 出力.getvalue()


def test_redirected_success_and_failure_have_readable_results_without_ansi(一時フォルダ, 差替え):
    出力, エラー = io.StringIO(), io.StringIO()
    差替え.setattr(sys, "stdout", 出力)
    差替え.setattr(sys, "stderr", エラー)
    差替え.setitem(sys.modules, "tkinter", None)
    対象 = ブックを生成(一時フォルダ / "確認 ブック.xlsm")
    assert cli.メイン([str(対象), "-o", str(一時フォルダ / "結果")]) == 0
    assert "抽出完了" in 出力.getvalue()
    assert "取得モジュール数：5" in 出力.getvalue()
    assert "最初に確認するファイル" in 出力.getvalue()
    assert cli.メイン(["missing.xlsm"]) == 1
    assert "INPUT_NOT_FOUND" in エラー.getvalue()
    assert "\x1b" not in 出力.getvalue() + エラー.getvalue()


def test_no_color_disables_escape_codes_even_for_terminal(差替え):
    class 端末出力(io.StringIO):
        def isatty(自身):
            return True

    出力 = 端末出力()
    差替え.setattr(sys, "stdout", 出力)
    差替え.setenv("NO_COLOR", "")
    ui.色付き表示("結果：成功", "成功")
    assert 出力.getvalue() == "結果：成功\n"


def test_partial_result_is_visible_and_distinct_from_success(一時フォルダ, 差替え):
    出力 = io.StringIO()
    差替え.setattr(sys, "stdout", 出力)
    対象 = ブックを生成(一時フォルダ / "部分.xlsm", モジュールを欠損=True)
    assert cli.メイン([str(対象), "--no-ui"]) == 3
    assert "一部の情報を取得できませんでした" in 出力.getvalue()
    assert "結果：一部成功" in 出力.getvalue()
    assert "診断コードと実行ログ" in 出力.getvalue()
