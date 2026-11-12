import hashlib
import io
import json
import zipfile

import olefile
import pytest
from oletools.olevba import VBA_Project

from tests.samples import ブックを生成, プロジェクトストリームを生成, 複合ファイルを生成
from vba2src_app.cli import メイン
from vba2src_app.extract import 抽出
from vba2src_app.output import 抽出結果を保存


def 比較スナップショット(フォルダ):
    return {
        p.relative_to(フォルダ).as_posix(): p.read_bytes()
        for p in フォルダ.rglob("*")
        if p.is_file() and "ログ" not in p.parts
    }


def test_all_types_code_metadata_forms_and_readonly(一時フォルダ):
    パス = ブックを生成(一時フォルダ / "対象 ファイル (テスト).xlsm")
    処理前データ = hashlib.sha256(パス.read_bytes()).digest()
    結果 = 抽出(パス)
    assert 結果.終了コード == 0
    assert 結果.VBA有無 == "あり"
    対象プロジェクト = 結果.プロジェクト一覧[0]
    assert {m.種類 for m in 対象プロジェクト.モジュール一覧} == {"standard", "class", "document", "form"}
    対象モジュール = next(m for m in 対象プロジェクト.モジュール一覧 if m.種類 == "standard")
    assert "\t' 日本語コメント\n" in 対象モジュール.コード
    assert 対象モジュール.コード.endswith("\n\n")
    assert "\r" not in 対象モジュール.コード
    assert "#Else\n" in 対象モジュール.コード
    assert 対象モジュール.属性情報 == ['Attribute VB_Name = "集計処理"']
    assert 対象プロジェクト.条件付きコンパイル引数["items"] == [
        {"name": "Alpha", "value": 3},
        {"name": "z", "value": -1},
    ]
    assert [r["name"] for r in 対象プロジェクト.参照一覧] == ["Alpha", "alpha", "zeta"]
    assert [r["original_order"] for r in 対象プロジェクト.参照一覧] == [1, 2, 0]
    assert 対象プロジェクト.参照一覧[0]["stored_path"] == r"C:\Library\Alpha.tlb"
    assert 対象プロジェクト.シート一覧[0]["display_name"] == "売上表"
    assert 対象プロジェクト.シート一覧[0]["codename"] == "Sheet1"
    フォーム = 対象プロジェクト.フォーム一覧[0]
    assert フォーム["properties"]["Caption"] == "確認フォーム"
    assert フォーム["controls"][0]["name"] == "ラベル1"
    assert フォーム["controls"][0]["properties"] == {"Caption": "表示文字列"}
    assert フォーム["controls"][0]["parent"] == "UserForm1"
    assert フォーム["complete_design"] is False
    assert 対象プロジェクト.モジュール取得状態 == "取得済み" and 対象プロジェクト.保存モジュール数 == 5
    出力先 = 一時フォルダ / "out"
    抽出結果を保存(出力先, 結果)
    assert len(list(出力先.rglob("*.bas"))) == 1
    assert len(list(出力先.rglob("*.cls"))) == 3
    assert len(list(出力先.rglob("*.code.txt"))) == 1
    assert not list(出力先.rglob("*.frm")) and not list(出力先.rglob("*.frx"))
    for ファイル名, データ in 比較スナップショット(出力先).items():
        assert not データ.startswith(b"\xef\xbb\xbf")
        assert b"\r" not in データ
        assert str(一時フォルダ).encode() not in データ
        if ファイル名.endswith(".json"):
            assert json.loads(データ)["スキーマバージョン"] == 2
    assert hashlib.sha256(パス.read_bytes()).digest() == 処理前データ


def test_oletools_api_agrees_with_synthetic_source():
    データ = 複合ファイルを生成(プロジェクトストリームを生成(フォームあり=False))
    with olefile.OleFileIO(io.BytesIO(データ)) as OLEファイル:
        外部API = VBA_Project(OLEファイル, "", "PROJECT", "VBA/dir", relaxed=False)
        外部API.parse_project_stream()
        assert len(list(外部API.parse_modules())) == 4
        assert "日本語コメント" in 外部API.modules[0].code_raw.decode("cp932")


@pytest.mark.parametrize("詳細ログ", [False, True])
def test_cli_determinism_debug_and_no_ui(一時フォルダ, 差替え, 詳細ログ):
    import sys

    パス = ブックを生成(一時フォルダ / "比較対象.xlsm", パスワード="確認パス")
    差替え.setitem(sys.modules, "tkinter", None)
    解析済み引数 = [str(パス), "-o", str(一時フォルダ / "結果"), "--no-ui"]
    assert メイン(解析済み引数) == 0
    assert メイン(解析済み引数 + (["--debug-log"] if 詳細ログ else [])) == 0
    フォルダ一覧 = sorted((一時フォルダ / "結果").iterdir())
    assert len(フォルダ一覧) == 2
    assert 比較スナップショット(フォルダ一覧[0]) == 比較スナップショット(フォルダ一覧[1])
    assert (フォルダ一覧[0] / "ログ/実行ログ.txt").exists()


def test_multi_project_separation_and_report_order(一時フォルダ):
    結果 = 抽出(ブックを生成(一時フォルダ / "multi.xlsm", プロジェクト一覧=2))
    assert 結果.終了コード == 0
    assert len(結果.プロジェクト一覧) == 2
    出力先 = 一時フォルダ / "out"
    抽出結果を保存(出力先, 結果)
    assert len(list((出力先 / "プロジェクト").iterdir())) == 2
    レポート = (出力先 / "抽出結果一覧.txt").read_text(encoding="utf-8")
    assert レポート.rfind("[07]") < レポート.index("[08]")
    assert レポート.count("[08]") == 2
    for 対象プロジェクト in 結果.プロジェクト一覧:
        assert all(
            m.保存パス.startswith(f"プロジェクト/{対象プロジェクト.識別子}/")
            for m in 対象プロジェクト.モジュール一覧
        )


def test_cli_group_by_type_preserves_sources_and_distinguishes_projects(一時フォルダ):
    パス = ブックを生成(一時フォルダ / "まとめ.xlsm", プロジェクト一覧=2)
    通常出力先 = 一時フォルダ / "通常"
    追加出力先 = 一時フォルダ / "追加"
    assert メイン([str(パス), "-o", str(通常出力先), "--no-ui"]) == 0
    assert メイン([str(パス), "-o", str(追加出力先), "--no-ui", "--group-by-type"]) == 0
    通常結果 = next(通常出力先.iterdir())
    追加結果 = next(追加出力先.iterdir())
    assert not (通常結果 / "VBAまとめ").exists()
    assert 比較スナップショット(通常結果) == {
        名前: データ for 名前, データ in 比較スナップショット(追加結果).items() if not 名前.startswith("VBAまとめ/")
    }
    種類別ファイル名 = {
        "standard": "標準モジュール.txt",
        "class": "クラスモジュール.txt",
        "document": "ドキュメントモジュール.txt",
        "form": "フォーム.txt",
    }
    まとめ先 = 追加結果 / "VBAまとめ"
    assert {p.name for p in まとめ先.iterdir()} == set(種類別ファイル名.values())
    プロジェクト一覧 = 抽出(パス).プロジェクト一覧
    for 種類, ファイル名 in 種類別ファイル名.items():
        期待内容 = "".join(
            f"' ===== モジュール：{m.名前}（プロジェクト：{p.識別子}） =====\n{m.コード}\n"
            for p in プロジェクト一覧
            for m in p.モジュール一覧
            if m.種類 == 種類
        )
        assert (まとめ先 / ファイル名).read_bytes() == 期待内容.encode("utf-8")


def test_cli_group_by_type_includes_unclassified_code_and_omits_unreadable_code(一時フォルダ):
    パス = ブックを生成(
        一時フォルダ / "partial.xlsm",
        フォームあり=False,
        モジュールを欠損=True,
        モジュール仕様=[("欠損", "Module"), ("分類不能", "Undefined")],
    )
    出力先 = 一時フォルダ / "結果"
    assert メイン([str(パス), "-o", str(出力先), "--group-by-type", "--no-ui"]) == 3
    まとめ先 = next(出力先.iterdir()) / "VBAまとめ"
    assert [p.name for p in まとめ先.iterdir()] == ["未分類.txt"]
    assert 'Attribute VB_Name = "分類不能"' in (まとめ先 / "未分類.txt").read_text(encoding="utf-8")


@pytest.mark.parametrize("破損状態,期待値", [(False, 0), (True, 3)])
def test_success_and_partial_input_unchanged(一時フォルダ, 破損状態, 期待値):
    パス = ブックを生成(一時フォルダ / "partial.xlsm", モジュールを欠損=破損状態)
    処理前データ = パス.read_bytes()
    結果 = 抽出(パス)
    assert 結果.終了コード == 期待値
    assert sum(m.コード is not None for m in 結果.プロジェクト一覧[0].モジュール一覧) == (
        4 if 破損状態 else 5
    )
    assert パス.read_bytes() == 処理前データ


def test_no_vba_and_constants_unset(一時フォルダ):
    平文データ = 抽出(ブックを生成(一時フォルダ / "plain.xlsx", プロジェクト一覧=0))
    assert 平文データ.終了コード == 0 and 平文データ.VBA有無 == "なし"
    抽出結果を保存(一時フォルダ / "plain", 平文データ)
    レポート = (一時フォルダ / "plain/抽出結果一覧.txt").read_text(encoding="utf-8")
    assert "候補照合：対象外" in レポート
    結果 = 抽出(ブックを生成(一時フォルダ / "unset.xlsm", 条件付きコンパイル引数=""))
    assert 結果.終了コード == 0
    assert 結果.プロジェクト一覧[0].条件付きコンパイル引数["status"] == "設定なし"


@pytest.mark.parametrize(
    "種類", ["encrypted", "corrupt", "unsupported", "missing", "missing_vba_part", "truncated_dir"]
)
def test_errors_are_not_no_vba(一時フォルダ, 種類):
    パス = 一時フォルダ / "input.xlsm"
    if 種類 == "encrypted":
        パス.write_bytes(複合ファイルを生成({"EncryptedPackage": bytes(4096), "EncryptionInfo": bytes(4096)}))
    elif 種類 == "corrupt":
        パス.write_bytes(b"PK\x03\x04broken")
    elif 種類 == "unsupported":
        パス.write_bytes(b"not Office")
    elif 種類 == "missing_vba_part":
        ブックを生成(パス)
        with zipfile.ZipFile(パス, "a") as アーカイブ:
            アーカイブ.writestr("xl/vba0.bin", b"broken")  # 重複パーツ自体が構造不整合となる。
    elif 種類 == "truncated_dir":
        ストリーム一覧 = プロジェクトストリームを生成()
        ストリーム一覧["VBA/dir"] = b"\1broken".ljust(4096, b"\0")
        パス.write_bytes(複合ファイルを生成(ストリーム一覧))
    処理前データ = パス.read_bytes() if パス.exists() else None
    結果 = 抽出(パス)
    assert 結果.終了コード in (1, 3)
    assert 結果.VBA有無 != "なし"
    if 処理前データ is not None:
        assert パス.read_bytes() == 処理前データ


def test_sources_survive_corrupt_protection(一時フォルダ):
    ストリーム一覧 = プロジェクトストリームを生成()
    対象プロジェクト = ストリーム一覧["PROJECT"].replace(b'DPB="', b'DPB="ZZ')
    ストリーム一覧["PROJECT"] = 対象プロジェクト
    パス = 一時フォルダ / "bad.bin"
    パス.write_bytes(複合ファイルを生成(ストリーム一覧))
    結果 = 抽出(パス)
    assert 結果.終了コード == 3
    assert all(m.コード for m in 結果.プロジェクト一覧[0].モジュール一覧)
    assert 結果.プロジェクト一覧[0].保護情報["password_set"] is None


def test_unread_module_directory_is_not_reported_as_zero_modules(一時フォルダ):
    ストリーム一覧 = プロジェクトストリームを生成()
    ストリーム一覧["VBA/dir"] = b"\1broken".ljust(4096, b"\0")
    パス = 一時フォルダ / "broken.bin"
    パス.write_bytes(複合ファイルを生成(ストリーム一覧))
    結果 = 抽出(パス)
    抽出結果を保存(一時フォルダ / "out", 結果)
    レポート = (一時フォルダ / "out/抽出結果一覧.txt").read_text(encoding="utf-8")
    assert "標準モジュール数：不明（取得済み 0）" in レポート
    assert "未分類モジュール数：" in レポート


def test_required_failure_is_partial_but_expected_limitations_are_not(一時フォルダ):
    正常結果 = 抽出(ブックを生成(一時フォルダ / "good.xlsm"))
    assert 正常結果.プロジェクト一覧[0].フォーム一覧[0]["complete_design"] is False
    assert 正常結果.終了コード == 0
    ストリーム一覧 = プロジェクトストリームを生成()
    ストリーム一覧["UserForm1/f"] = bytes(4096)
    パス = 一時フォルダ / "formfail.bin"
    パス.write_bytes(複合ファイルを生成(ストリーム一覧))
    異常結果 = 抽出(パス)
    assert 異常結果.終了コード == 3 and "FORM_READ_FAILED" in 異常結果.プロジェクト一覧[0].失敗一覧


def test_debug_logs_do_not_contain_password_candidates(一時フォルダ, 差替え):
    from vba2src_app import cli

    パスワード = "一致 パス\n"
    対象 = ブックを生成(一時フォルダ / "secret.xlsm", パスワード=パスワード, モジュールを欠損=True)
    差替え.setattr(cli.config, "VBA保護パスワード候補", ("不一致の秘密候補", パスワード))
    assert メイン([str(対象), "-o", str(一時フォルダ / "out"), "--no-ui", "--debug-log"]) == 3
    フォルダ = next((一時フォルダ / "out").iterdir())
    ログ内容 = (フォルダ / "ログ/実行ログ.txt").read_text(encoding="utf-8")
    assert "不一致の秘密候補" not in ログ内容 and パスワード not in ログ内容
    assert "Traceback" not in ログ内容 or "exception=" in ログ内容
    assert b"\r" not in (フォルダ / "ログ/実行ログ.txt").read_bytes()
    レポート = (フォルダ / "抽出結果一覧.txt").read_text(encoding="utf-8")
    assert "パスワード：一致 パス\\u000a\n" in レポート
