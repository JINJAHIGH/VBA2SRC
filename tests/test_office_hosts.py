"""Word・AccessのVBA、設定、破損検出と読み取り専用性を検証する。"""

import hashlib
import json
import struct
import zipfile

import pytest
from pyopenvba.access.database import AccessDatabase

from tests.samples import (
    Accessデータベースを生成,
    Word文書を生成,
    プロジェクトストリームを生成,
    圧縮データを生成,
    複合ファイルを生成,
)
from tests.test_extraction import 比較スナップショット
from vba2src_app.cli import メイン
from vba2src_app.extract import 抽出
from vba2src_app.output import 抽出結果を保存


@pytest.mark.parametrize(
    "生成器,拡張子,ホスト",
    [(Word文書を生成, ".docm", "Word"), (Accessデータベースを生成, ".accdb", "Access")],
)
def test_hosts_export_code_constants_references_protection_and_readonly(一時フォルダ, 生成器, 拡張子, ホスト):
    パス = 生成器(
        一時フォルダ / ("入力" + 拡張子), 条件付きコンパイル引数="Debug = -1 : Release = 2", パスワード="確認"
    )
    元データ = パス.read_bytes()
    結果 = 抽出(パス, ("確認",))
    assert 結果.終了コード == 0 and 結果.VBA有無 == "あり"
    assert 結果.形式.startswith(ホスト)
    assert 結果.名前定義情報["status"] == f"対象外（{ホスト}）"
    プロジェクト = 結果.プロジェクト一覧[0]
    assert プロジェクト.シート取得状態 == f"対象外（{ホスト}）"
    assert プロジェクト.条件付きコンパイル引数["items"] == [
        {"name": "Debug", "value": -1},
        {"name": "Release", "value": 2},
    ]
    assert プロジェクト.保護情報["password"] == "確認"
    assert プロジェクト.参照取得状態 == "取得済み"
    assert [r["name"] for r in プロジェクト.参照一覧] == ["Alpha", "alpha", "zeta"]
    assert プロジェクト.参照一覧[0]["stored_path"] == r"C:\Library\Alpha.tlb"
    assert {m.種類 for m in プロジェクト.モジュール一覧} >= {"standard", "class", "document"}
    assert all(m.コード and "日本語コメント" in m.コード for m in プロジェクト.モジュール一覧)
    if ホスト == "Access":
        assert {m.名前 for m in プロジェクト.モジュール一覧 if m.種類 == "document"} == {
            "Form_Orders",
            "Report_Summary",
        }
    出力先 = 一時フォルダ / "out"
    抽出結果を保存(出力先, 結果)
    assert list(出力先.rglob("*.bas")) and list(出力先.rglob("*.cls"))
    assert (
        json.loads((出力先 / "設定情報/条件付きコンパイル引数.json").read_text(encoding="utf-8"))["元の値"]
        == "Debug = -1 : Release = 2"
    )
    assert パス.read_bytes() == 元データ
    assert メイン([str(パス), "-o", str(一時フォルダ / "cli"), "--no-ui"]) == 0


@pytest.mark.parametrize(
    "生成器,名前",
    [
        (Word文書を生成, "plain.docx"),
        (Word文書を生成, "plain.dotx"),
        (Accessデータベースを生成, "plain.accdb"),
    ],
)
def test_no_vba_is_success(一時フォルダ, 生成器, 名前):
    結果 = 抽出(生成器(一時フォルダ / 名前, VBAあり=False))
    assert 結果.終了コード == 0 and 結果.VBA有無 == "なし"
    assert 結果.名前定義情報["status"].startswith("対象外")


def test_word_template_and_content_detection(一時フォルダ):
    結果 = 抽出(Word文書を生成(一時フォルダ / "disguised.dat", テンプレート=True))
    assert 結果.形式 == "Word OOXML" and 結果.終了コード == 0


@pytest.mark.parametrize("VBAあり,暗号化", [(True, False), (False, False), (True, True)])
def test_legacy_word(一時フォルダ, VBAあり, 暗号化):
    ストリーム一覧 = (
        {"Macros/" + k: v for k, v in プロジェクトストリームを生成(フォームあり=False).items()}
        if VBAあり
        else {}
    )
    ヘッダー = bytearray(4096)
    struct.pack_into("<HH", ヘッダー, 0, 0xA5EC, 0xC1)
    struct.pack_into("<H", ヘッダー, 10, 0x100 if 暗号化 else 0)
    ストリーム一覧.update(WordDocument=bytes(ヘッダー), **{"0Table": bytes(4096)})
    パス = 一時フォルダ / "legacy.dot"
    パス.write_bytes(複合ファイルを生成(ストリーム一覧))
    結果 = 抽出(パス)
    assert 結果.形式 == "Word OLE"
    if 暗号化:
        assert 結果.終了コード == 1 and 結果.VBA有無 != "なし"
        assert "OFFICE_ENCRYPTED_DECRYPTION_REQUIRED" in 結果.失敗一覧
    else:
        assert 結果.終了コード == 0
        assert 結果.VBA有無 == ("あり" if VBAあり else "なし")


@pytest.mark.parametrize(
    "生成器,名前", [(Word文書を生成, "partial.docm"), (Accessデータベースを生成, "partial.accdb")]
)
def test_missing_module_preserves_settings_and_other_sources(一時フォルダ, 生成器, 名前):
    結果 = 抽出(生成器(一時フォルダ / 名前, モジュールを欠損=True))
    assert 結果.終了コード == 3 and 結果.VBA有無 == "あり"
    プロジェクト = 結果.プロジェクト一覧[0]
    assert "MODULE_SOURCE_READ_FAILED" in プロジェクト.失敗一覧
    assert プロジェクト.参照一覧 and プロジェクト.条件付きコンパイル引数["status"] == "設定あり"
    assert any(m.コード for m in プロジェクト.モジュール一覧)


def ZIPを書き換え(パス, 変更):
    with zipfile.ZipFile(パス) as アーカイブ:
        パーツ = {n: アーカイブ.read(n) for n in アーカイブ.namelist()}
    変更(パーツ)
    with zipfile.ZipFile(パス, "w") as アーカイブ:
        for 名前, 内容 in パーツ.items():
            アーカイブ.writestr(名前, 内容)


@pytest.mark.parametrize("異常", ["missing_main", "bad_xml", "missing_vba", "external_vba", "traversal"])
def test_corrupt_word_is_not_no_vba(一時フォルダ, 異常):
    パス = Word文書を生成(一時フォルダ / "broken.docm")

    def 変更(パーツ):
        if 異常 == "missing_main":
            del パーツ["custom/main.xml"]
        elif 異常 == "bad_xml":
            パーツ["custom/main.xml"] = b"<broken>"
        elif 異常 == "missing_vba":
            del パーツ["macros/project.bin"]
        else:
            パーツ["custom/_rels/main.xml.rels"] = パーツ["custom/_rels/main.xml.rels"].replace(
                b'Target="../macros/project.bin"',
                b'Target="https://example.com/vba.bin" TargetMode="External"'
                if 異常 == "external_vba"
                else b'Target="../../escape.bin"',
            )

    ZIPを書き換え(パス, 変更)
    結果 = 抽出(パス)
    assert 結果.終了コード == 1 and 結果.VBA有無 != "なし"


@pytest.mark.parametrize(
    "異常", ["missing_dir", "bad_dir", "missing_project", "missing_vba_project", "bad_source"]
)
def test_access_storage_corruption_is_partial(一時フォルダ, 異常):
    パス = Accessデータベースを生成(一時フォルダ / "broken.accdb")
    db = AccessDatabase(パス)
    t = db.table("MSysAccessStorage")
    名前 = {
        "missing_dir": "dir",
        "bad_dir": "dir",
        "missing_project": "PROJECT",
        "missing_vba_project": "_VBA_PROJECT",
        "bad_source": "source0",
    }[異常]
    位置 = next(rid for rid, row in t.rows_with_ids() if row["Name"] == 名前)
    if 異常.startswith("missing"):
        t.delete_row(位置)
    else:
        t.update_row(位置, {"Lv": b"broken"})
    db.save()
    結果 = 抽出(パス)
    assert 結果.終了コード in (1, 3) and 結果.VBA有無 == "あり"
    assert 結果.プロジェクト一覧[0].失敗一覧


def test_access_ignores_deleted_and_scratch_source(一時フォルダ):
    パス = Accessデータベースを生成(一時フォルダ / "scratch.accdb")
    前 = 抽出(パス)
    db = AccessDatabase(パス)
    t = db.table("MSysAccessStorage")
    t.insert_row({"Id": 999, "ParentId": 2, "Name": "dir", "Type": 2, "Lv": b"broken"})
    db.save()
    後 = 抽出(パス)
    assert 後.終了コード == 0
    assert 後.プロジェクト一覧[0].モジュール一覧 == 前.プロジェクト一覧[0].モジュール一覧


@pytest.mark.parametrize("種類", ["truncated", "jet3", "encrypted", "old_storage"])
def test_access_unreadable_formats_are_not_no_vba(一時フォルダ, 種類):
    パス = Accessデータベースを生成(一時フォルダ / "bad.mdb")
    if 種類 == "old_storage":
        db = AccessDatabase(パス)
        db.table("MSysObjects")
        db.rename_table("MSysAccessStorage", "MSysAccessObjects")
        db.save()
    elif 種類 == "encrypted":
        # ヘッダーの固定マスクはXORなので、保存値への差分も同じ位置へ反映できる。
        b = bytearray(パス.read_bytes())
        b[62] ^= 123
        パス.write_bytes(b)
    else:
        b = bytearray(パス.read_bytes())
        if 種類 == "jet3":
            b[4:20] = b"Standard Jet DB\0"
            b[20:24] = bytes(4)
        else:
            b = b[:-1]
        パス.write_bytes(b)
    結果 = 抽出(パス)
    assert 結果.終了コード == 1 and 結果.VBA有無 != "なし"


def test_access_exports_are_deterministic(一時フォルダ):
    パス = Accessデータベースを生成(一時フォルダ / "input.accdb")
    ハッシュ = hashlib.sha256(パス.read_bytes()).digest()
    for 名前 in ("a", "b"):
        抽出結果を保存(一時フォルダ / 名前, 抽出(パス))
    assert 比較スナップショット(一時フォルダ / "a") == 比較スナップショット(一時フォルダ / "b")
    assert hashlib.sha256(パス.read_bytes()).digest() == ハッシュ


def test_access_source_spanning_multiple_long_value_pages(一時フォルダ):
    パス = Accessデータベースを生成(一時フォルダ / "long.accdb")
    元 = 抽出(パス).プロジェクト一覧[0]
    モジュール = next(m for m in 元.モジュール一覧 if m.名前 == "集計処理")
    本文 = 'Attribute VB_Name = "集計処理"\r\nOption Explicit\r\n' + "' 長い日本語コメント\r\n" * 3000
    db = AccessDatabase(パス)
    t = db.table("MSysAccessStorage")
    位置 = next(rid for rid, row in t.rows_with_ids() if row["Name"] == モジュール.メタデータ["stream"])
    t.update_row(
        位置, {"Lv": bytes(モジュール.メタデータ["offset"]) + 圧縮データを生成(本文.encode("cp932"))}
    )
    db.save()
    結果 = 抽出(パス)
    assert 結果.終了コード == 0
    assert next(
        m for m in 結果.プロジェクト一覧[0].モジュール一覧 if m.名前 == "集計処理"
    ).コード == 本文.replace("\r\n", "\n")


def test_access_form_without_code_is_no_vba(一時フォルダ):
    パス = Accessデータベースを生成(一時フォルダ / "no_code.accdb", VBAあり=False, コードなしフォーム=True)
    結果 = 抽出(パス)
    assert 結果.終了コード == 0 and 結果.VBA有無 == "なし"


def test_word_strict_ooxml(一時フォルダ):
    パス = Word文書を生成(一時フォルダ / "strict.docx", VBAあり=False)

    def 変更(パーツ):
        パーツ["custom/main.xml"] = パーツ["custom/main.xml"].replace(
            b"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            b"http://purl.oclc.org/ooxml/wordprocessingml/main",
        )

    ZIPを書き換え(パス, 変更)
    結果 = 抽出(パス)
    assert 結果.終了コード == 0 and 結果.VBA有無 == "なし"
