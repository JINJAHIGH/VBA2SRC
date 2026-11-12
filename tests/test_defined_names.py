"""名前定義の適用範囲、保存式、部分回収、比較の安定性を検証する。"""

import io
import json
import zipfile

import pytest

from tests.samples import ブックを生成, プロジェクトストリームを生成, 複合ファイルを生成
from tests.test_extraction import 比較スナップショット
from vba2src_app.cli import メイン
from vba2src_app.extract import 抽出
from vba2src_app.output import 抽出結果を保存

名前定義XML = """<definedNames>
<definedName name="z定数" comment="税率">0.1</definedName>
<definedName name="売上合計" localSheetId="0" hidden="1" comment="日本語の説明">'売上表'!$B$2:$B$20</definedName>
<definedName name="Alpha">SUM('売上表'!$B$2:$B$20)</definedName>
<definedName name="Alpha" localSheetId="0" hidden="false">#REF!</definedName>
<definedName name="_xlnm.Print_Area" localSheetId="0">'売上表'!$A$1:$F$30</definedName>
<definedName name="外部参照">'[別 ブック.xlsx]表1'!$A$1</definedName>
</definedNames>"""


def test_names_are_workbook_metadata_without_vba_and_keep_saved_values(一時フォルダ):
    対象 = ブックを生成(一時フォルダ / "日本語 名前 (1).xlsx", プロジェクト一覧=0, 名前定義XML=名前定義XML)
    元内容 = 対象.read_bytes()
    結果 = 抽出(対象)
    assert 結果.終了コード == 0 and 結果.VBA有無 == "なし"
    assert 結果.名前定義情報["status"] == "設定あり"
    項目一覧 = 結果.名前定義情報["items"]
    assert [項目["name"] for 項目 in 項目一覧] == [
        "_xlnm.Print_Area",
        "Alpha",
        "Alpha",
        "z定数",
        "売上合計",
        "外部参照",
    ]
    assert [項目["original_order"] for 項目 in 項目一覧] == [4, 2, 3, 0, 1, 5]
    assert 項目一覧[1]["scope"] == "ブック" and 項目一覧[1]["scope_index"] is None
    assert 項目一覧[1]["hidden"] is None
    assert 項目一覧[2]["refers_to"] == "#REF!" and 項目一覧[2]["hidden"] is False
    assert 項目一覧[4]["scope_sheet"] == "売上表" and 項目一覧[4]["hidden"] is True
    assert 項目一覧[4]["comment"] == "日本語の説明"
    assert 項目一覧[5]["refers_to"] == "'[別 ブック.xlsx]表1'!$A$1"
    抽出結果を保存(一時フォルダ / "結果", 結果)
    保存内容 = json.loads((一時フォルダ / "結果/設定情報/名前定義一覧.json").read_text(encoding="utf-8"))
    assert 保存内容["プロジェクト識別子"] is None
    売上 = next(項目 for 項目 in 保存内容["名前定義一覧"] if 項目["名前"] == "売上合計")
    assert 売上["参照先の保存式"] == "'売上表'!$B$2:$B$20"
    assert 売上["適用範囲"] == "シート" and 売上["適用先シート名"] == "売上表"
    レポート = (一時フォルダ / "結果/抽出結果一覧.txt").read_text(encoding="utf-8")
    assert "名前定義数：6" in レポート and "設定情報/名前定義一覧.json" in レポート
    assert 対象.read_bytes() == 元内容


def test_names_keep_multiple_project_separation_and_debug_determinism(一時フォルダ):
    対象 = ブックを生成(一時フォルダ / "対象.xlsm", プロジェクト一覧=2, 名前定義XML=名前定義XML)
    for オプション in ([], ["--debug-log"]):
        assert メイン([str(対象), "-o", str(一時フォルダ / "保存先"), "--no-ui", *オプション]) == 0
    出力一覧 = sorted((一時フォルダ / "保存先").iterdir())
    assert 比較スナップショット(出力一覧[0]) == 比較スナップショット(出力一覧[1])
    assert len(list(出力一覧[0].rglob("名前定義一覧.json"))) == 1


@pytest.mark.parametrize(
    "追加定義,診断",
    [
        ('<definedName name="bad" localSheetId="-1">A1</definedName>', "DEFINED_NAME_SCOPE_INVALID"),
        ('<definedName name="bad" localSheetId="4">A1</definedName>', "DEFINED_NAME_SCOPE_INVALID"),
        ('<definedName name="bad" localSheetId="x">A1</definedName>', "DEFINED_NAME_SCOPE_INVALID"),
        ('<definedName name="bad" hidden="maybe">A1</definedName>', "DEFINED_NAME_HIDDEN_INVALID"),
        ("<definedName>A1</definedName>", "DEFINED_NAME_NAME_MISSING"),
        ("<unexpected/>", "DEFINED_NAMES_UNEXPECTED_ELEMENT"),
        ('<definedName name="ok">A1</definedName>', "DEFINED_NAME_DUPLICATE"),
    ],
)
def test_malformed_names_preserve_good_rows_and_vba_absence(一時フォルダ, 追加定義, 診断):
    対象 = ブックを生成(
        一時フォルダ / "不整合.xlsx",
        プロジェクト一覧=0,
        名前定義XML='<definedNames><definedName name="ok">42</definedName>' + 追加定義 + "</definedNames>",
    )
    元内容 = 対象.read_bytes()
    結果 = 抽出(対象)
    assert 結果.終了コード == 3 and 結果.VBA有無 == "なし"
    assert 結果.名前定義情報["status"] == "一部取得"
    assert 診断 in 結果.名前定義情報["diagnostics"]
    assert any(項目["refers_to"] == "42" for 項目 in 結果.名前定義情報["items"])
    assert 対象.read_bytes() == 元内容


def test_sheet_scope_uses_position_not_sheet_id_and_strict_namespace(一時フォルダ):
    対象 = ブックを生成(
        一時フォルダ / "シート.xlsx",
        プロジェクト一覧=0,
        名前定義XML='<definedNames><definedName name="定義" localSheetId="0">A1</definedName><definedName name="定義" localSheetId="+0000000000000">B1</definedName></definedNames>',
    )
    with zipfile.ZipFile(io.BytesIO(対象.read_bytes())) as アーカイブ:
        パーツ = {名前: アーカイブ.read(名前) for 名前 in アーカイブ.namelist()}
    パーツ["xl/workbook.xml"] = (
        パーツ["xl/workbook.xml"]
        .replace(b'sheetId="1"', b'sheetId="52"')
        .replace(
            b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            b"http://purl.oclc.org/ooxml/spreadsheetml/main",
        )
    )
    with zipfile.ZipFile(対象, "w") as アーカイブ:
        for 名前, 内容 in パーツ.items():
            アーカイブ.writestr(名前, 内容)
    結果 = 抽出(対象)
    assert 結果.名前定義情報["items"][0]["scope_sheet"] == "売上表"
    assert 結果.名前定義情報["items"][0]["scope_index"] == 0
    assert "DEFINED_NAME_DUPLICATE" in 結果.名前定義情報["diagnostics"]


def test_absent_names_and_read_failure_are_distinct(一時フォルダ):
    正常 = ブックを生成(一時フォルダ / "正常.xlsx", プロジェクト一覧=0)
    assert 抽出(正常).名前定義情報["status"] == "設定なし"
    破損 = ブックを生成(一時フォルダ / "破損.xlsx", プロジェクト一覧=0, 名前定義XML="<definedNames>")
    結果 = 抽出(破損)
    assert 結果.終了コード == 1
    assert 結果.名前定義情報["status"] == "読取失敗"


def test_names_survive_unreadable_vba_part(一時フォルダ):
    対象 = ブックを生成(一時フォルダ / "VBA破損.xlsm", 名前定義XML=名前定義XML)
    with zipfile.ZipFile(io.BytesIO(対象.read_bytes())) as アーカイブ:
        パーツ = {名前: アーカイブ.read(名前) for 名前 in アーカイブ.namelist()}
    パーツ["xl/vba0.bin"] = b"broken VBA part"
    with zipfile.ZipFile(対象, "w") as アーカイブ:
        for 名前, 内容 in パーツ.items():
            アーカイブ.writestr(名前, 内容)
    結果 = 抽出(対象)
    assert 結果.終了コード == 3 and 結果.VBA有無 == "判定不能"
    assert 結果.名前定義情報["status"] == "設定あり"
    assert len(結果.名前定義情報["items"]) == 6
    assert "VBA_PART_READ_FAILED" in 結果.失敗一覧


def test_binary_formats_do_not_claim_names_are_absent(一時フォルダ):
    単体 = 一時フォルダ / "vbaProject.bin"
    単体.write_bytes(複合ファイルを生成(プロジェクトストリームを生成()))
    assert 抽出(単体).名前定義情報["status"] == "対象外（VBAプロジェクト単体）"
    対象 = ブックを生成(一時フォルダ / "binary.xlsb", プロジェクト一覧=0)
    with zipfile.ZipFile(io.BytesIO(対象.read_bytes())) as アーカイブ:
        パーツ = {名前: アーカイブ.read(名前) for 名前 in アーカイブ.namelist()}
    for 名前 in ("[Content_Types].xml", "_rels/.rels"):
        パーツ[名前] = (
            パーツ[名前]
            .replace(b"workbook.xml", b"workbook.bin")
            .replace(
                b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                b"application/vnd.ms-excel.sheet.binary.macroEnabled.main",
            )
        )
    del パーツ["xl/workbook.xml"]
    パーツ["xl/workbook.bin"] = b"binary"
    with zipfile.ZipFile(対象, "w") as アーカイブ:
        for 名前, 内容 in パーツ.items():
            アーカイブ.writestr(名前, 内容)
    結果 = 抽出(対象)
    assert 結果.終了コード == 0 and 結果.VBA有無 == "なし"
    assert 結果.名前定義情報["status"] == "未対応（バイナリブック）"
