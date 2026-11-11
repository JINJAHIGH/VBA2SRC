"""Jet 4/ACEの現行ストレージを読み、共通VBA解析器へ渡す。Officeは起動しない。"""

import io
from dataclasses import replace

from pyopenvba.access._rows import split_row
from pyopenvba.access.database import AccessDatabase

from .ovba import 解析エラー

Access署名一覧 = (b"Standard ACE DB\0", b"Standard Jet DB\0")


class Accessストレージ:
    """MSysAccessStorageをOLEと同じ読取インターフェイスで参照する。

    Lvは各ストリームの読取時に復元する。壊れた一つの長値があっても、
    dirや他のモジュールを回収できる。削除済み・scratchの行は使用しない。
    pyOpenVBAの内部RawRow APIは、このアダプター内だけで使用する。
    """

    def __init__(自身, テーブル):
        自身.テーブル = テーブル
        自身.行一覧 = {}
        自身.子一覧 = {}
        自身.生行一覧 = {}
        長値列 = テーブル.definition.column("Lv").number
        for ページ, スロット, データ in テーブル.raw_rows():
            生行 = split_row(テーブル.definition, データ)
            行 = テーブル.decode(replace(生行, values={**生行.values, 長値列: None}))
            識別子 = 行["Id"]
            if 識別子 in 自身.行一覧:
                raise 解析エラー("ACCESS_STORAGE_DUPLICATE_ID")
            自身.行一覧[識別子] = 行
            自身.生行一覧[識別子] = 生行
            自身.子一覧.setdefault(行["ParentId"], []).append(識別子)
        ルート一覧 = [
            行["Id"]
            for 行 in 自身.行一覧.values()
            if 行["Name"] == "MSysAccessStorage_ROOT" and 行["Type"] == 1
        ]
        if len(ルート一覧) != 1:
            raise 解析エラー("ACCESS_STORAGE_ROOT_INVALID")
        自身.ルート = ルート一覧[0]
        自身.プロジェクト = 自身.子を取得(自身.子を取得(自身.ルート, "VBA"), "VBAProject")

    def 子を取得(自身, 親, 名前):
        if 親 is None:
            return None
        候補 = [
            i
            for i in 自身.子一覧.get(親, [])
            if i != 親 and 自身.行一覧[i]["Name"].casefold() == 名前.casefold()
        ]
        if len(候補) > 1:
            raise 解析エラー("ACCESS_STORAGE_DUPLICATE_NAME")
        return 候補[0] if 候補 else None

    def 項目を取得(自身, パス):
        パーツ一覧 = パス.split("/") if isinstance(パス, str) else パス
        現在 = 自身.プロジェクト
        for パーツ in パーツ一覧:
            if 現在 is None or 自身.行一覧[現在]["Type"] != 1:
                return None
            現在 = 自身.子を取得(現在, パーツ)
        return 現在

    def exists(自身, パス):
        return 自身.項目を取得(パス) is not None

    def openstream(自身, パス):
        識別子 = 自身.項目を取得(パス)
        if 識別子 is None or 自身.行一覧[識別子]["Type"] != 2:
            raise 解析エラー("ACCESS_STREAM_MISSING")
        データ = 自身.テーブル.decode(自身.生行一覧[識別子])["Lv"]
        if not isinstance(データ, bytes):
            raise 解析エラー("ACCESS_STREAM_VALUE_MISSING")
        return io.BytesIO(データ)


def Access入力を解析(データ, 候補一覧, 結果):
    """ページをメモリ上で読み、現在のVBAProjectだけを解析する。"""
    from .extract import プロジェクトを抽出

    結果.形式 = "Access ACE" if データ[4:20] == Access署名一覧[0] else "Access Jet 4"
    結果.名前定義情報["status"] = "対象外（Access）"
    if len(データ) < 24 or int.from_bytes(データ[20:24], "little") not in (1, 2, 3, 4, 5, 6):
        結果.致命的失敗 = True
        結果.失敗一覧.append("ACCESS_VERSION_UNSUPPORTED")
        return
    データベース = AccessDatabase(データ)  # bytesを渡し、保存先パスを与えない。
    if データベース.header.encoding_key or データベース.header.password:
        結果.致命的失敗 = True
        結果.失敗一覧.append("OFFICE_ENCRYPTED_DECRYPTION_REQUIRED")
        return
    テーブル名一覧 = データベース.table_names(include_system=True)
    if "MSysAccessStorage" not in テーブル名一覧:
        # Access 2000の別ストレージをVBAなしと誤判定しない。
        結果.致命的失敗 = True
        結果.失敗一覧.append("ACCESS_STORAGE_LAYOUT_UNSUPPORTED")
        return
    ストレージ = Accessストレージ(データベース.table("MSysAccessStorage"))
    if ストレージ.プロジェクト is None:
        if any(行.type == -32761 for 行 in データベース.catalog()):
            raise 解析エラー("ACCESS_VBA_STORAGE_MISSING")
        return
    対象プロジェクト = プロジェクトを抽出(
        ストレージ, "", "MSysAccessStorage!VBA/VBAProject", "project_001", 候補一覧
    )
    対象プロジェクト.シート取得状態 = "対象外（Access）"
    対象プロジェクト.診断一覧.append("ACCESS_FORM_REPORT_DESIGN_UNSUPPORTED")
    if not ストレージ.exists("VBA/_VBA_PROJECT"):
        対象プロジェクト.失敗一覧.append("VBA_PROJECT_STREAM_MISSING")
    結果.プロジェクト一覧.append(対象プロジェクト)
