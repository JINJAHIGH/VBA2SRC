"""OLE・OOXMLの入力を識別し、プロジェクトとソースを読み取る。"""

import io
import logging
import posixpath
import re
import struct
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import msoffcrypto
import olefile
from oletools.olevba import decompress_stream

from .forms import フォームを抽出
from .model import プロジェクト, モジュール, 抽出結果
from .ovba import PROJECTを解析, ディレクトリを解析, 解析エラー
from .protection import 保護情報を解析
from .runtime import 例外を記録

解析ログ = logging.getLogger("vba2src")
コンテンツ種別名前空間 = "http://schemas.openxmlformats.org/package/2006/content-types"
関連付け名前空間 = "http://schemas.openxmlformats.org/package/2006/relationships"
Excelコンテンツ種別 = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml",
    "application/vnd.ms-excel.addin.macroEnabled.main+xml",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
}
Wordコンテンツ種別 = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}


# ── 1. OOXMLの関連付けとブック情報 ──────────────────────────
def XMLを解析(データ):
    # Officeの通常XMLに不要なDTDを拒否し、外部実体を展開しない。
    """DTDと外部実体を拒否してOfficeのXMLを読み取る。"""
    if b"<!DOCTYPE" in データ.upper() or b"<!ENTITY" in データ.upper():
        raise 解析エラー("XML_DTD_UNSUPPORTED")
    return ET.fromstring(データ)


def 参照パーツを解決(基準名, 対象):
    """関係先をZIP内の相対パーツとして解決し、コンテナ外への逸脱を拒否する。"""
    from urllib.parse import unquote

    パス = unquote(対象)
    if "\\" in パス or ":" in パス:
        raise 解析エラー("PACKAGE_TARGET_INVALID")
    解決済みパス = posixpath.normpath(
        パス.lstrip("/") if パス.startswith("/") else posixpath.join(posixpath.dirname(基準名), パス)
    )
    if 解決済みパス.startswith("../") or 解決済みパス in ("..", "."):
        raise 解析エラー("PACKAGE_TARGET_INVALID")
    return 解決済みパス


def 関連付けを取得(アーカイブ, パーツ):
    """Relationshipsの保存値を読み、外部リンクと内部パーツを区別する。"""
    関連付けパス = (
        (posixpath.dirname(パーツ) + "/" if posixpath.dirname(パーツ) else "")
        + "_rels/"
        + posixpath.basename(パーツ)
        + ".rels"
    )
    if not パーツ:
        関連付けパス = "_rels/.rels"
    if 関連付けパス not in アーカイブ.namelist():
        return []
    ルート = XMLを解析(アーカイブ.read(関連付けパス))
    if ルート.tag != f"{{{関連付け名前空間}}}Relationships":
        raise 解析エラー("RELATIONSHIPS_INVALID")
    return [r.attrib for r in ルート]


def シート対応を取得(アーカイブ, メイン):
    """ブックと各シートの内部情報から表示名とコードネームの対応を取得する。"""
    if not メイン.endswith(".xml"):
        return [], "未対応（バイナリブック）"
    ルート = XMLを解析(アーカイブ.read(メイン))
    関連付け表 = {r["Id"]: r for r in 関連付けを取得(アーカイブ, メイン)}
    シート一覧 = []
    for シート in ルート.findall(".//{*}sheets/{*}sheet"):
        関連付け識別子 = next((v for k, v in シート.attrib.items() if k.endswith("}id")), None)
        項目 = {"display_name": シート.get("name"), "codename": None, "part": None}
        if 関連付け識別子 not in 関連付け表 or 関連付け表[関連付け識別子].get("TargetMode") == "External":
            raise 解析エラー("SHEET_RELATIONSHIP_INVALID")
        項目["part"] = 参照パーツを解決(メイン, 関連付け表[関連付け識別子]["Target"])
        シートルート = XMLを解析(アーカイブ.read(項目["part"]))
        保存プロパティ = シートルート.find("{*}sheetPr")
        if 保存プロパティ is not None:
            項目["codename"] = 保存プロパティ.get("codeName")
        シート一覧.append(項目)
    return シート一覧, "取得済み（未保存コードネームは不明）"


def 名前定義を取得(アーカイブ, メイン):
    """セル・範囲・数式・定数の名前定義を保存式のまま取得し、評価しない。"""
    情報 = {"status": "判定不能", "items": [], "diagnostics": [], "part": メイン}
    if not メイン.endswith(".xml"):
        情報["status"] = "未対応（バイナリブック）"
        return 情報
    ルート = XMLを解析(アーカイブ.read(メイン))
    名前空間 = ルート.tag.removesuffix("workbook")
    if ルート.tag not in (
        "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbook",
        "{http://purl.oclc.org/ooxml/spreadsheetml/main}workbook",
    ):
        raise 解析エラー("DEFINED_NAMES_WORKBOOK_INVALID")
    シート一覧 = ルート.findall(f"{名前空間}sheets/{名前空間}sheet")
    定義群 = ルート.findall(f"{名前空間}definedNames")
    if len(定義群) > 1:
        情報["diagnostics"].append("DEFINED_NAMES_MULTIPLE_COLLECTIONS")
    使用済み = set()
    for 定義 in (要素 for 群 in 定義群 for 要素 in 群):
        if 定義.tag != f"{名前空間}definedName":
            情報["diagnostics"].append("DEFINED_NAMES_UNEXPECTED_ELEMENT")
            continue
        診断一覧 = []
        項目 = {
            "name": 定義.get("name"),
            "refers_to": "".join(定義.itertext()),
            "scope": "ブック",
            "scope_sheet": None,
            "scope_index": None,
            "hidden": None,
            "comment": 定義.get("comment"),
            "original_order": len(情報["items"]),
            "stored_attributes": [
                {"attribute_name": キー, "attribute_value": 値} for キー, 値 in sorted(定義.attrib.items())
            ],
        }
        if not 項目["name"] or not 項目["name"].strip():
            診断一覧.append("DEFINED_NAME_NAME_MISSING")
        if len(定義):
            診断一覧.append("DEFINED_NAME_CHILD_ELEMENT")
        if "localSheetId" in 定義.attrib:
            項目["scope"] = "不明"
            保存位置 = 定義.get("localSheetId").strip()
            # localSheetIdはsheets内の位置。sheetIdの保存値とは別の識別情報。
            if re.fullmatch(r"\+?[0-9]+", 保存位置):
                # XMLのunsignedIntは先頭ゼロを許す。巨大な値は整数化の前に制限する。
                数字列 = 保存位置.removeprefix("+").lstrip("0") or "0"
                if len(数字列) <= 10 and int(数字列) <= 0xFFFFFFFF:
                    項目["scope_index"] = int(数字列)
            if 項目["scope_index"] is not None and 項目["scope_index"] < len(シート一覧):
                項目["scope_sheet"] = シート一覧[項目["scope_index"]].get("name")
                if 項目["scope_sheet"]:
                    項目["scope"] = "シート"
            if 項目["scope"] == "不明":
                診断一覧.append("DEFINED_NAME_SCOPE_INVALID")
        if "hidden" in 定義.attrib:
            保存非表示 = 定義.get("hidden").strip()
            if 保存非表示 in ("1", "true", "0", "false"):
                項目["hidden"] = 保存非表示 in ("1", "true")
            else:
                診断一覧.append("DEFINED_NAME_HIDDEN_INVALID")
        # 同じ名前でもブックと各シートでは別定義。同じ適用範囲の重複だけを診断する。
        適用先識別子 = (
            項目["scope_index"]
            if 項目["scope"] == "シート"
            else 定義.get("localSheetId")
            if "localSheetId" in 定義.attrib
            else None
        )
        識別子 = ((項目["name"] or "").casefold(), 適用先識別子)
        if 識別子 in 使用済み:
            診断一覧.append("DEFINED_NAME_DUPLICATE")
        使用済み.add(識別子)
        項目["diagnostics"] = sorted(診断一覧)
        項目["status"] = "一部取得" if 診断一覧 else "取得済み"
        情報["diagnostics"].extend(診断一覧)
        情報["items"].append(項目)
    情報["items"].sort(
        key=lambda x: (
            (x["name"] or "").casefold(),
            x["name"] or "",
            {"ブック": 0, "シート": 1, "不明": 2}[x["scope"]],
            x["scope_index"] if x["scope_index"] is not None else -1,
            x["original_order"],
        )
    )
    情報["diagnostics"] = sorted(set(情報["diagnostics"]))
    情報["status"] = "一部取得" if 情報["diagnostics"] else "設定あり" if 情報["items"] else "設定なし"
    return 情報


# ── 2. VBAプロジェクトとソースの取得 ──────────────────────────
def 種類を判定(記述情報, 種類一覧):
    """dirの種類とPROJECT宣言を照合し、モジュール名からは推測しない。"""
    保存情報 = 種類一覧.get(記述情報["name"].casefold(), [])
    重複除外集合 = set(保存情報)
    if 記述情報["type_id"] == 0x21:
        return "standard" if not 重複除外集合 or 重複除外集合 == {"standard"} else "unknown"
    if 記述情報["type_id"] == 0x22 and len(重複除外集合) == 1 and "standard" not in 重複除外集合:
        return 保存情報[0]
    return "unknown"


def プロジェクトを抽出(OLEファイル, ルート, 抽出元識別子, 識別子, 候補一覧):
    """dir・PROJECT・ソース・フォームを独立に取得し、回収結果と失敗を共存させる。"""
    対象プロジェクト = プロジェクト(識別子, 抽出元識別子)
    解析ログ.debug("project=%s source=%s", 識別子, 抽出元識別子)
    try:
        元データ = decompress_stream(bytearray(OLEファイル.openstream(ルート + "VBA/dir").read()))
        解析情報 = ディレクトリを解析(元データ)
        対象プロジェクト.名前, 対象プロジェクト.説明 = 解析情報.名前, 解析情報.説明
        対象プロジェクト.コードページ, 対象プロジェクト.文字コード = (
            解析情報.コードページ,
            解析情報.文字コード,
        )
        対象プロジェクト.条件付きコンパイル引数, 対象プロジェクト.参照一覧 = (
            解析情報.条件付きコンパイル引数,
            解析情報.参照一覧,
        )
        対象プロジェクト.参照取得状態 = 解析情報.参照取得状態
        対象プロジェクト.保存モジュール数 = 解析情報.モジュール数
        対象プロジェクト.モジュール取得状態 = (
            "取得済み"
            if 解析情報.モジュール数 is not None
            and len(解析情報.モジュール一覧) == 解析情報.モジュール数
            and "MODULE_DIRECTORY_READ_FAILED" not in 解析情報.失敗一覧
            else "読取失敗"
        )
        対象プロジェクト.失敗一覧.extend(解析情報.失敗一覧)
    except Exception:
        例外を記録("DIR_READ_FAILED")
        対象プロジェクト.失敗一覧.append("DIR_READ_FAILED")
        対象プロジェクト.保護情報 = 保護情報を解析({}, None, 候補一覧)
        return 対象プロジェクト
    プロパティ, 種類一覧 = {}, {}
    if 解析情報.文字コード:
        try:
            プロパティ, 種類一覧 = PROJECTを解析(
                OLEファイル.openstream(ルート + "PROJECT").read(), 解析情報.文字コード
            )
        except Exception:
            例外を記録("PROJECT_READ_FAILED")
            対象プロジェクト.失敗一覧.append("PROJECT_READ_FAILED")
    else:
        対象プロジェクト.失敗一覧.append("PROJECT_CODEPAGE_UNAVAILABLE")
    対象プロジェクト.保護情報 = 保護情報を解析(プロパティ, 解析情報.文字コード, 候補一覧)
    # 候補の符号化不能は照合の制限として扱い、保存情報の破損と混同しない。
    対象プロジェクト.失敗一覧.extend(
        コード
        for コード in 対象プロジェクト.保護情報["diagnostics"]
        if コード != "CANDIDATE_ENCODING_UNAVAILABLE"
    )
    for 記述情報 in 解析情報.モジュール一覧:
        ストリーム = ルート + "VBA/" + (記述情報["stream"] or 記述情報["name"])
        対象モジュール = モジュール(
            記述情報["name"], 種類を判定(記述情報, 種類一覧), ストリーム, メタデータ=記述情報
        )
        対象プロジェクト.モジュール一覧.append(対象モジュール)
        解析ログ.debug("module source=%s type=%s", ストリーム, 対象モジュール.種類)
        try:
            if 記述情報["offset"] is None or 記述情報["stream"] is None:
                raise 解析エラー("MODULE_DESCRIPTOR_INCOMPLETE")
            # パス要素を分離してOLEへ渡し、モジュール名をパスとして解釈させない。
            コードデータ = OLEファイル.openstream(
                ルート.rstrip("/").split("/") + ["VBA", 記述情報["stream"]]
                if ルート
                else ["VBA", 記述情報["stream"]]
            ).read()
            開始位置 = 記述情報["offset"]
            if 開始位置 >= len(コードデータ):
                raise 解析エラー("MODULE_OFFSET_INVALID")
            元コード = decompress_stream(bytearray(コードデータ[開始位置:]))
            コード = 元コード.decode(解析情報.文字コード, errors="strict")
            対象モジュール.コード = コード.replace("\r\n", "\n").replace("\r", "\n")
            if not 対象モジュール.コード.endswith("\n"):
                対象モジュール.コード += "\n"
            対象モジュール.属性情報 = [
                行
                for 行 in 対象モジュール.コード.splitlines()
                if re.match(r"^\s*Attribute\s", 行, re.IGNORECASE)
            ]
            if 対象モジュール.種類 == "unknown":
                対象プロジェクト.失敗一覧.append("MODULE_TYPE_UNDETERMINED")
        except Exception:
            例外を記録("MODULE_SOURCE_READ_FAILED")
            対象プロジェクト.失敗一覧.append("MODULE_SOURCE_READ_FAILED")
        if 対象モジュール.種類 == "form":
            フォーム = フォームを抽出(OLEファイル, ルート + 対象モジュール.名前, 対象モジュール.名前)
            対象プロジェクト.フォーム一覧.append(フォーム)
            対象プロジェクト.失敗一覧.extend(フォーム["diagnostics"])
    対象プロジェクト.モジュール一覧.sort(
        key=lambda m: (m.名前.casefold(), m.名前, m.抽出元ストリーム, m.メタデータ["original_order"])
    )
    対象プロジェクト.フォーム一覧.sort(key=lambda f: (f["name"].casefold(), f["name"]))
    return 対象プロジェクト


def OLEプロジェクトを取得(データ, 抽出元識別子, 候補一覧, 結果):
    """OLE内のVBAストレージを列挙し、不完全なプロジェクトも検出する。"""
    件数 = len(結果.プロジェクト一覧)
    with olefile.OleFileIO(io.BytesIO(データ), raise_defects=olefile.DEFECT_INCORRECT) as OLEファイル:
        # olevbaの通常列挙で省略される不完全なVBAストレージも検出する。
        ルート一覧 = sorted(
            {
                "/".join(パス[:-1])
                for パス in OLEファイル.listdir(streams=False, storages=True)
                if パス[-1].casefold() == "vba"
            },
            key=lambda s: (s.casefold(), s),
        )
        if not ルート一覧 and any(
            パス[-1].casefold() in ("project", "_vba_project") for パス in OLEファイル.listdir()
        ):
            結果.失敗一覧.append("VBA_STORAGE_MISSING")
        for ルート in ルート一覧:
            接頭辞 = ルート + "/" if ルート else ""
            識別子 = f"project_{len(結果.プロジェクト一覧) + 1:03d}"
            対象プロジェクト = プロジェクトを抽出(
                OLEファイル, 接頭辞, 抽出元識別子 + "!" + (ルート or "/"), 識別子, 候補一覧
            )
            結果.プロジェクト一覧.append(対象プロジェクト)
            if not OLEファイル.exists(接頭辞 + "VBA/_VBA_PROJECT"):
                対象プロジェクト.失敗一覧.append("VBA_PROJECT_STREAM_MISSING")
        if OLEファイル.parsing_issues:
            結果.失敗一覧.append("OLE_STRUCTURAL_ISSUES")
    return len(結果.プロジェクト一覧) - 件数


# ── 3. 入力コンテナの識別と抽出 ──────────────────────────
def BIFF暗号化を判定(データ):
    """BIFFレコードのFILEPASSを調べ、文書全体の暗号化を識別する。"""
    読取器 = io.BytesIO(データ)
    開始レコード検出, 終端レコード検出 = False, False
    while 読取器.tell() < len(データ):
        ヘッダー = 読取器.read(4)
        if len(ヘッダー) != 4:
            raise 解析エラー("BIFF_TRUNCATED")
        識別子, 長さ = struct.unpack("<HH", ヘッダー)
        ペイロード = 読取器.read(長さ)
        if len(ペイロード) != 長さ:
            raise 解析エラー("BIFF_TRUNCATED")
        if not 開始レコード検出:
            if 識別子 not in (0x09, 0x209, 0x409, 0x809) or 長さ < 4:
                raise 解析エラー("BIFF_BOF_INVALID")
            開始レコード検出 = True
        if 識別子 == 0x2F:
            return True
        if 識別子 == 0x0A:
            終端レコード検出 = True
            break
    if not 終端レコード検出:
        raise 解析エラー("BIFF_EOF_MISSING")
    return False


def OLE入力を解析(データ, 候補一覧, 結果):
    """OLEの暗号化と内容を確認して、独立したVBAプロジェクトを取得する。"""
    結果.形式 = "OLE"
    with olefile.OleFileIO(io.BytesIO(データ), raise_defects=olefile.DEFECT_INCORRECT) as OLEファイル:
        if OLEファイル.exists("EncryptedPackage") or OLEファイル.exists("EncryptionInfo"):
            結果.致命的失敗 = True
            結果.失敗一覧.append("OFFICE_ENCRYPTED_DECRYPTION_REQUIRED")
            return 結果
        ブック = next((名前 for 名前 in ("Workbook", "Book") if OLEファイル.exists(名前)), None)
        if ブック:
            # 文書の暗号化メタデータだけを確認し、VBA用候補を復号へ流用しない。
            Officeファイル = msoffcrypto.OfficeFile(io.BytesIO(データ))
            if Officeファイル.is_encrypted() or BIFF暗号化を判定(OLEファイル.openstream(ブック).read()):
                結果.致命的失敗 = True
                結果.失敗一覧.append("OFFICE_ENCRYPTED_DECRYPTION_REQUIRED")
                return 結果
            結果.形式 = "Excel BIFF/OLE"
        elif OLEファイル.exists("WordDocument"):
            結果.形式 = "Word OLE"
            ヘッダー = OLEファイル.openstream("WordDocument").read(32)
            if len(ヘッダー) != 32 or ヘッダー[:2] != b"\xec\xa5":
                raise 解析エラー("WORD_FIB_INVALID")
            if int.from_bytes(ヘッダー[2:4], "little") not in (0xC1, 0xD9, 0x101, 0x10C, 0x112):
                raise 解析エラー("WORD_VERSION_UNSUPPORTED")
            if msoffcrypto.OfficeFile(io.BytesIO(データ)).is_encrypted():
                結果.致命的失敗 = True
                結果.失敗一覧.append("OFFICE_ENCRYPTED_DECRYPTION_REQUIRED")
                return 結果
            if not OLEファイル.exists("1Table" if ヘッダー[11] & 2 else "0Table"):
                raise 解析エラー("WORD_TABLE_STREAM_MISSING")
        elif OLEファイル.exists("PowerPoint Document"):
            結果.致命的失敗 = True
            結果.失敗一覧.append("INPUT_FORMAT_UNSUPPORTED")
            return 結果
    検出済み = OLEプロジェクトを取得(データ, "OLE", 候補一覧, 結果)
    if not ブック and 結果.形式 != "Word OLE" and not 検出済み:
        結果.致命的失敗 = True
        結果.失敗一覧.append("INPUT_FORMAT_UNSUPPORTED")
        return 結果
    結果.名前定義情報["status"] = (
        "未対応（BIFF/OLE）"
        if ブック
        else "対象外（Word）"
        if 結果.形式 == "Word OLE"
        else "対象外（VBAプロジェクト単体）"
    )
    if 結果.形式 == "Word OLE":
        for 対象プロジェクト in 結果.プロジェクト一覧:
            対象プロジェクト.シート取得状態 = "対象外（Word）"


def OOXML入力を解析(データ, 候補一覧, 結果):
    """Content TypesとRelationshipsを照合してExcel・WordとVBAパーツを取得する。"""
    結果.形式 = "OOXML"
    with zipfile.ZipFile(io.BytesIO(データ)) as アーカイブ:
        名前一覧 = アーカイブ.namelist()
        if len(名前一覧) != len(set(名前一覧)):
            raise 解析エラー("PACKAGE_DUPLICATE_PARTS")
        if アーカイブ.testzip():
            raise 解析エラー("PACKAGE_CRC_FAILED")
        種別一覧 = XMLを解析(アーカイブ.read("[Content_Types].xml"))
        if 種別一覧.tag != f"{{{コンテンツ種別名前空間}}}Types":
            raise 解析エラー("PACKAGE_CONTENT_TYPES_INVALID")
        上書き種別表 = {
            p.attrib["PartName"].lstrip("/"): p.attrib["ContentType"]
            for p in 種別一覧
            if p.tag.endswith("}Override")
        }
        既定種別表 = {
            p.attrib["Extension"]: p.attrib["ContentType"] for p in 種別一覧 if p.tag.endswith("}Default")
        }
        Office関連付け一覧 = [
            r
            for r in 関連付けを取得(アーカイブ, "")
            if r.get("Type", "").endswith("/officeDocument") and r.get("TargetMode") != "External"
        ]
        if len(Office関連付け一覧) != 1:
            raise 解析エラー("PACKAGE_MAIN_RELATIONSHIP_INVALID")
        メイン = 参照パーツを解決("", Office関連付け一覧[0]["Target"])
        メイン種別 = 上書き種別表.get(メイン)
        if メイン種別 not in Excelコンテンツ種別 | Wordコンテンツ種別:
            結果.致命的失敗 = True
            結果.失敗一覧.append("INPUT_FORMAT_UNSUPPORTED")
            return 結果
        アーカイブ.read(メイン)  # 主パーツの欠損は破損として扱い、VBAなしへ分類しない。
        Excel入力 = メイン種別 in Excelコンテンツ種別
        結果.形式 = "Excel OOXML" if Excel入力 else "Word OOXML"
        if Excel入力:
            try:
                結果.名前定義情報 = 名前定義を取得(アーカイブ, メイン)
                if 結果.名前定義情報["diagnostics"]:
                    結果.失敗一覧.append("DEFINED_NAMES_INCOMPLETE")
            except Exception:
                結果.名前定義情報.update(status="読取失敗", part=メイン)
                結果.名前定義情報["diagnostics"].append("DEFINED_NAMES_READ_FAILED")
                結果.失敗一覧.append("DEFINED_NAMES_READ_FAILED")
                例外を記録("DEFINED_NAMES_READ_FAILED")
        else:
            結果.名前定義情報["status"] = "対象外（Word）"
            if XMLを解析(アーカイブ.read(メイン)).tag not in (
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document",
                "{http://purl.oclc.org/ooxml/wordprocessingml/main}document",
            ):
                raise 解析エラー("WORD_DOCUMENT_INVALID")
        VBAパーツ一覧 = {
            名前
            for 名前 in 名前一覧
            if 上書き種別表.get(名前, 既定種別表.get(名前.rsplit(".", 1)[-1]))
            == "application/vnd.ms-office.vbaProject"
        }
        VBAパーツ一覧.update(
            名前
            for 名前, コンテンツ種別 in 上書き種別表.items()
            if コンテンツ種別 == "application/vnd.ms-office.vbaProject"
        )
        主文書のVBAパーツ = set()
        for 名前 in sorted(名前一覧):
            if 名前.endswith(".rels"):
                関連付けルート = XMLを解析(アーカイブ.read(名前))
                基準名 = "" if 名前 == "_rels/.rels" else 名前.replace("/_rels/", "/")[:-5]
                for 関連付け in 関連付けルート:
                    if 関連付け.get("Type", "").endswith("/vbaProject"):
                        if 関連付け.get("TargetMode") == "External":
                            raise 解析エラー("VBA_EXTERNAL_TARGET")
                        パーツ = 参照パーツを解決(基準名, 関連付け.attrib["Target"])
                        VBAパーツ一覧.add(パーツ)
                        if 基準名 == メイン:
                            主文書のVBAパーツ.add(パーツ)
        # 未宣言のVBAパーツを検出し、構造不整合をVBAなしと誤判定しない。
        for 名前 in 名前一覧:
            if 名前.rsplit("/", 1)[-1].casefold() == "vbaproject.bin" and 名前 not in VBAパーツ一覧:
                VBAパーツ一覧.add(名前)
                結果.失敗一覧.append("PACKAGE_VBA_UNDECLARED")
        シート一覧, シート取得状態 = [], "対象外（Word）"
        if Excel入力:
            try:
                シート一覧, シート取得状態 = シート対応を取得(アーカイブ, メイン)
            except Exception:
                シート一覧, シート取得状態 = [], "読取失敗"
                結果.失敗一覧.append("SHEET_MAPPING_READ_FAILED")
                例外を記録("SHEET_MAPPING_READ_FAILED")
        for パーツ in sorted(VBAパーツ一覧, key=lambda s: (s.casefold(), s)):
            処理前データ = len(結果.プロジェクト一覧)
            try:
                if not OLEプロジェクトを取得(アーカイブ.read(パーツ), パーツ, 候補一覧, 結果):
                    raise 解析エラー("VBA_PROJECT_MISSING")
            except Exception:
                例外を記録("VBA_PART_READ_FAILED")
                結果.失敗一覧.append("VBA_PART_READ_FAILED")
            if not Excel入力 or (パーツ in 主文書のVBAパーツ and len(主文書のVBAパーツ) == 1):
                for 対象プロジェクト in 結果.プロジェクト一覧[処理前データ:]:
                    対象プロジェクト.シート一覧 = シート一覧
                    対象プロジェクト.シート取得状態 = シート取得状態


def 抽出(対象, 候補一覧=()):
    """入力を一度だけ読み取り、内容に基づいてOLEまたはOOXMLへ振り分ける。"""
    対象 = Path(対象)
    結果 = 抽出結果(対象.name)
    try:
        データ = 対象.read_bytes()  # 元ファイルは一度だけ読み、以降はメモリ上のデータを解析する。
        if データ.startswith(olefile.MAGIC):
            OLE入力を解析(データ, 候補一覧, 結果)
        elif データ.startswith(b"PK"):
            OOXML入力を解析(データ, 候補一覧, 結果)
        elif データ[4:20] in (b"Standard ACE DB\0", b"Standard Jet DB\0"):
            from .access import Access入力を解析

            Access入力を解析(データ, 候補一覧, 結果)
        else:
            結果.致命的失敗 = True
            結果.失敗一覧.append("INPUT_FORMAT_UNSUPPORTED")
        if 結果.致命的失敗:
            return 結果
        # 名前定義の異常を、独立して確認したVBA有無の判定へ混在させない。
        VBA判定に関わる失敗 = [c for c in 結果.失敗一覧 if not c.startswith("DEFINED_NAMES_")]
        結果.VBA有無 = "あり" if 結果.プロジェクト一覧 else ("判定不能" if VBA判定に関わる失敗 else "なし")
        if VBA判定に関わる失敗 and not 結果.プロジェクト一覧:
            結果.致命的失敗 = not bool(結果.名前定義情報["items"])
        解析ログ.debug(
            "format=%s projects=%d modules=%d",
            結果.形式,
            len(結果.プロジェクト一覧),
            sum(len(p.モジュール一覧) for p in 結果.プロジェクト一覧),
        )
    except PermissionError:
        結果.致命的失敗 = True
        結果.失敗一覧.append("INPUT_ACCESS_DENIED")
        例外を記録("INPUT_ACCESS_DENIED")
    except FileNotFoundError:
        結果.致命的失敗 = True
        結果.失敗一覧.append("INPUT_NOT_FOUND")
        例外を記録("INPUT_NOT_FOUND")
    except OSError:
        結果.致命的失敗 = True
        結果.失敗一覧.append("INPUT_IO_OR_CORRUPT")
        例外を記録("INPUT_IO_OR_CORRUPT")
    except Exception:
        結果.致命的失敗 = not bool(結果.プロジェクト一覧 or 結果.名前定義情報["items"])
        結果.失敗一覧.append("INPUT_CORRUPT_OR_PARSE_FAILED")
        例外を記録("INPUT_CORRUPT_OR_PARSE_FAILED")
    if 結果.プロジェクト一覧:
        結果.VBA有無 = "あり"
    return 結果
