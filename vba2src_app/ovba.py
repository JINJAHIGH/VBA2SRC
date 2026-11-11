"""MS-OVBAのdir・PROJECTをレコード境界に沿って解析する。未知の構造を空の設定とみなさない。"""

import codecs
import re
import struct
from dataclasses import dataclass
from dataclasses import field as フィールド

from oletools.common import codepages


# ── 1. 境界を検査する読取器 ──────────────────────────
class 解析エラー(ValueError):
    """レコード構造の異常を識別コード付きで通知する。"""

    pass


class バイナリ読取器:
    """境界を検査しながらリトルエンディアンのレコードを順に読み取る。"""

    def __init__(自身, データ):
        自身.データ = データ
        自身.位置 = 0

    def 読取(自身, 長さ):
        if 長さ < 0 or 自身.位置 + 長さ > len(自身.データ):
            raise 解析エラー("TRUNCATED_RECORD")
        値 = 自身.データ[自身.位置 : 自身.位置 + 長さ]
        自身.位置 += 長さ
        return 値

    def u16(自身):
        return struct.unpack("<H", 自身.読取(2))[0]

    def u32(自身):
        return struct.unpack("<I", 自身.読取(4))[0]

    def 可変長データ(自身):
        return 自身.読取(自身.u32())

    def 期待値を確認(自身, 実際値, 期待値):
        if 実際値 != 期待値:
            raise 解析エラー("UNEXPECTED_RECORD")

    def レコード(自身, 識別子):
        自身.期待値を確認(自身.u16(), 識別子)
        return 自身.可変長データ()


# ── 2. PROJECTと保存文字列 ──────────────────────────
def 文字コードを取得(コードページ):
    # 推測や近似codecへ置き換えず、保存コードページの解釈を優先する。
    """保存コードページに正確に対応するPython codecを取得する。"""
    if コードページ in (708, 709, 710, 870, 1047, 1141, 10001, 10002, 10003, 10005, 10008, 10021):
        raise LookupError("CODEPAGE_EXACT_CODEC_UNAVAILABLE")
    文字コード = codepages.CODEPAGE_TO_CODEC.get(コードページ, f"cp{コードページ}")
    codecs.lookup(文字コード)
    return 文字コード


def 条件付きコンパイル引数(元データ):
    """PROJECTCONSTANTSの原文を保持し、重複・構文・数値範囲を検証する。"""
    項目一覧 = []
    if 元データ:
        処理済み = set()
        for エントリ in 元データ.split(":"):
            一致情報 = re.fullmatch(r"\s*([^\W\d]\w*)\s*=\s*(-?\d{1,5})\s*", エントリ, re.UNICODE)
            if not 一致情報:
                raise 解析エラー("CONSTANTS_SYNTAX")
            名前, 値 = 一致情報.groups()
            if 名前.casefold() in 処理済み or not -32768 <= int(値) <= 32767:
                raise 解析エラー("CONSTANTS_VALUE")
            処理済み.add(名前.casefold())
            項目一覧.append({"name": 名前, "value": int(値)})
    return {
        "status": "設定あり" if 元データ else "設定なし",
        "raw": 元データ,
        "items": sorted(項目一覧, key=lambda x: (x["name"].casefold(), x["name"])),
    }


def PROJECTを解析(データ, 文字コード):
    """PROJECTのプロパティとモジュール宣言を読み取る。"""
    プロパティ, 種類一覧 = {}, {}
    for 行 in データ.decode(文字コード, errors="strict").splitlines():
        if 行.startswith("["):
            break  # Workspace以降はプロジェクトプロパティではない。
        if "=" not in 行:
            continue
        キー, 値 = 行.split("=", 1)
        if len(値) >= 2 and 値.startswith('"') and 値.endswith('"'):
            値 = 値[1:-1]
        プロパティ.setdefault(キー, []).append(値)
        種類 = {"Module": "standard", "Class": "class", "Document": "document", "BaseClass": "form"}.get(キー)
        if 種類:
            名前 = 値.split("/", 1)[0] if キー == "Document" else 値
            種類一覧.setdefault(名前.casefold(), []).append(種類)
    return プロパティ, 種類一覧


def ライブラリ識別情報を解析(元データ):
    """LIBIDの保存パス・GUID・版・表示名を分離する。実PCでの解決は行わない。"""
    結果 = {
        "raw_libid": 元データ,
        "display_name": None,
        "stored_path": None,
        "guid": None,
        "major": None,
        "minor": None,
        "lcid": None,
    }
    一致情報 = re.fullmatch(
        r"\*\\([GH])(\{[^}]+\})#([0-9a-fA-F]+)\.([0-9a-fA-F]+)#([0-9a-fA-F]+)#([^#]*)#(.*)", 元データ
    )
    if 一致情報:
        種類, GUID値, メジャー版, マイナー版, 言語識別子, パス, 表示名 = 一致情報.groups()
        結果.update(
            guid=GUID値,
            major=int(メジャー版, 16),
            minor=int(マイナー版, 16),
            lcid=int(言語識別子, 16),
            stored_path=パス,
            display_name=表示名,
            path_kind=種類,
        )
    return 結果


def 参照の並び順(参照情報):
    """OSのロケールに依存せず、名前・GUID・版・保存値で整列する。"""
    名前 = 参照情報.get("name") or ""
    return (
        名前.casefold(),
        名前,
        参照情報.get("guid") or "",
        str(参照情報.get("major")),
        str(参照情報.get("minor")),
        参照情報.get("stored_path") or "",
        参照情報["original_order"],
    )


# ── 3. dirの参照・モジュールレコード ──────────────────────────
@dataclass
class ディレクトリ情報:
    """dirから回収した値と途中の解析失敗を保持する。"""

    名前: str | None = None
    説明: str | None = None
    コードページ: int | None = None
    文字コード: str | None = None
    条件付きコンパイル引数: dict = フィールド(
        default_factory=lambda: {"status": "読取失敗", "raw": None, "items": []}
    )
    参照一覧: list = フィールド(default_factory=list)
    参照取得状態: str = "読取失敗"
    モジュール一覧: list = フィールド(default_factory=list)
    モジュール数: int | None = None
    失敗一覧: list = フィールド(default_factory=list)


def 名前レコードを取得(読取器, 文字列化):
    """MBCS名とUnicode名の対応レコードを読み取る。"""
    名前 = 文字列化(読取器.可変長データ())
    if 読取器.データ[読取器.位置 : 読取器.位置 + 2] == b"\x3e\x00":
        読取器.u16()
        Unicode名前 = 読取器.可変長データ().decode("utf-16-le", errors="strict")
        if 名前 != Unicode名前:
            raise 解析エラー("REFERENCE_NAME_INCONSISTENT")
    return 名前


def 参照を取得(読取器, 文字列化, 解析情報):
    """参照の種類に従ってレコードを消費し、元の参照順も保持する。"""
    名前, 元情報 = None, None
    while True:
        識別子 = 読取器.u16()
        if 識別子 == 0x0F:
            if 名前 is not None or 元情報 is not None:
                raise 解析エラー("REFERENCE_INCOMPLETE")
            return
        if 識別子 == 0x16:
            if 名前 is not None:
                raise 解析エラー("REFERENCE_DUPLICATE_NAME")
            名前 = 名前レコードを取得(読取器, 文字列化)
            continue
        if 識別子 == 0x33:
            元情報 = 文字列化(読取器.可変長データ())
            continue
        参照情報 = {"name": 名前, "original_order": len(解析情報.参照一覧), "kind": None}
        if 識別子 == 0x0D:
            読取器.u32()  # 各サブフィールドの長さに従って読む。全体サイズだけで飛ばさない。
            参照情報.update(ライブラリ識別情報を解析(文字列化(読取器.可変長データ())), kind="registered")
            読取器.期待値を確認(読取器.u32(), 0)
            読取器.期待値を確認(読取器.u16(), 0)
        elif 識別子 == 0x0E:
            読取器.u32()
            絶対参照, 相対参照 = 文字列化(読取器.可変長データ()), 文字列化(読取器.可変長データ())
            参照情報.update(
                ライブラリ識別情報を解析(絶対参照),
                kind="project",
                raw_libid_absolute=絶対参照,
                raw_libid_relative=相対参照,
                absolute_path=絶対参照[2:] if 絶対参照.startswith("*\\C") else None,
                relative_path=相対参照[3:] if 相対参照.startswith(("*\\C", "*\\D")) else None,
                major=読取器.u32(),
                minor=読取器.u16(),
            )
            # LibidProjectのパスは3文字の種別接頭辞に続く。
            if 絶対参照.startswith("*\\C"):
                参照情報["absolute_path"] = 絶対参照[3:]
                参照情報["stored_path"] = 絶対参照[3:]
            if 相対参照.startswith(("*\\C", "*\\D")):
                参照情報["relative_path"] = 相対参照[3:]
        elif 識別子 == 0x2F:
            読取器.u32()
            変換済み識別子 = 文字列化(読取器.可変長データ())
            読取器.期待値を確認(読取器.u32(), 0)
            読取器.期待値を確認(読取器.u16(), 0)
            追加識別子 = 読取器.u16()
            拡張名 = None
            if 追加識別子 == 0x16:
                拡張名 = 名前レコードを取得(読取器, 文字列化)
                追加識別子 = 読取器.u16()
            読取器.期待値を確認(追加識別子, 0x30)
            読取器.u32()
            拡張情報 = 文字列化(読取器.可変長データ())
            読取器.期待値を確認(読取器.u32(), 0)
            読取器.期待値を確認(読取器.u16(), 0)
            型ライブラリ, クッキー = 読取器.読取(16), 読取器.u32()
            import uuid

            参照情報.update(
                ライブラリ識別情報を解析(元情報 or 拡張情報),
                kind="control",
                raw_libid_original=元情報,
                raw_libid_twiddled=変換済み識別子,
                raw_libid_extended=拡張情報,
                extended_name=拡張名,
                original_typelib=str(uuid.UUID(bytes_le=型ライブラリ)),
                cookie=クッキー,
            )
        else:
            raise 解析エラー("REFERENCE_UNSUPPORTED_RECORD")
        if 参照情報["kind"] in ("registered", "control") and 参照情報.get("guid") is None:
            解析情報.失敗一覧.append("REFERENCE_LIBID_UNPARSED")
        参照情報["resolution_status"] = "未確認（保存情報のみ）"
        解析情報.参照一覧.append(参照情報)
        名前, 元情報 = None, None


def ディレクトリを解析(データ):
    """プロジェクト・参照・モジュールの順にdirを解析し、途中の失敗を記録する。"""
    解析情報 = ディレクトリ情報()
    読取器 = バイナリ読取器(データ)
    処理段階 = "DIR_HEADER"
    try:
        読取器.期待値を確認(len(読取器.レコード(1)), 4)
        if 読取器.データ[読取器.位置 : 読取器.位置 + 2] == b"\x4a\x00":
            読取器.期待値を確認(len(読取器.レコード(0x4A)), 4)
        読取器.期待値を確認(len(読取器.レコード(2)), 4)
        読取器.期待値を確認(len(読取器.レコード(0x14)), 4)
        コードページデータ = 読取器.レコード(3)
        読取器.期待値を確認(len(コードページデータ), 2)
        解析情報.コードページ = struct.unpack("<H", コードページデータ)[0]
        解析情報.文字コード = 文字コードを取得(解析情報.コードページ)

        def 文字列化(値):
            return 値.decode(解析情報.文字コード, errors="strict")

        解析情報.名前 = 文字列化(読取器.レコード(4))
        説明 = 文字列化(読取器.レコード(5))
        Unicode説明 = 読取器.レコード(0x40).decode("utf-16-le", errors="strict")
        解析情報.説明 = 説明
        if 説明 != Unicode説明:
            解析情報.失敗一覧.append("DIR_DESCRIPTION_INCONSISTENT")
        ヘルプ情報甲 = 読取器.レコード(6)
        ヘルプ情報乙 = 読取器.レコード(0x3D)
        if ヘルプ情報甲 != ヘルプ情報乙:
            解析情報.失敗一覧.append("DIR_HELP_INCONSISTENT")
        読取器.期待値を確認(len(読取器.レコード(7)), 4)
        読取器.期待値を確認(len(読取器.レコード(8)), 4)
        読取器.期待値を確認(読取器.u16(), 9)
        読取器.期待値を確認(読取器.u32(), 4)
        読取器.読取(6)  # PROJECTVERSIONは通常の可変長レコードと構造が異なる。
        処理段階 = "PROJECTCONSTANTS"
        元データ = 文字列化(読取器.レコード(0x0C))
        Unicode元データ = 読取器.レコード(0x3C).decode("utf-16-le", errors="strict")
        解析情報.条件付きコンパイル引数["raw"] = 元データ
        解析情報.条件付きコンパイル引数["raw_unicode"] = Unicode元データ
        try:
            解析済み情報 = 条件付きコンパイル引数(元データ)
            解析情報.条件付きコンパイル引数.update(解析済み情報)
            if 元データ != Unicode元データ:
                解析情報.条件付きコンパイル引数["status"] = "不整合"
                解析情報.失敗一覧.append("PROJECTCONSTANTS_INCONSISTENT")
        except 解析エラー:
            解析情報.失敗一覧.append("PROJECTCONSTANTS_SYNTAX")
        処理段階 = "REFERENCES"
        参照を取得(読取器, 文字列化, 解析情報)
        解析情報.参照一覧.sort(key=参照の並び順)
        解析情報.参照取得状態 = "取得済み" if 解析情報.参照一覧 else "設定なし"
        処理段階 = "MODULE_DIRECTORY"
        読取器.期待値を確認(読取器.u32(), 2)  # 0x000Fの識別子は参照の読取時に消費済み。
        解析情報.モジュール数 = 読取器.u16()
        読取器.期待値を確認(len(読取器.レコード(0x13)), 2)
        for 添字 in range(解析情報.モジュール数):
            対象モジュール = {
                "original_order": 添字,
                "name": 文字列化(読取器.レコード(0x19)),
                "stream": None,
                "offset": None,
                "type_id": None,
            }
            while True:
                識別子 = 読取器.u16()
                # 種類・読取専用・非公開・終端レコードはDWORD予約値を持つ。可変長ではない。
                値 = 読取器.読取(4) if 識別子 in (0x21, 0x22, 0x25, 0x28, 0x2B) else 読取器.可変長データ()
                if 識別子 == 0x47:
                    対象モジュール["unicode_name"] = 値.decode("utf-16-le", errors="strict")
                    if 対象モジュール["unicode_name"] != 対象モジュール["name"]:
                        解析情報.失敗一覧.append("MODULE_NAME_INCONSISTENT")
                elif 識別子 == 0x1A:
                    対象モジュール["stream"] = 文字列化(値)
                    Unicodeストリーム名 = 読取器.レコード(0x32).decode("utf-16-le", errors="strict")
                    対象モジュール["unicode_stream"] = Unicodeストリーム名
                    if Unicodeストリーム名 != 対象モジュール["stream"]:
                        解析情報.失敗一覧.append("MODULE_STREAM_INCONSISTENT")
                elif 識別子 == 0x1C:
                    対象モジュール["description"] = 文字列化(値)
                    対象モジュール["description_unicode"] = 読取器.レコード(0x48).decode(
                        "utf-16-le", errors="strict"
                    )
                elif 識別子 == 0x31:
                    読取器.期待値を確認(len(値), 4)
                    対象モジュール["offset"] = int.from_bytes(値, "little")
                elif 識別子 in (0x21, 0x22):
                    読取器.期待値を確認(len(値), 4)
                    対象モジュール["type_id"] = 識別子
                elif 識別子 in (0x1E, 0x2C, 0x25, 0x28):
                    対象モジュール[f"record_{識別子:04x}"] = 値.hex()
                elif 識別子 == 0x2B:
                    読取器.期待値を確認(len(値), 4)
                    break
                else:
                    raise 解析エラー("MODULE_UNSUPPORTED_RECORD")
            解析情報.モジュール一覧.append(対象モジュール)
        読取器.期待値を確認(読取器.u16(), 0x10)
        読取器.期待値を確認(読取器.u32(), 0)
    except ValueError, LookupError, struct.error:
        解析情報.失敗一覧.append(処理段階 + "_READ_FAILED")
    return 解析情報
