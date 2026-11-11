"""MS-OFORMSの保存文字列・サイト・対応コントロールを部分解析する。完全デザインの復元は行わない。"""

from oletools.oleform import ExtendedStream, consume_GuidAndFont, consume_GuidAndPicture

from .ovba import バイナリ読取器, 解析エラー
from .runtime import 例外を記録

コントロール種類表 = {
    7: "Form",
    12: "Image",
    14: "Frame",
    15: "MorphData",
    16: "SpinButton",
    17: "CommandButton",
    18: "TabStrip",
    21: "Label",
    23: "TextBox",
    24: "ListBox",
    25: "ComboBox",
    26: "CheckBox",
    27: "OptionButton",
    28: "ToggleButton",
    47: "ScrollBar",
    57: "MultiPage",
}


# ── 1. 保存プロパティの読取 ──────────────────────────
def 位置を整列(読取器, 開始, サイズ=4):
    """指定境界まで読取位置を進め、プロパティのアラインメントを保つ。"""
    読取器.読取((-(読取器.位置 - 開始)) % サイズ)


def フィールド一覧(読取器, ビットマスク, 記述情報一覧):
    """保存マスクで存在が示されたフィールドだけを取り出す。"""
    開始 = 読取器.位置
    値一覧 = {}
    for ビット番号, 名前, width in 記述情報一覧:
        if ビットマスク & (1 << ビット番号):
            位置を整列(読取器, 開始, min(width, 4))
            値一覧[名前] = int.from_bytes(読取器.読取(width), "little")
    位置を整列(読取器, 開始)
    return 値一覧


def フォーム文字列を取得(読取器, 記述情報):
    """圧縮UnicodeまたはUTF-16の保存文字列を長さどおりに復元する。"""
    値 = 読取器.読取(記述情報 & 0x7FFFFFFF)
    # 圧縮Unicodeは各文字の下位バイトを保存する。VBAのコードページではない。
    return 値.decode("latin-1" if 記述情報 & 0x80000000 else "utf-16-le", errors="strict")


def コントロールプロパティを取得(データ, 種類):
    """対応する標準コントロールの保存済みプロパティを取得する。"""
    if 種類 not in (
        "Label",
        "CommandButton",
        "TextBox",
        "ListBox",
        "ComboBox",
        "CheckBox",
        "OptionButton",
        "ToggleButton",
    ):
        return {}
    読取器 = バイナリ読取器(データ)
    読取器.期待値を確認(読取器.読取(2), b"\x00\x02")
    ブロック = バイナリ読取器(読取器.読取(読取器.u16()))
    ビットマスク = ブロック.u32()
    if 種類 in ("Label", "CommandButton"):
        記述情報一覧 = [
            (0, "fore", 4),
            (1, "back", 4),
            (2, "various", 4),
            (3, "Caption", 4),
            (4, "picture_position", 4),
            (6, "mouse", 1),
        ]
        if 種類 == "Label":
            記述情報一覧 += [
                (7, "border_color", 4),
                (8, "border_style", 2),
                (9, "effect", 2),
                (10, "picture", 2),
                (11, "accelerator", 2),
                (12, "icon", 2),
            ]
        else:
            記述情報一覧 += [(7, "picture", 2), (8, "accelerator", 2), (10, "icon", 2)]
    elif 種類 in ("TextBox", "ListBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton"):
        ビットマスク |= ブロック.u32() << 32
        widths = {
            0: 4,
            1: 4,
            2: 4,
            3: 4,
            4: 1,
            5: 1,
            6: 1,
            7: 1,
            9: 2,
            10: 4,
            11: 2,
            12: 2,
            13: 2,
            14: 2,
            15: 2,
            16: 1,
            17: 1,
            18: 1,
            20: 1,
            21: 1,
            22: 4,
            23: 4,
            24: 4,
            25: 4,
            26: 4,
            27: 2,
            28: 2,
            29: 2,
            32: 4,
        }
        名前一覧 = {22: "Value", 23: "Caption", 32: "GroupName"}
        記述情報一覧 = [
            (ビット番号, 名前一覧.get(ビット番号, f"field_{ビット番号}"), width)
            for ビット番号, width in widths.items()
        ]
    else:
        return {}
    値一覧 = フィールド一覧(ブロック, ビットマスク, 記述情報一覧)
    if 種類 not in ("Label", "CommandButton") and ビットマスク & (1 << 8):
        ブロック.読取(8)
    結果 = {}
    for 名前 in ("Value", "Caption", "GroupName"):
        if 名前 in 値一覧:
            結果[名前] = フォーム文字列を取得(ブロック, 値一覧[名前])
    return 結果


# ── 2. サイトとフォームの部分解析 ──────────────────────────
def サイトを取得(OLEファイル, ストレージ):
    """サイトの識別子・種類・名前・保存フラグを読み取る。"""
    データ = OLEファイル.openstream(ストレージ + "/f").read()
    読取器 = バイナリ読取器(データ)
    読取器.期待値を確認(読取器.読取(2), b"\x00\x04")
    ブロック = バイナリ読取器(読取器.読取(読取器.u16()))
    ビットマスク = ブロック.u32()
    項目定義 = [
        (1, "back", 4),
        (2, "fore", 4),
        (3, "next_id", 4),
        (6, "boolean", 4),
        (7, "border", 1),
        (8, "mouse", 1),
        (9, "scroll", 1),
        (13, "group", 4),
        (15, "icon", 2),
        (16, "cycle", 1),
        (17, "effect", 1),
        (18, "border_color", 4),
        (19, "Caption", 4),
        (20, "font", 2),
        (21, "picture", 2),
        (22, "zoom", 4),
        (23, "picture_align", 1),
        (25, "picture_mode", 1),
        (26, "cookie", 4),
        (27, "buffer", 4),
    ]
    値一覧 = フィールド一覧(ブロック, ビットマスク, 項目定義)
    for ビット番号 in (10, 11, 12):
        if ビットマスク & (1 << ビット番号):
            ブロック.読取(8)
    プロパティ = {}
    if "Caption" in 値一覧:
        プロパティ["Caption"] = フォーム文字列を取得(ブロック, 値一覧["Caption"])
    # 可変長の画像・フォントは公開APIで境界だけ読み進める。復元データとは扱わない。
    import io

    残りデータ = io.BytesIO(データ[読取器.位置 :])
    拡張情報 = ExtendedStream(残りデータ, ストレージ + "/f")
    for ビット番号, 読取処理 in (
        (15, consume_GuidAndPicture),
        (20, consume_GuidAndFont),
        (21, consume_GuidAndPicture),
    ):
        if ビットマスク & (1 << ビット番号):
            読取処理(拡張情報)
    読取器.位置 += 残りデータ.tell()
    if not 値一覧.get("boolean", 0) & (1 << 15):
        for 未使用 in range(読取器.u16()):
            読取器.期待値を確認(読取器.u16(), 0)
            読取器.読取(読取器.u16())
    件数, 長さ = 読取器.u32(), 読取器.u32()
    サイトデータ = バイナリ読取器(読取器.読取(長さ))
    深さ一覧 = []
    while len(深さ一覧) < 件数:
        深さ, 混合値 = サイトデータ.読取(2)
        繰返し数 = 1
        if 混合値 & 0x80:
            繰返し数 = 混合値 & 0x7F
            サイトデータ.期待値を確認(サイトデータ.読取(1), b"\x01")
        else:
            サイトデータ.期待値を確認(混合値, 1)
        if 繰返し数 == 0 or len(深さ一覧) + 繰返し数 > 件数:
            raise 解析エラー("FORM_DEPTH_COUNT")
        深さ一覧.extend([深さ] * 繰返し数)
    位置を整列(サイトデータ, 0)
    サイト一覧 = []
    for 深さ in 深さ一覧:
        サイトデータ.期待値を確認(サイトデータ.u16(), 0)
        サイト = バイナリ読取器(サイトデータ.読取(サイトデータ.u16()))
        サイトマスク = サイト.u32()
        項目定義 = [
            (0, "Name", 4),
            (1, "Tag", 4),
            (2, "id", 4),
            (3, "help", 4),
            (4, "flags", 4),
            (5, "object_size", 4),
            (6, "tab_index", 2),
            (7, "class_index", 2),
            (9, "group", 2),
            (11, "ControlTipText", 4),
            (12, "license", 4),
            (13, "control_source", 4),
            (14, "row_source", 4),
        ]
        保存値一覧 = フィールド一覧(サイト, サイトマスク, 項目定義)
        結果 = {"depth": 深さ}
        for 名前 in ("Name", "Tag"):
            if 名前 in 保存値一覧:
                結果[名前] = フォーム文字列を取得(サイト, 保存値一覧[名前])
        if サイトマスク & (1 << 8):
            サイト.読取(8)
        if "ControlTipText" in 保存値一覧:
            結果["ControlTipText"] = フォーム文字列を取得(サイト, 保存値一覧["ControlTipText"])
        for 名前 in ("id", "flags", "object_size", "tab_index", "class_index"):
            if 名前 in 保存値一覧:
                結果[名前] = 保存値一覧[名前]
        サイト一覧.append(結果)
    return プロパティ, サイト一覧


def フォームを抽出(OLEファイル, ストレージ, 名前):
    """フォームとサイトを読み、未解析領域を状態と診断で明示する。"""
    結果 = {
        "name": 名前,
        "status": "一部取得（仕様範囲）",
        "source_storage": ストレージ,
        "properties": {},
        "controls": [],
        "complete_design": False,
        "unsupported": ["画像・フォント・完全なデザイン復元", "未対応コントロールの詳細プロパティ"],
        "diagnostics": [],
    }
    try:
        保存プロパティ, サイト一覧 = サイトを取得(OLEファイル, ストレージ)
        結果["properties"] = 保存プロパティ
        オブジェクトデータ = (
            OLEファイル.openstream(ストレージ + "/o").read() if OLEファイル.exists(ストレージ + "/o") else b""
        )
        開始位置 = 0
        for サイト in サイト一覧:
            種類 = コントロール種類表.get(サイト.get("class_index"), "不明（カスタムを含む）")
            コントロール = {
                "name": サイト.get("Name"),
                "type": 種類,
                "site": サイト,
                "parent": 名前 if サイト["depth"] == 0 else None,
                "properties": {},
            }
            結果["controls"].append(コントロール)
            # fStreamedが立つサイトだけ、共有oストリームからプロパティを取得する。
            if サイト.get("flags", 0) & 0x10 and "object_size" in サイト:
                サイズ = サイト["object_size"]
                if 開始位置 + サイズ > len(オブジェクトデータ):
                    raise 解析エラー("FORM_OBJECT_TRUNCATED")
                try:
                    コントロール["properties"] = コントロールプロパティを取得(
                        オブジェクトデータ[開始位置 : 開始位置 + サイズ], 種類
                    )
                except ValueError, IndexError:
                    結果["diagnostics"].append("FORM_CONTROL_READ_FAILED")
                開始位置 += サイズ
            else:
                コントロール["properties_status"] = "未対応（別ストレージまたは保存情報不足）"
        # 深さだけでは親を確定できない。未検証のコンテナ対応は不明のまま保持する。
        結果["parent_mapping_status"] = "一部未対応" if any(s["depth"] for s in サイト一覧) else "取得済み"
    except Exception:
        例外を記録("FORM_READ_FAILED")
        結果["status"] = "読取失敗"
        結果["diagnostics"].append("FORM_READ_FAILED")
    結果["diagnostics"] = sorted(set(結果["diagnostics"]))
    return 結果
