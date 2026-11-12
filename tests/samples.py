"""Original synthetic fixtures, MIT licensed with this repository.

No Office application, private documents or downloaded macro samples are used.
The minimal CFB writer deliberately uses regular (>=4096 byte) streams only.
"""

import hashlib
import math
import struct
import zipfile

終端 = 0xFFFFFFFE
未使用領域 = 0xFFFFFFFF
プロジェクト識別定数 = "{00000000-0000-0000-0000-000000000000}"


def レコード(識別子, 値):
    return struct.pack("<HI", 識別子, len(値)) + 値


def 難読化(データ, シード=0x05, プロジェクト識別子=プロジェクト識別定数):
    キー = sum(プロジェクト識別子.encode("ascii")) & 255
    結果 = bytearray((シード, シード ^ 2, シード ^ キー))
    平文データ, 直前値, 前々値 = キー, 結果[2], 結果[1]
    無視データ = bytes(range((シード & 6) // 2))
    for 値 in 無視データ + struct.pack("<I", len(データ)) + データ:
        暗号化データ = 値 ^ ((前々値 + 平文データ) & 255)
        結果.append(暗号化データ)
        平文データ, 前々値, 直前値 = 値, 直前値, 暗号化データ
    return 結果.hex().upper()


def パスワードハッシュを生成(パスワード, 文字コード="cp932", キー=b"\0\x22\x03\0"):
    元データ = キー + hashlib.sha1(パスワード.encode(文字コード) + キー).digest()
    フラグ群 = sum((値 != 0) << (23 - 添字) for 添字, 値 in enumerate(元データ))
    return b"\xff" + フラグ群.to_bytes(3, "big") + bytes(値 or 1 for 値 in 元データ) + b"\0"


def 圧縮データを生成(データ):
    結果 = bytearray(b"\x01")
    # リテラルだけのチャンクもMS-OVBAの有効な圧縮コンテナとなる。
    for 開始 in range(0, len(データ), 3500):
        チャンク = データ[開始 : 開始 + 3500]
        ペイロード = b"".join(b"\0" + チャンク[i : i + 8] for i in range(0, len(チャンク), 8))
        結果 += struct.pack("<H", 0xB000 | (len(ペイロード) - 1)) + ペイロード
    return bytes(結果)


def ディレクトリ(
    モジュール一覧,
    条件付きコンパイル引数="z = -1 : Alpha = 3",
    参照名一覧=("zeta", "Alpha", "alpha"),
    コードページ=932,
):
    文字コード = f"cp{コードページ}"
    データ = レコード(1, struct.pack("<I", 3)) + レコード(2, struct.pack("<I", 0x409))
    データ += レコード(0x14, struct.pack("<I", 0x409)) + レコード(3, struct.pack("<H", コードページ))
    データ += レコード(4, "テストProject".encode(文字コード))
    説明 = "比較テスト"
    データ += レコード(5, 説明.encode(文字コード)) + レコード(0x40, 説明.encode("utf-16-le"))
    データ += レコード(6, b"") + レコード(0x3D, b"") + レコード(7, bytes(4)) + レコード(8, bytes(4))
    データ += struct.pack("<HIIH", 9, 4, 1, 0)
    データ += レコード(0x0C, 条件付きコンパイル引数.encode(文字コード)) + レコード(
        0x3C, 条件付きコンパイル引数.encode("utf-16-le")
    )
    for 名前 in 参照名一覧:
        ライブラリ識別子 = (
            rf"*\G{{00020430-0000-0000-C000-000000000046}}#2.0#0#C:\Library\{名前}.tlb#{名前} library".encode(
                文字コード
            )
        )
        データ += レコード(0x16, 名前.encode(文字コード)) + レコード(0x3E, 名前.encode("utf-16-le"))
        ペイロード = struct.pack("<I", len(ライブラリ識別子)) + ライブラリ識別子 + bytes(6)
        データ += レコード(0x0D, ペイロード)
    データ += レコード(0x0F, struct.pack("<H", len(モジュール一覧))) + レコード(0x13, bytes(2))
    for 対象モジュール in モジュール一覧:
        名前, ストリーム = 対象モジュール["name"], 対象モジュール["stream"]
        データ += レコード(0x19, 名前.encode(文字コード)) + レコード(0x47, 名前.encode("utf-16-le"))
        データ += レコード(0x1A, ストリーム.encode(文字コード)) + レコード(
            0x32, ストリーム.encode("utf-16-le")
        )
        データ += レコード(0x1C, b"") + レコード(0x48, b"")
        データ += レコード(0x31, struct.pack("<I", 対象モジュール["offset"]))
        データ += レコード(0x1E, bytes(4)) + レコード(0x2C, bytes(2))
        データ += struct.pack("<HIHI", 0x21 if 対象モジュール["kind"] == "Module" else 0x22, 0, 0x2B, 0)
    return データ + レコード(0x10, b"")


def フォームストリームを生成():
    # ルートCaptionと、Unicodeの名前・Captionを持つLabelを一つ保存する。
    表示文字列 = "確認フォーム".encode("utf-16-le")
    ブロック = struct.pack("<III", (1 << 6) | (1 << 19), 1 << 15, len(表示文字列)) + 表示文字列
    ラベル文字列 = "表示文字列".encode("utf-16-le")
    ラベルブロック = struct.pack("<II", 1 << 3, len(ラベル文字列)) + ラベル文字列
    表示ラベル = b"\0\x02" + struct.pack("<H", len(ラベルブロック)) + ラベルブロック
    名前 = "ラベル1".encode("utf-16-le")
    ビットマスク = (1 << 0) | (1 << 2) | (1 << 4) | (1 << 5) | (1 << 7)
    サイト = struct.pack("<IIIIIH", ビットマスク, len(名前), 1, 0x10, len(表示ラベル), 21) + bytes(2) + 名前
    サイト = b"\0\0" + struct.pack("<H", len(サイト)) + サイト
    サイトデータ = b"\0\x01\0\0" + サイト
    f = (
        b"\0\x04"
        + struct.pack("<H", len(ブロック))
        + ブロック
        + struct.pack("<II", 1, len(サイトデータ))
        + サイトデータ
    )
    return f, 表示ラベル


def プロジェクトストリームを生成(
    パスワード=None,
    旧方式=False,
    条件付きコンパイル引数="z = -1 : Alpha = 3",
    フォームあり=True,
    モジュールを欠損=False,
    モジュール仕様=None,
):
    仕様一覧 = (
        list(モジュール仕様)
        if モジュール仕様 is not None
        else [
            ("集計処理", "Module"),
            ("Customer", "Class"),
            ("ThisWorkbook", "Document"),
            ("Sheet1", "Document"),
        ]
    )
    if フォームあり:
        仕様一覧.append(("UserForm1", "BaseClass"))
    モジュール一覧, ストリーム一覧, 対象プロジェクト = [], {}, [f'ID="{プロジェクト識別定数}"']
    for 添字, (名前, 種類) in enumerate(仕様一覧):
        コード = f'Attribute VB_Name = "{名前}"\r\nOption Explicit\r\n\r\n\t\' 日本語コメント\r\n#If VBA7 Then\r\n#Else\r\n#End If\r\n\r\n'
        ペイロード = 圧縮データを生成(コード.encode("cp932"))
        開始位置 = 4096 - len(ペイロード)
        ストリーム = f"source{添字}"
        if not (モジュールを欠損 and 添字 == 0):
            ストリーム一覧[f"VBA/{ストリーム}"] = bytes(開始位置) + ペイロード
        モジュール一覧.append({"name": 名前, "kind": 種類, "stream": ストリーム, "offset": 開始位置})
        対象プロジェクト.append(f"{種類}={名前}" + ("/&H00000000" if 種類 == "Document" else ""))
    パスワードデータ = (
        b"\0"
        if パスワード is None
        else (パスワード.encode("cp932") + b"\0" if 旧方式 else パスワードハッシュを生成(パスワード))
    )
    対象プロジェクト += [
        'Name="テストProject"',
        f'CMG="{難読化(struct.pack("<I", 0 if パスワード is None else 4))}"',
        f'DPB="{難読化(パスワードデータ)}"',
        f'GC="{難読化(bytes([255 if パスワード is None else 0]))}"',
        "",
        "[Workspace]",
    ]
    ストリーム一覧["PROJECT"] = "\r\n".join(対象プロジェクト).encode("cp932").ljust(4096, b" ")
    元ディレクトリ = ディレクトリ(モジュール一覧, 条件付きコンパイル引数=条件付きコンパイル引数)
    ストリーム一覧["VBA/dir"] = b"\x01\xff\x3f" + 元ディレクトリ.ljust(4096, b"\0")
    ストリーム一覧["VBA/_VBA_PROJECT"] = b"\xcc\x61" + bytes(4094)
    if フォームあり:
        f, o = フォームストリームを生成()
        ストリーム一覧["UserForm1/f"] = f.ljust(4096, b"\0")
        ストリーム一覧["UserForm1/o"] = o.ljust(4096, b"\0")
    return ストリーム一覧


def 複合ファイルを生成(ストリーム一覧):
    ノード一覧 = [
        {
            "name": "Root Entry",
            "type": 5,
            "children": {},
            "left": 未使用領域,
            "right": 未使用領域,
            "child": 未使用領域,
            "start": 終端,
            "size": 0,
        }
    ]
    for パス, データ in sorted(ストリーム一覧.items()):
        親 = 0
        パーツ一覧 = パス.split("/")
        for 添字, 名前 in enumerate(パーツ一覧):
            子要素 = ノード一覧[親]["children"]
            if 名前 not in 子要素:
                子要素[名前] = len(ノード一覧)
                ストリーム判定 = 添字 == len(パーツ一覧) - 1
                ノード一覧.append(
                    {
                        "name": 名前,
                        "type": 2 if ストリーム判定 else 1,
                        "children": {},
                        "left": 未使用領域,
                        "right": 未使用領域,
                        "child": 未使用領域,
                        "start": 終端,
                        "size": 0,
                        "data": データ if ストリーム判定 else b"",
                    }
                )
            親 = 子要素[名前]
    for ノード in ノード一覧:
        識別子一覧 = sorted(
            ノード["children"].values(),
            key=lambda i: (len(ノード一覧[i]["name"]), ノード一覧[i]["name"].upper()),
        )

        def 木構造を作成(項目一覧):
            if not 項目一覧:
                return 未使用領域
            中央位置 = len(項目一覧) // 2
            現在値 = 項目一覧[中央位置]
            ノード一覧[現在値]["left"] = 木構造を作成(項目一覧[:中央位置])
            ノード一覧[現在値]["right"] = 木構造を作成(項目一覧[中央位置 + 1 :])
            return 現在値

        ノード["child"] = 木構造を作成(識別子一覧)
    セクター一覧, 割当表 = [], []

    def 割当(データ):
        開始 = len(セクター一覧)
        for i in range(math.ceil(len(データ) / 512)):
            セクター一覧.append(データ[i * 512 : (i + 1) * 512].ljust(512, b"\0"))
            割当表.append(終端 if i == math.ceil(len(データ) / 512) - 1 else len(セクター一覧))
        return 開始

    for ノード in ノード一覧[1:]:
        if ノード["type"] == 2:
            assert len(ノード["data"]) >= 4096
            ノード["size"] = len(ノード["data"])
            ノード["start"] = 割当(ノード["data"])
    エントリ一覧 = bytearray()
    for ノード in ノード一覧:
        名前 = (ノード["name"] + "\0").encode("utf-16-le")
        assert len(名前) <= 64
        エントリ = bytearray(128)
        エントリ[: len(名前)] = 名前
        struct.pack_into(
            "<HBBIII",
            エントリ,
            64,
            len(名前),
            ノード["type"],
            1,
            ノード["left"],
            ノード["right"],
            ノード["child"],
        )
        struct.pack_into("<IQ", エントリ, 116, ノード["start"], ノード["size"])
        エントリ一覧 += エントリ
    ディレクトリセクター = 割当(エントリ一覧)
    割当表数 = 1
    while 割当表数 * 128 < len(セクター一覧) + 割当表数:
        割当表数 += 1
    割当表開始位置 = len(セクター一覧)
    割当表.extend([0xFFFFFFFD] * 割当表数)
    割当表データ = struct.pack(
        "<" + "I" * (割当表数 * 128), *(割当表 + [未使用領域] * (割当表数 * 128 - len(割当表)))
    )
    セクター一覧.extend(割当表データ[i : i + 512] for i in range(0, len(割当表データ), 512))
    ヘッダー = bytearray(512)
    ヘッダー[:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", ヘッダー, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", ヘッダー, 40, 0, 割当表数, ディレクトリセクター, 0, 4096, 終端, 0, 終端, 0)
    struct.pack_into(
        "<109I",
        ヘッダー,
        76,
        *list(range(割当表開始位置, 割当表開始位置 + 割当表数)),
        *([未使用領域] * (109 - 割当表数)),
    )
    return bytes(ヘッダー) + b"".join(セクター一覧)


def ブックを生成(パス, プロジェクト一覧=1, 名前定義XML="", **追加引数):
    ペイロード = 複合ファイルを生成(プロジェクトストリームを生成(**追加引数)) if プロジェクト一覧 else None
    with zipfile.ZipFile(パス, "w", zipfile.ZIP_DEFLATED) as アーカイブ:
        種別一覧 = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
        for i in range(プロジェクト一覧):
            種別一覧 += (
                f'<Override PartName="/xl/vba{i}.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
            )
            アーカイブ.writestr(f"xl/vba{i}.bin", ペイロード)
        アーカイブ.writestr("[Content_Types].xml", 種別一覧 + "</Types>")
        アーカイブ.writestr(
            "_rels/.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        アーカイブ.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="売上表" sheetId="1" r:id="s1"/></sheets>'
            + 名前定義XML
            + "</workbook>",
        )
        マクロ関連付け = (
            '<Relationship Id="v1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vba0.bin"/>'
            if プロジェクト一覧
            else ""
        )
        アーカイブ.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            + マクロ関連付け
            + "</Relationships>",
        )
        アーカイブ.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr codeName="Sheet1"/></worksheet>',
        )
    return パス


def Word文書を生成(パス, VBAあり=True, テンプレート=False, **追加引数):
    """主文書を独自位置へ置き、拡張子や固定パスに依存しないWordを生成する。"""
    with zipfile.ZipFile(パス, "w", zipfile.ZIP_DEFLATED) as アーカイブ:
        主種別 = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            if パス.suffix.lower() == ".docx"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
            if パス.suffix.lower() == ".dotx"
            else "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
            if テンプレート or パス.suffix.lower() == ".dotm"
            else "application/vnd.ms-word.document.macroEnabled.main+xml"
        )
        アーカイブ.writestr(
            "[Content_Types].xml",
            f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'<Override PartName="/custom/main.xml" ContentType="{主種別}"/>'
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
        )
        アーカイブ.writestr(
            "_rels/.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/main.xml"/></Relationships>',
        )
        アーカイブ.writestr(
            "custom/main.xml",
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
        )
        if VBAあり:
            ストリーム一覧 = プロジェクトストリームを生成(
                モジュール仕様=[("集計処理", "Module"), ("Customer", "Class"), ("ThisDocument", "Document")],
                **追加引数,
            )
            アーカイブ.writestr("macros/project.bin", 複合ファイルを生成(ストリーム一覧))
            アーカイブ.writestr(
                "custom/_rels/main.xml.rels",
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="vba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="../macros/project.bin"/></Relationships>',
            )
    return パス


def Accessデータベースを生成(パス, VBAあり=True, コードなしフォーム=False, **追加引数):
    """依存ライブラリの空DBから、独自のVBAを持つページ・長値のfixtureを生成する。

    ユーザーの実文書は使わない。VBAProject以外のテンプレート情報は解析対象外。
    """
    from pyopenvba.access.database import AccessDatabase

    データベース = AccessDatabase.create_new(パス)
    if コードなしフォーム:
        データベース.create_form("Orders")
    テーブル = データベース.table("MSysAccessStorage")
    行一覧 = list(テーブル.rows_with_ids())
    プロジェクト親 = next(行["Id"] for _, 行 in 行一覧 if 行["Name"] == "VBAProject")
    VBA親 = next(行["Id"] for _, 行 in 行一覧 if 行["Name"] == "VBA" and 行["ParentId"] == プロジェクト親)
    for 位置, 行 in 行一覧:
        if 行["ParentId"] in (プロジェクト親, VBA親) and (行["Type"] == 2 or not VBAあり):
            テーブル.delete_row(位置)
        elif not VBAあり and 行["Id"] == プロジェクト親:
            テーブル.delete_row(位置)
    if VBAあり:
        ストリーム一覧 = プロジェクトストリームを生成(
            フォームあり=False,
            モジュール仕様=[
                ("集計処理", "Module"),
                ("Customer", "Class"),
                ("Form_Orders", "Document"),
                ("Report_Summary", "Document"),
            ],
            **追加引数,
        )
        次の識別子 = max(行["Id"] for _, 行 in 行一覧) + 1
        for 名前, 値 in ストリーム一覧.items():
            テーブル.insert_row(
                {
                    "Id": 次の識別子,
                    "ParentId": VBA親 if 名前.startswith("VBA/") else プロジェクト親,
                    "Name": 名前.removeprefix("VBA/"),
                    "Type": 2,
                    "Lv": 値,
                }
            )
            次の識別子 += 1
    else:
        # 空DBと同じ、モジュールを持たないカタログにする。
        カタログ = データベース.table("MSysObjects")
        for 位置, 行 in list(カタログ.rows_with_ids()):
            if 行["Type"] == -32761:
                カタログ.delete_row(位置)
    データベース.save()
    return パス
