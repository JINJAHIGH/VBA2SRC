import struct

from tests.samples import ディレクトリ, レコード
from vba2src_app.forms import コントロールプロパティを取得, フォーム文字列を取得
from vba2src_app.ovba import ディレクトリを解析, バイナリ読取器, 文字コードを取得


def 参照レコードを追加(参照一覧):
    元データ = ディレクトリ([], 条件付きコンパイル引数="", 参照名一覧=())
    識別マーカー = レコード(0x0F, bytes(2))
    位置 = 元データ.index(識別マーカー)
    return 元データ[:位置] + 参照一覧 + 元データ[位置:]


def 名前レコードを生成(名前):
    return レコード(0x16, 名前.encode("cp932")) + レコード(0x3E, 名前.encode("utf-16-le"))


def test_project_reference_absolute_relative_paths():
    絶対参照, 相対参照 = rb"*\CC:\Example Path\other.xls", rb"*\Cother.xls"
    ペイロード = struct.pack("<I", len(絶対参照)) + 絶対参照 + struct.pack("<I", len(相対参照)) + 相対参照
    ペイロード += struct.pack("<IH", 99, 3)
    解析情報 = ディレクトリを解析(
        参照レコードを追加(名前レコードを生成("外部Project") + レコード(0x0E, ペイロード))
    )
    assert not 解析情報.失敗一覧
    参照情報 = 解析情報.参照一覧[0]
    assert 参照情報["absolute_path"] == r"C:\Example Path\other.xls"
    assert 参照情報["relative_path"] == "other.xls"
    assert 参照情報["major"] == 99 and 参照情報["minor"] == 3


def test_control_reference_preserves_all_libids():
    元情報 = rb"*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Forms.dll#Forms"
    変換済み識別子 = rb"*\G{00000000-0000-0000-0000-000000000000}#0.0#0##"
    拡張情報 = rb"*\G{896C2D83-5466-46ED-8FAE-4C3E4F85E710}#2.0#0#C:\MSForms.exd#Microsoft Forms"
    ペイロード = (
        struct.pack("<I", len(変換済み識別子)) + 変換済み識別子 + bytes(6) + 名前レコードを生成("MSFormsEx")
    )
    ペイロード += struct.pack("<HI", 0x30, 0) + struct.pack("<I", len(拡張情報)) + 拡張情報 + bytes(6)
    ペイロード += bytes(16) + struct.pack("<I", 123)
    解析情報 = ディレクトリを解析(
        参照レコードを追加(
            名前レコードを生成("MSForms") + レコード(0x33, 元情報) + レコード(0x2F, ペイロード)
        )
    )
    assert not 解析情報.失敗一覧
    参照情報 = 解析情報.参照一覧[0]
    assert 参照情報["raw_libid_original"] == 元情報.decode()
    assert 参照情報["raw_libid_twiddled"] == 変換済み識別子.decode()
    assert 参照情報["raw_libid_extended"] == 拡張情報.decode()
    assert 参照情報["extended_name"] == "MSFormsEx"
    assert 参照情報["cookie"] == 123


def test_codepage_unknown_is_not_utf8_fallback():
    import pytest

    assert 文字コードを取得(932) == "cp932"
    with pytest.raises(LookupError):
        文字コードを取得(65535)


def test_forms_compressed_unicode_and_persisted_values():
    assert フォーム文字列を取得(バイナリ読取器(b"caf\xe9"), 0x80000004) == "café"
    値 = "入力値".encode("utf-16-le")
    表示文字列 = "確認".encode("utf-16-le")
    ビットマスク = (1 << 22) | (1 << 23)
    ブロック = struct.pack("<QII", ビットマスク, len(値), len(表示文字列)) + 値 + 表示文字列
    データ = b"\0\x02" + struct.pack("<H", len(ブロック)) + ブロック
    assert コントロールプロパティを取得(データ, "TextBox") == {"Value": "入力値", "Caption": "確認"}
    assert コントロールプロパティを取得(b"unparsed", "Frame") == {}


def test_invalid_constants_preserve_original_and_do_not_erase_modules():
    解析情報 = ディレクトリを解析(ディレクトリ([], 条件付きコンパイル引数="not a setting", 参照名一覧=()))
    assert 解析情報.条件付きコンパイル引数["raw"] == "not a setting"
    assert 解析情報.条件付きコンパイル引数["status"] == "読取失敗"
    assert "PROJECTCONSTANTS_SYNTAX" in 解析情報.失敗一覧
    assert 解析情報.参照取得状態 == "設定なし"
