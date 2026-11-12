import struct

import pytest

from tests.samples import パスワードハッシュを生成, プロジェクト識別定数, 難読化
from vba2src_app.output import 一行化
from vba2src_app.protection import ハッシュ情報を復元, 保護情報を解析, 難読化を復元


def プロパティ(パスワード保存データ=b"\0", フラグ群=0, 表示状態=255):
    return {
        "ID": [プロジェクト識別定数],
        "CMG": [難読化(struct.pack("<I", フラグ群))],
        "DPB": [難読化(パスワード保存データ)],
        "GC": [難読化(bytes([表示状態]))],
    }


@pytest.mark.parametrize("シード", range(256))
def test_obfuscation_all_seeds(シード):
    ペイロード = bytes(range(32))
    assert 難読化を復元(難読化(ペイロード, シード)) == ペイロード


def test_published_oletools_project_example_vectors():
    # oletoolsのVBA_Project.parse_project_streamにある公開例を検証する。
    assert 難読化を復元("F1F301E705E705E705E705") == bytes(4)
    assert 難読化を復元("8F8D7FE3831F2020202020") == b"\0"
    assert 難読化を復元("2D2FDD81E51EE61EE6E1") == b"\xff"


@pytest.mark.parametrize("旧方式", [False, True])
def test_candidate_order_whitespace_case_and_skip(旧方式):
    パスワード = " 確認Pass "
    保存データ = パスワード.encode("cp932") + b"\0" if 旧方式 else パスワードハッシュを生成(パスワード)
    保存プロパティ = プロパティ(保存データ, フラグ群=4, 表示状態=0)
    空設定結果 = 保護情報を解析(保存プロパティ, "cp932", ())
    assert 空設定結果["password"] == "不明"
    assert 空設定結果["candidate_status"] == "スキップ（候補未設定）"
    一致候補 = 保護情報を解析(保存プロパティ, "cp932", ("確認Pass", " 確認pass ", パスワード, パスワード))
    assert 一致候補["candidate_status"] == "一致" and 一致候補["password"] == パスワード
    不一致結果 = 保護情報を解析(保存プロパティ, "cp932", ("確認Pass", " 確認pass "))
    assert 不一致結果["candidate_status"] == "不一致" and 不一致結果["password"] == "不明"


def test_no_password_is_not_dpb_presence_or_protection_flag():
    結果 = 保護情報を解析(プロパティ(フラグ群=4, 表示状態=0), "cp932", ("unused",))
    assert 結果["protection_state"] == "あり"
    assert 結果["password_set"] is False
    assert 結果["password"] == "なし" and 結果["candidate_status"] == "対象外"


def test_unencodable_candidate_unknown_codec_and_corruption():
    保存プロパティ = プロパティ(パスワードハッシュを生成("secret"), フラグ群=4, 表示状態=0)
    assert 保護情報を解析(保存プロパティ, "cp932", ("😀",))["candidate_status"] == "判定不能"
    assert 保護情報を解析(保存プロパティ, None, ("secret",))["candidate_status"] == "判定不能"
    assert 保護情報を解析({}, "cp932", ("secret",))["password_set"] is None
    保存プロパティ["DPB"] = ["AB"]
    assert 保護情報を解析(保存プロパティ, "cp932", ("secret",))["candidate_status"] == "判定不能"
    保存プロパティ = プロパティ(表示状態=0)
    assert "GC_CMG_INCONSISTENT" in 保護情報を解析(保存プロパティ, "cp932", ())["diagnostics"]


def test_null_hash_mask_decoding():
    # 鍵の先頭・末尾とハッシュの先頭がゼロになるビット位置を検証する。
    符号化データ = bytes([1, 34, 3, 1]) + bytes([1]) + bytes(range(1, 20))
    フラグ群 = int("0110" + "0" + "1" * 19, 2)
    キー, ダイジェスト = ハッシュ情報を復元(b"\xff" + フラグ群.to_bytes(3, "big") + 符号化データ + b"\0")
    assert キー == b"\0\x22\x03\0" and ダイジェスト == b"\0" + bytes(range(1, 20))


def test_one_line_escapes_controls_and_backslash():
    assert 一行化("a\r\nb\t\\c") == r"a\u000d\u000ab\u0009\\c"
