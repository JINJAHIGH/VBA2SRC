"""公開JSONの項目名と原文保持、形式の安定性を検証する。"""

import json

import pytest

from tests.samples import ブックを生成
from vba2src_app.extract import 抽出
from vba2src_app.json_output import JSON文字列を作成
from vba2src_app.output import 抽出結果を保存


def test_export_has_japanese_keys_and_keeps_original_values(一時フォルダ):
    対象 = ブックを生成(一時フォルダ / "原文.xlsm")
    抽出結果を保存(一時フォルダ / "結果", 抽出(対象))

    def 項目名を検査(値):
        if isinstance(値, dict):
            for キー, 項目値 in 値.items():
                assert not キー.isascii(), キー
                assert not キー.startswith("_説明")
                項目名を検査(項目値)
        elif isinstance(値, list):
            for 項目値 in 値:
                項目名を検査(項目値)

    for パス in (一時フォルダ / "結果").rglob("*.json"):
        データ = json.loads(パス.read_text(encoding="utf-8"))
        assert データ["スキーマバージョン"] == 2
        項目名を検査(データ)
    設定フォルダ = 一時フォルダ / "結果/設定情報"
    モジュール一覧 = json.loads((設定フォルダ / "モジュール一覧.json").read_text(encoding="utf-8"))
    標準モジュール = next(値 for 値 in モジュール一覧["モジュール一覧"] if 値["種類"] == "標準モジュール")
    assert 標準モジュール["名前"] == "集計処理"
    assert 標準モジュール["Attribute情報"] == ['Attribute VB_Name = "集計処理"']
    参照設定 = json.loads((設定フォルダ / "参照設定.json").read_text(encoding="utf-8"))
    assert 参照設定["参照設定一覧"][0]["保存パス"] == r"C:\Library\Alpha.tlb"
    フォーム = json.loads((設定フォルダ / "フォーム情報/UserForm1.json").read_text(encoding="utf-8"))
    assert フォーム["保存プロパティ"]["表示文字列"] == "確認フォーム"
    assert フォーム["取得済みコントロール一覧"][0]["名前"] == "ラベル1"


def test_json_conversion_does_not_mutate_input_and_has_stable_order():
    原文 = {"name": "Caption", "raw": "B = 2 : A = 1", "items": [{"name": "B", "value": 2}]}
    内容 = JSON文字列を作成(原文)
    assert 内容 == JSON文字列を作成(dict(reversed(list(原文.items()))))
    assert 原文["name"] == "Caption" and 原文["items"][0]["name"] == "B"
    assert json.loads(内容)["元の値"] == "B = 2 : A = 1"


@pytest.mark.parametrize("データ", [{"unknown_key": 1}, {"name": 1, "Name": 2}])
def test_undefined_or_duplicate_display_keys_are_rejected(データ):
    with pytest.raises(ValueError):
        JSON文字列を作成(データ)


def test_internal_record_identifier_is_preserved():
    assert json.loads(JSON文字列を作成({"record_0025": "00000000"})) == {"内部レコード（0x0025）": "00000000"}
