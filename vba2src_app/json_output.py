"""内部データを、項目名が日本語の比較用JSONへ変換する。"""

import json
import re

# ── 1. 出力スキーマ ───────────────────────────────────────────
# 内部の解析キーと表示名を分離する。新項目の追加時はここで意味を確定する。
スキーマバージョン = 2
項目表示名 = {
    "schema_version": "スキーマバージョン",
    "project_id": "プロジェクト識別子",
    "original_filename": "元ファイル名",
    "sanitized_filename": "保存用ファイル名",
    "format": "ファイル形式",
    "vba": "VBA有無",
    "projects": "プロジェクト一覧",
    "id": "識別子",
    "origin": "抽出元識別子",
    "saved_prefix": "保存先接頭辞",
    "name": "名前",
    "Name": "名前",
    "description": "説明",
    "description_unicode": "Unicode説明",
    "codepage": "コードページ",
    "codec": "文字コード",
    "modules": "モジュール一覧",
    "modules_status": "モジュール取得状態",
    "modules_expected": "保存モジュール数",
    "source_origin": "抽出元識別子",
    "forms": "フォーム一覧",
    "sheets": "シート一覧",
    "sheets_status": "シート対応取得状態",
    "failures": "失敗コード",
    "diagnostics": "診断コード",
    "kind": "種類",
    "source_stream": "抽出元ストリーム",
    "attributes": "Attribute情報",
    "saved_path": "保存ファイル",
    "metadata": "内部保存情報",
    "codename": "コードネーム",
    "source_status": "ソース取得状態",
    "original_order": "元の順序（0始まり）",
    "stream": "ストリーム名",
    "offset": "ソース開始位置（バイト）",
    "type_id": "モジュール種別ID",
    "unicode_name": "Unicode名",
    "unicode_stream": "Unicodeストリーム名",
    "status": "取得状態",
    "raw": "元の値",
    "raw_unicode": "Unicode原文",
    "items": "設定一覧",
    "value": "設定値",
    "references": "参照設定一覧",
    "display_name": "表示名",
    "stored_path": "保存パス",
    "guid": "ライブラリGUID",
    "major": "メジャーバージョン",
    "minor": "マイナーバージョン",
    "lcid": "言語識別子（LCID）",
    "path_kind": "保存パス形式",
    "raw_libid": "元のLIBID",
    "raw_libid_absolute": "絶対参照のLIBID",
    "raw_libid_relative": "相対参照のLIBID",
    "absolute_path": "外部プロジェクト絶対パス",
    "relative_path": "外部プロジェクト相対パス",
    "raw_libid_original": "変換前のLIBID",
    "raw_libid_twiddled": "変換済みLIBID",
    "raw_libid_extended": "拡張LIBID",
    "extended_name": "拡張参照名",
    "original_typelib": "元の型ライブラリGUID",
    "cookie": "保存クッキー",
    "resolution_status": "現在PCでの解決確認状態",
    "protection_state": "VBA保護状態",
    "flags": "保護・動作フラグ",
    "user": "利用者による保護",
    "host": "ホストによる保護",
    "vbe": "VBEによる保護",
    "visibility": "プロジェクト表示状態",
    "password_set": "パスワード設定有無",
    "password_method": "パスワード保存方式",
    "password": "一致したパスワード",
    "candidate_status": "候補照合結果",
    "fields": "保護フィールド解析情報",
    "CMG": "保護フラグ（CMG）",
    "DPB": "パスワード情報（DPB）",
    "GC": "表示状態（GC）",
    "present": "保存項目の有無",
    "count": "保存項目数",
    "decoded_length": "復元後のバイト数",
    "source_storage": "抽出元ストレージ",
    "properties": "保存プロパティ",
    "controls": "取得済みコントロール一覧",
    "complete_design": "完全なデザイン復元",
    "unsupported": "未対応項目",
    "parent_mapping_status": "親子関係取得状態",
    "type": "コントロール種類",
    "site": "サイト保存情報",
    "parent": "親コントロール名",
    "properties_status": "プロパティ取得状態",
    "depth": "階層の深さ",
    "Tag": "タグ",
    "ControlTipText": "ヒント文字列",
    "object_size": "オブジェクトデータのバイト数",
    "tab_index": "タブ順序",
    "class_index": "コントロール種類ID",
    "Caption": "表示文字列",
    "Value": "保存値",
    "GroupName": "グループ名",
    "part": "パッケージ内パーツ",
    "refers_to": "参照先の保存式",
    "defined_names": "名前定義一覧",
    "scope": "適用範囲",
    "scope_sheet": "適用先シート名",
    "scope_index": "適用先シートの位置（0始まり）",
    "hidden": "非表示フラグ（保存値）",
    "comment": "コメント",
    "stored_attributes": "保存属性一覧",
    "attribute_name": "保存属性名",
    "attribute_value": "保存属性値",
}
種類表示名 = {
    "standard": "標準モジュール",
    "class": "クラスモジュール",
    "document": "ドキュメントモジュール",
    "form": "ユーザーフォーム",
    "unknown": "未分類",
    "registered": "登録型ライブラリ参照",
    "project": "外部VBAプロジェクト参照",
    "control": "コントロール型ライブラリ参照",
}


# ── 2. 項目名の変換 ─────────────────────────────────────────
def 項目名を取得(キー):
    """既知の項目を日本語化し、仕様のない項目名を出力へ流さない。"""
    if キー in 項目表示名:
        return 項目表示名[キー]
    if isinstance(キー, str) and re.fullmatch(r"record_[0-9a-f]{4}", キー):
        return f"内部レコード（0x{キー[7:].upper()}）"
    raise ValueError(f"JSON項目の表示名が未定義：{キー}")


def 日本語項目へ変換(値):
    """入れ子と配列を再帰変換する。元データを書き換えず、値の原文を保持する。"""
    if isinstance(値, dict):
        変換済み = {}
        for キー, 項目値 in 値.items():
            表示名 = 項目名を取得(キー)
            if 表示名 in 変換済み:
                raise ValueError(f"JSON項目名が重複：{表示名}")
            # 種類コードは本ツールの判定値。原文名やLIBIDの文字列は置換しない。
            if キー == "kind" and isinstance(項目値, str):
                項目値 = 種類表示名.get(項目値, 項目値)
            変換済み[表示名] = 日本語項目へ変換(項目値)
        return 変換済み
    if isinstance(値, list):
        return [日本語項目へ変換(項目値) for 項目値 in 値]
    return 値


# ── 3. 比較用JSONの整形 ──────────────────────────────────────
def JSON文字列を作成(データ):
    """項目名の辞書順・2スペース・末尾LFを固定して比較差分を安定させる。"""
    return (
        json.dumps(日本語項目へ変換(データ), ensure_ascii=False, indent=2, sort_keys=True, allow_nan=False)
        + "\n"
    )
