"""保存結果の診断コードを、取得できた範囲と事象の説明に変換する。"""

from ..console import 表示値

事象説明 = {
    "DEFINED_NAMES_INCOMPLETE": "Excelの名前定義の一部を完全に取得できませんでした。",
    "DEFINED_NAMES_READ_FAILED": "Excelの名前定義を読み取れませんでした。",
    "DEFINED_NAMES_MULTIPLE_COLLECTIONS": "ブック内に名前定義の一覧が複数あります。",
    "DEFINED_NAMES_UNEXPECTED_ELEMENT": "名前定義の一覧に想定外の項目があります。",
    "DEFINED_NAME_CHILD_ELEMENT": "名前定義の書式が想定と異なり、参照式を完全には確認できませんでした。",
    "DEFINED_NAME_DUPLICATE": "同じ適用範囲に同名の名前定義が重複しています。",
    "DEFINED_NAME_HIDDEN_INVALID": "名前定義の非表示設定が不正です。",
    "DEFINED_NAME_NAME_MISSING": "名前が保存されていない名前定義があります。",
    "DEFINED_NAME_SCOPE_INVALID": "名前定義の適用先シートを特定できませんでした。",
    "SHEET_MAPPING_READ_FAILED": "シートの表示名とコードネームの対応を取得できませんでした。",
    "DIR_HEADER_READ_FAILED": "VBAプロジェクトの基本情報を読み取れませんでした。",
    "DIR_DESCRIPTION_INCONSISTENT": "VBAプロジェクトの説明が、保存されている文字形式によって異なります。",
    "DIR_HELP_INCONSISTENT": "VBAプロジェクトのヘルプファイルの指定が一致しません。",
    "DIR_READ_FAILED": "VBAのプロジェクト内にあるディレクトリ情報を読み取れませんでした。",
    "PROJECT_READ_FAILED": "VBAプロジェクトの設定を読み取れませんでした。",
    "PROJECT_CODEPAGE_UNAVAILABLE": "VBAプロジェクトの文字コードを特定できませんでした。",
    "PROJECTCONSTANTS_READ_FAILED": "条件付きコンパイル引数を読み取れませんでした。",
    "PROJECTCONSTANTS_INCONSISTENT": "条件付きコンパイル引数の保存内容が文字形式によって異なります。",
    "PROJECTCONSTANTS_SYNTAX": "条件付きコンパイル引数に解釈できない記述があります。",
    "REFERENCES_READ_FAILED": "参照設定を最後まで読み取れませんでした。",
    "REFERENCE_LIBID_UNPARSED": "参照設定の一部で、保存先やライブラリ識別子を解釈できませんでした。",
    "MODULE_DIRECTORY_READ_FAILED": "VBAモジュールの一覧を最後まで読み取れませんでした。",
    "MODULE_SOURCE_READ_FAILED": "VBAモジュールの一部で、ソースを取得できませんでした。",
    "MODULE_TYPE_UNDETERMINED": "VBAモジュールの一部で種類を判定できず、未分類として出力しました。",
    "MODULE_NAME_INCONSISTENT": "モジュール名の保存内容が文字形式によって異なります。",
    "MODULE_STREAM_INCONSISTENT": "モジュールの格納先の指定が文字形式によって異なります。",
    "FORM_READ_FAILED": "ユーザーフォームのデザイン情報を読み取れませんでした。",
    "FORM_CONTROL_READ_FAILED": "ユーザーフォーム内の一部のコントロール情報を取得できませんでした。",
    "VBA_PART_READ_FAILED": "Officeファイルに含まれるVBAデータの一部を読み取れませんでした。",
    "VBA_PROJECT_STREAM_MISSING": "VBAプロジェクトの保存設定が見つかりませんでした。",
    "VBA_STORAGE_MISSING": "VBAの格納場所を特定できませんでした。",
    "OLE_STRUCTURAL_ISSUES": "VBAを含む内部ストレージの構成に不整合があります。",
    "PACKAGE_VBA_UNDECLARED": "VBAデータはありますが、Officeファイル内に正式な参照の登録がありません。",
    "ACCESS_FORM_REPORT_DESIGN_UNSUPPORTED": "Accessのフォームやレポートの画面デザインは取得対象外です。",
    "ACCESS_STORAGE_LAYOUT_UNSUPPORTED": "このAccessファイルのVBA格納形式には対応していません。",
    "ACCESS_VERSION_UNSUPPORTED": "このAccessファイルの保存形式には対応していません。",
    "OFFICE_ENCRYPTED_DECRYPTION_REQUIRED": "ファイル全体が暗号化されているため、読み取れませんでした。",
    "INPUT_ACCESS_DENIED": "対象ファイルへのアクセスが拒否されました。",
    "INPUT_NOT_FOUND": "対象ファイルが見つかりませんでした。",
    "INPUT_FORMAT_UNSUPPORTED": "対象ファイルの形式には対応していません。",
    "INPUT_IO_OR_CORRUPT": "対象ファイルを読み取れませんでした。読み取り権限またはファイルの破損を確認してください。",
    "INPUT_CORRUPT_OR_PARSE_FAILED": "対象ファイルの保存内容を解析できませんでした。",
    "CMG_LENGTH": "VBA保護フラグの保存データ長が不正です。",
    "CMG_RESERVED_BITS": "VBA保護フラグに未定義の値があります。",
    "GC_CMG_INCONSISTENT": "VBAの表示状態と保護フラグが一致しません。",
    "GC_INVALID_STATE": "VBAの表示状態に解釈できない値があります。",
    "DPB_INVALID_STRUCTURE": "VBAパスワードの保存情報を解釈できませんでした。",
    "HASH_PROJECT_ID_INCONSISTENT": "VBAパスワードの保護情報がプロジェクトの識別子と一致しません。",
    "CANDIDATE_ENCODING_UNAVAILABLE": "一部のパスワード候補を、このVBAプロジェクトの文字コードで扱えませんでした。",
}


def 診断の詳細を取得(結果):
    """全体とプロジェクト別の診断を集め、影響を受けた項目を明示する。"""
    全体コード = set(結果.入力.get("失敗コード", [])) | set(結果.名前定義.get("診断コード", []))
    グループ = [("", 全体コード)]
    for プロジェクト in 結果.プロジェクト一覧:
        情報 = プロジェクト.情報
        コード一覧 = (
            set(情報.get("失敗コード", []))
            | set(情報.get("診断コード", []))
            | set(プロジェクト.保護.get("診断コード", []))
        )
        グループ.append((表示値(情報.get("名前") or プロジェクト.識別子), コード一覧))

    詳細 = []
    全コード = set()
    for 名前, コード一覧 in グループ:
        全コード.update(コード一覧)
        for コード in sorted(コード一覧):
            説明 = 事象説明.get(コード)
            if 説明 is None and コード.startswith(("CMG_", "DPB_", "GC_")):
                説明 = "VBAの保護設定の保存情報に欠落または不整合があります。"
            elif 説明 is None:
                説明 = f"{コード} の詳細は保存先のJSONを確認してください。"
            詳細.append(f"{名前}  /  {説明}" if 名前 else 説明)

    不完全な名前定義 = [
        表示値(項目.get("名前") or "名前なし")
        for 項目 in 結果.名前定義.get("名前定義一覧", [])
        if 項目.get("取得状態") == "一部取得"
    ]
    if 不完全な名前定義:
        詳細.append(f"一部の情報を確認できなかった名前定義  {', '.join(不完全な名前定義)}")

    for プロジェクト in 結果.プロジェクト一覧:
        未取得モジュール = [
            表示値(項目.get("名前") or "不明")
            for 項目 in プロジェクト.モジュール.get("モジュール一覧", [])
            if 項目.get("ソース取得状態") == "読取失敗"
        ]
        if 未取得モジュール:
            名前 = 表示値(プロジェクト.情報.get("名前") or プロジェクト.識別子)
            詳細.append(f"{名前}  /  ソース未取得のモジュール  {', '.join(未取得モジュール)}")
    return sorted(全コード), list(dict.fromkeys(詳細))


def VBA取得状況を取得(結果):
    """一覧の読取状態と保存済み件数から、取得の完了を確認できる場合だけ案内する。"""
    if 結果.入力.get("VBA有無") == "なし":
        return "このファイルにVBAは含まれていません。"
    プロジェクト一覧 = 結果.プロジェクト一覧
    if not プロジェクト一覧:
        return "VBAソースの取得状況を確認できませんでした。"
    取得数 = 0
    全件取得 = not (
        {"VBA_PART_READ_FAILED", "VBA_STORAGE_MISSING", "OLE_STRUCTURAL_ISSUES"}
        & set(結果.入力.get("失敗コード", []))
    )
    for プロジェクト in プロジェクト一覧:
        一覧 = プロジェクト.モジュール.get("モジュール一覧", [])
        件数 = sum(項目.get("ソース取得状態") == "取得済み" for 項目 in 一覧)
        取得数 += 件数
        全件取得 = (
            全件取得
            and プロジェクト.情報.get("モジュール取得状態") == "取得済み"
            and プロジェクト.情報.get("保存モジュール数") == len(一覧) == 件数
        )
    if 全件取得:
        return f"VBAソースは問題なく取得できました（{取得数}件）。"
    return f"取得できたVBAソースは{取得数}件です。すべての取得は確認できませんでした。"
