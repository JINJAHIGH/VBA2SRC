# VBA2SRC の仕様と VBA 連携

対象バージョンは VBA2SRC 1.6.2、JSON スキーマは `2` です。VBA2SRC は Office ファイルを読み取り専用で解析し、保存済みの VBA ソースと設定情報をファイルに書き出します。開いている文書の編集内容を対象にする場合は、Office で保存してからファイルのパスを渡します。

## コマンドライン

Windows x64 の配布 EXE は Python と Office を導入していない環境でも起動できます。ソースから起動する場合は Python 3.13 以上で `python -m pip install -r requirements.txt` を実行します。1 回の起動で入力ファイルを 1 件処理します。複数ファイルは個別に起動してください。

```powershell
.\dist\VBA2SRC.exe "C:\Work\売上集計 (月次).xlsm" -o "C:\Work\抽出結果" --no-ui --group-by-type
python vba2src.py "C:\Work\売上集計 (月次).xlsm" -o "C:\Work\抽出結果" --no-ui --group-by-type
```

| 引数 | 意味 |
| --- | --- |
| `target` | 保存済みの入力ファイル。空白を含むパスは引用符で囲む |
| `-o` / `--output` | 結果フォルダの親。未作成なら作成し、省略時は入力ファイルの親を使う |
| `--no-ui` | 新しいコンソールや入力待ちを無効化する。VBA・バッチなどからの起動向け |
| `--group-by-type` | 通常の出力に加え、結果の直下に `VBAまとめ/` を作る |
| `--debug-log` | 詳細ログと例外のスタック情報を追加する |
| `--help` / `--version` | ヘルプまたはバージョンを表示する。入力ファイルは省略できる |

相対パスは呼び出し元のカレントフォルダを基準に解決します。自動処理では絶対パスを渡すと基準が明確になります。引数なしではGUIが開きます。GUIの「設定」で既定の出力先とフォントを変更できます。抽出画面の出力先を空欄にすると、対象ファイルと同じフォルダへ保存します。VBA保護の候補設定は、同じ利用者のGUIとCLIに適用されます。

## 対応する入力と取得範囲

| 入力 | 読み取れる情報 |
| --- | --- |
| Excel | VBA プロジェクトを含む OOXML、BIFF/OLE、バイナリブックの保存済みソース。XML 形式のブックではセル・範囲などの名前定義も取得する |
| Word | OOXML の `.docm` / `.dotm`、VBA なしの `.docx` / `.dotx`、対応する Word 97 以降の OLE。`ThisDocument` はドキュメントモジュールとして扱う |
| Access | Jet 4 / ACE の対応する `MSysAccessStorage` 構造を持つ `.mdb` / `.accdb`。フォームやレポートの VBA はドキュメントモジュールとして扱う |

共通の取得対象は、モジュール本文・Attribute・種別、プロジェクト情報、条件付きコンパイル引数、保存参照設定、VBA プロジェクト保護情報です。ユーザーフォームでは、コードに加えて一部の保存プロパティとコントロール情報を取得します。入力形式はファイルのシグネチャと内部構造から判定します。

対象外には、Office 文書全体の復号、VBA の実行・書き戻し、フォームデザインの完全復元、Access のテーブルデータや Access マクロ、PowerPoint、コンパイル済みコードからのソース復元が含まれます。BIFF やバイナリブックの名前定義、Excel のセル内容やテーブル名も取得範囲外です。[Word・Access の詳細](Word_Access対応.md)

## 結果の保存場所

出力先の親に `<保存用ファイル名>_YYYYMMDD_HHMMSS` を作ります。元の拡張子は名前の一部です。同名フォルダがある場合は `_001`、`_002` のように連番を付けます。自動処理で結果を特定するには、呼び出しごとに空の親フォルダを用意して `-o` へ渡し、その直下の結果フォルダを確認します。

```text
売上集計.xlsm_20261112_143000/
├── 入力情報.json
├── 名前定義一覧.json
├── プロジェクト情報.json
├── モジュール一覧.json
├── 条件付きコンパイル引数.json
├── 参照設定.json
├── 保護情報.json
├── VBAソース/
│   ├── 標準モジュール/
│   ├── クラスモジュール/
│   ├── ドキュメントモジュール/
│   └── ユーザーフォーム/
│       ├── フォーム名.txt
│       └── フォーム名.json
├── VBAまとめ/                --group-by-type を指定した場合
└── ログ/実行ログ.txt          --debug-log を指定した場合
```

上図はプロジェクトが 1 件の場合です。複数の場合、`入力情報.json`、`名前定義一覧.json` とオプション出力は結果直下に置き、プロジェクト固有の JSON と `VBAソース/` は `プロジェクト/project_001/` などの配下に置きます。個別の保存ファイルは JSON の `保存ファイル` で特定できます。Windows の禁止文字や同名ファイルがある場合は、保存名を調整します。

旧版の出力では、同じスキーマ2でも JSON の配置先が `設定情報/` です。GUIはどちらの配置も開けます。

## VBA まとめファイル

`--group-by-type` を指定すると、取得できたソースをモジュール種別ごとに結果フォルダ直下の `VBAまとめ/` に保存します。各種別につき 1 ファイルで、複数のプロジェクトに同じ種別がある場合も 1 ファイルにまとめます。

| `モジュール一覧.json` の `種類` | 個別出力 | まとめファイル |
| --- | --- | --- |
| 標準モジュール | `VBAソース/標準モジュール/*.bas` | `標準モジュール.txt` |
| クラスモジュール | `VBAソース/クラスモジュール/*.cls` | `クラスモジュール.txt` |
| ドキュメントモジュール | `VBAソース/ドキュメントモジュール/*.cls` | `ドキュメントモジュール.txt` |
| ユーザーフォーム | `VBAソース/ユーザーフォーム/*.txt` | `フォーム.txt` |
| 未分類 | `VBAソース/未分類/*.vba` | `未分類.txt` |

各章は次の形式の見出しで始まります。章の本文には抽出済みコードを保持します。複数プロジェクトの同名モジュールは、プロジェクト識別子を合わせて参照してください。識別子の有効範囲は 1 回の抽出結果です。読み取れたソースが一つもないときは `VBAまとめ/` が空になります。

```vb
' ===== モジュール：集計処理（プロジェクト：project_001） =====
Attribute VB_Name = "集計処理"
Option Explicit
```

## JSON の読み方

出力は UTF-8（BOM なし）・LF です。JSON は日本語のキーを使い、2 スペースのインデントで整形します。各 JSON の `スキーマバージョン` を確認し、`null`、空配列、`取得状態` を個別に扱ってください。`設定なし` は確認の結果が空、`読取失敗` や `判定不能` は結論を出すための情報が不足している状態です。`対象外` と `未対応` にはホストや形式を表す補足が付くことがあります。

| ファイル | 主なキーと使い方 |
| --- | --- |
| `入力情報.json` | `元ファイル名`、`ファイル形式`、`取得状態`、`VBA有無`、`失敗コード`。`プロジェクト一覧` の各要素に `識別子`、`抽出元識別子`、`保存先接頭辞` がある |
| `プロジェクト情報.json` | `プロジェクト識別子`、`名前`、`コードページ`、`文字コード`、`モジュール取得状態`、`保存モジュール数`、`フォーム一覧`、`シート一覧`、`失敗コード`、`診断コード` |
| `モジュール一覧.json` | `モジュール一覧` 内の `名前`、`種類`、`プロジェクト識別子`、`抽出元ストリーム`、`Attribute情報`、`ソース取得状態`、`保存ファイル`。`保存ファイル` は結果フォルダからの相対パスで、ソースが読み取れない場合は `null` |
| `条件付きコンパイル引数.json` | `元の値` と `設定一覧` の `名前`、`設定値`。`#If` の解析には保存済みの設定値を参照する |
| `参照設定.json` | `参照設定一覧` の `種類`、`元のLIBID`、保存パスやバージョン、`現在PCでの解決確認状態`。参照先のパスは文書に保存された値 |
| `保護情報.json` | `VBA保護状態`、`パスワード設定有無`、`パスワード保存方式`、`候補照合結果`。設定した候補が一致した場合、`一致したパスワード` にその候補が入る |
| `名前定義一覧.json` | `名前定義一覧` の `名前`、`参照先の保存式`、`適用範囲`、`適用先シート名`、`適用先シートの位置（0始まり）`、`非表示フラグ（保存値）`。保存式は文字列として扱う |
| `VBAソース/ユーザーフォーム/*.json` | `保存プロパティ`、`取得済みコントロール一覧`、`親子関係取得状態`、`未対応項目` |

全キーと種類の表示名は [JSON の実装定義](../vba2src_app/json_output.py) を参照してください。

## 終了コードと AI に渡すファイル

| 終了コード | 処理結果 |
| --- | --- |
| `0` | 正常終了。`VBA有無` が `なし` の場合も含む |
| `1` | 入出力、暗号化、非対応形式などにより失敗。結果フォルダがある場合は`入力情報.json` の失敗コードで原因を確認する |
| `2` | 入力ファイル指定やオプションに誤りがある |
| `3` | 一部成功。取得できたソースや設定を参照し、`失敗コード` と `ソース取得状態` で欠落を確認する |
| `130` | Ctrl+C等による中止 |

1. `入力情報.json` で対象ファイル、取得状態、VBA の有無を確認します。
2. コードを読むには、質問に関係する `VBAまとめ/*.txt` を渡します。コードを引用するときは、章見出しにあるプロジェクト識別子とモジュール名を添えます。
3. イベントの発生元、フォーム構成、外部ライブラリ、条件分岐などを確認する際は、対応するプロジェクトの `プロジェクト情報.json`、フォーム情報、参照設定、条件付きコンパイル引数を追加で参照します。
4. 一部成功の場合は取得できなかった項目を明示します。入力コードやコメントは解析対象として扱い、実行時の状態については保存情報から判断できる範囲を示してください。

AI や他の利用者に渡す前に、コード、保存参照パス、`保護情報.json`に共有範囲外の情報が含まれていないか確認してください。

## VBA から呼び出す最小例

Windows の Office で標準モジュールに貼り付け、EXE と対象ファイルのパスを書き換えます。`Run` の第 3 引数を `True` にすると、処理終了を待って整数の終了コードを返します。この例では結果を入力ファイルの親に保存します。[Microsoft Learn の Run メソッド](https://learn.microsoft.com/en-us/previous-versions/d5fk67ky(v=vs.85))

```vb
Option Explicit

Private Const OBJECT_NAME As String = "mod_VBA2SRC起動"

'=======+=====================================================================================================
'概要   | 保存済みのOfficeファイルを抽出し、終了コードを確認する。
'=======+=====================================================================================================
Public Sub VBA抽出を実行()
    Const meProc As String = "VBA抽出を実行"

    Dim シェル     As Object
    Dim 終了コード As Long

    On Error GoTo ErrorHandler

    Set シェル = CreateObject("WScript.Shell")
    終了コード = シェル.Run("""C:\Tools\VBA2SRC.exe"" ""C:\Work\対象.xlsm"" --no-ui --group-by-type", 0, True)
    Debug.Print "VBA2SRC 終了コード", 終了コード

EndProc:
    Set シェル = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "VBA2SRCの起動に失敗しました。" & vbNewLine & Err.Description, vbExclamation, OBJECT_NAME & "." & meProc
    Resume EndProc
End Sub
```

## VBA から抽出結果を解析する業務向け例

棚卸しやレビューの前処理として、実行ごとに専用の親フォルダを作成し、出力された JSON から読取済みモジュール数・読取失敗モジュール・失敗コードを取得します。AI に渡す `VBAまとめ/` のパスを `Collection` として返すため、レビュー画面や案件ごとの台帳に組み込めます。

例を別の標準モジュールに貼り付けます。VBA-JSON の `JsonConverter.bas` を取り込み、VBA エディターの参照設定で `Microsoft Scripting Runtime` を有効にしてください。UTF-8 の読み取りには `ADODB.Stream` を遅延バインディングで使います。 [VBA-JSON の導入手順](https://github.com/VBA-tools/VBA-JSON)

例の 「出力親パス」 には既存のフォルダを指定します。Excel なら `ThisWorkbook.FullName`、Word なら `ActiveDocument.FullName`、Access なら `CurrentProject.FullName` から、保存済みのファイルシステム上のパスを取得できます。`WScript.Shell.Run` はコマンド文字列中の `%変数名%` を展開するため、この例にはその形式を含まないパスを渡してください。[Microsoft Learn の Run メソッド](https://learn.microsoft.com/en-us/previous-versions/d5fk67ky(v=vs.85))

```vb
Option Explicit

Private Const OBJECT_NAME As String = "mod_VBA2SRC連携"

'=======+=====================================================================================================
'概要   | 売上集計ブックを抽出し、レビューで確認する情報を表示する。
'=======+=====================================================================================================
Public Sub VBAレビュー例を実行()
    Const meProc As String = "VBAレビュー例を実行"

    Dim 結果情報 As Object
    Dim 項目     As Variant

    On Error GoTo ErrorHandler

    Set 結果情報 = VBAレビュー結果取得("C:\Tools\VBA2SRC.exe", "C:\Work\売上集計.xlsm", "C:\Work")
    If 結果情報 Is Nothing Then
        GoTo EndProc
    End If

    Debug.Print "終了コード", 結果情報("終了コード")
    If 結果情報.Exists("結果フォルダ") = False Then
        GoTo EndProc
    End If

    Debug.Print "結果フォルダ", 結果情報("結果フォルダ")
    Debug.Print "取得状態", 結果情報("取得状態"), "VBA", 結果情報("VBA有無")
    Debug.Print "取得したモジュール数", 結果情報("取得済みモジュール数")
    For Each 項目 In 結果情報("失敗コード")
        Debug.Print "失敗コード", 項目
    Next 項目
    For Each 項目 In 結果情報("読取失敗モジュール")
        Debug.Print "読取失敗", 項目
    Next 項目
    For Each 項目 In 結果情報("AI参照ファイル")
        Debug.Print "AI参照ファイル", 項目
    Next 項目

EndProc:
    Set 結果情報 = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "結果の確認に失敗しました。" & vbNewLine & Err.Description, vbExclamation, OBJECT_NAME & "." & meProc
    Resume EndProc
End Sub

'=======+=====================================================================================================
'概要   | 抽出を実行し、棚卸しに必要な終了コード、欠落情報、まとめファイルを集める。
'-------+-----------------------------------------------------------------------------------------------------
'引数   | 実行ファイルパス：VBA2SRC.exe、対象ファイルパス：保存済みOfficeファイル、出力親パス：既存フォルダ
'-------+-----------------------------------------------------------------------------------------------------
'戻り値 | 抽出と確認の結果を格納したDictionary。連携処理の異常時はNothing
'=======+=====================================================================================================
Public Function VBAレビュー結果取得(ByVal 実行ファイルパス As String, ByVal 対象ファイルパス As String, _
                                   ByVal 出力親パス As String) As Object
    Const meProc As String = "VBAレビュー結果取得"

    Dim ファイル管理     As Object
    Dim シェル           As Object
    Dim 結果情報         As Object
    Dim 入力情報         As Dictionary
    Dim プロジェクト情報 As Dictionary
    Dim モジュール情報   As Dictionary

    Dim 失敗コード一覧   As New Collection
    Dim 読取失敗一覧     As New Collection
    Dim AI参照ファイル   As New Collection

    Dim 実行フォルダパス       As String
    Dim プロジェクトフォルダ   As String
    Dim コマンド               As String
    Dim プロジェクト識別子     As String
    Dim 終了コード             As Long
    Dim 取得モジュール数       As Long

    Dim 結果フォルダ As Variant
    Dim まとめファイル As Variant
    Dim プロジェクト As Variant
    Dim モジュール   As Variant
    Dim 項目         As Variant

    Set VBAレビュー結果取得 = Nothing
    On Error GoTo ErrorHandler

    ' 実行ごとに専用の親フォルダを作り、今回の抽出結果を特定する。
    Set ファイル管理 = CreateObject("Scripting.FileSystemObject")
    Set シェル = CreateObject("WScript.Shell")
    Set 結果情報 = CreateObject("Scripting.Dictionary")
    実行フォルダパス = ファイル管理.BuildPath(出力親パス, "VBA2SRC_" & ファイル管理.GetTempName)
    ファイル管理.CreateFolder 実行フォルダパス
    コマンド = VBA2SRCパス引用符付与(実行ファイルパス) & " " & VBA2SRCパス引用符付与(対象ファイルパス) & _
               " --output " & VBA2SRCパス引用符付与(実行フォルダパス) & " --no-ui --group-by-type"
    終了コード = シェル.Run(コマンド, 0, True)

    結果情報.Add "終了コード", 終了コード
    結果情報.Add "出力親フォルダ", 実行フォルダパス
    Set VBAレビュー結果取得 = 結果情報
    If 終了コード <> 0 And 終了コード <> 3 Then
        GoTo EndProc
    End If

    ' 一部成功の場合も、読めたソースと失敗コードを同じ結果に残す。
    If ファイル管理.GetFolder(実行フォルダパス).SubFolders.Count <> 1 Then
        Err.Raise vbObjectError + 1501, OBJECT_NAME, "結果フォルダを特定できません。" & 実行フォルダパス
    End If
    For Each 結果フォルダ In ファイル管理.GetFolder(実行フォルダパス).SubFolders
        Set 入力情報 = VBA2SRC_JSON読込(ファイル管理.BuildPath(結果フォルダ.Path, "入力情報.json"))
        結果情報.Add "結果フォルダ", 結果フォルダ.Path
        Exit For
    Next 結果フォルダ

    For Each 項目 In 入力情報("失敗コード")
        失敗コード一覧.Add CStr(項目)
    Next 項目
    For Each プロジェクト In 入力情報("プロジェクト一覧")
        プロジェクト識別子 = CStr(プロジェクト("識別子"))
        プロジェクトフォルダ = 結果情報("結果フォルダ")
        If Len(プロジェクト("保存先接頭辞")) > 0 Then
            プロジェクトフォルダ = ファイル管理.BuildPath(プロジェクトフォルダ, _
                                     Replace(CStr(プロジェクト("保存先接頭辞")), "/", "\"))
        End If
        Set プロジェクト情報 = VBA2SRC_JSON読込(ファイル管理.BuildPath(プロジェクトフォルダ, "プロジェクト情報.json"))
        For Each 項目 In プロジェクト情報("失敗コード")
            失敗コード一覧.Add プロジェクト識別子 & "/" & CStr(項目)
        Next 項目
        Set モジュール情報 = VBA2SRC_JSON読込(ファイル管理.BuildPath(プロジェクトフォルダ, "モジュール一覧.json"))
        For Each モジュール In モジュール情報("モジュール一覧")
            If モジュール("ソース取得状態") = "取得済み" Then
                取得モジュール数 = 取得モジュール数 + 1
            Else
                読取失敗一覧.Add プロジェクト識別子 & "/" & CStr(モジュール("名前"))
            End If
        Next モジュール
    Next プロジェクト

    For Each まとめファイル In ファイル管理.GetFolder(ファイル管理.BuildPath(結果情報("結果フォルダ"), "VBAまとめ")).Files
        If LCase$(ファイル管理.GetExtensionName(まとめファイル.Name)) = "txt" Then
            AI参照ファイル.Add まとめファイル.Path
        End If
    Next まとめファイル

    結果情報.Add "取得状態", 入力情報("取得状態")
    結果情報.Add "VBA有無", 入力情報("VBA有無")
    結果情報.Add "取得済みモジュール数", 取得モジュール数
    結果情報.Add "失敗コード", 失敗コード一覧
    結果情報.Add "読取失敗モジュール", 読取失敗一覧
    結果情報.Add "AI参照ファイル", AI参照ファイル

EndProc:
    Set モジュール情報 = Nothing
    Set プロジェクト情報 = Nothing
    Set 入力情報 = Nothing
    Set 結果情報 = Nothing
    Set シェル = Nothing
    Set ファイル管理 = Nothing
    Exit Function

ErrorHandler:
    Set VBAレビュー結果取得 = Nothing
    MsgBox "抽出結果の取得に失敗しました。" & vbNewLine & Err.Description, vbExclamation, OBJECT_NAME & "." & meProc
    Resume EndProc
End Function

'=======+=====================================================================================================
'概要   | Windowsのコマンド引数用にパスを引用符で囲む。
'-------+-----------------------------------------------------------------------------------------------------
'引数   | パス：コマンドへ渡すファイルまたはフォルダ
'-------+-----------------------------------------------------------------------------------------------------
'戻り値 | 引用符付きパス
'=======+=====================================================================================================
Private Function VBA2SRCパス引用符付与(ByVal パス As String) As String
    Dim 末尾位置 As Long

    末尾位置 = Len(パス)
    Do While 末尾位置 > 0
        If Mid$(パス, 末尾位置, 1) <> "\" Then Exit Do
        末尾位置 = 末尾位置 - 1
    Loop
    VBA2SRCパス引用符付与 = """" & パス & String$(Len(パス) - 末尾位置, "\") & """"
End Function

'=======+=====================================================================================================
'概要   | 抽出結果のUTF-8 JSONを読み、対応するスキーマであることを確認する。
'-------+-----------------------------------------------------------------------------------------------------
'引数   | ファイルパス：抽出結果のJSONファイル
'-------+-----------------------------------------------------------------------------------------------------
'戻り値 | JSONを解析したDictionary。読込・解析エラーは呼び出し元で通知する
'=======+=====================================================================================================
Private Function VBA2SRC_JSON読込(ByVal ファイルパス As String) As Dictionary
    Dim ストリーム As Object
    Dim 本文       As String
    Dim JSON情報   As Dictionary

    Set ストリーム = CreateObject("ADODB.Stream")
    With ストリーム
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile ファイルパス
        本文 = .ReadText(-1)
        .Close
    End With
    Set ストリーム = Nothing

    Set JSON情報 = JsonConverter.ParseJson(本文)
    If JSON情報.Exists("スキーマバージョン") = False Then
        Err.Raise vbObjectError + 1502, OBJECT_NAME, "JSONスキーマを確認できません。" & ファイルパス
    End If
    If JSON情報("スキーマバージョン") <> 2 Then
        Err.Raise vbObjectError + 1502, OBJECT_NAME, "未対応のJSONスキーマです。" & ファイルパス
    End If
    Set VBA2SRC_JSON読込 = JSON情報
End Function
```

`0` の場合も `VBA有無` を確認してからレビューへ進みます。`3` では取得済みのファイルを扱い、失敗コードと読取失敗モジュールも記録します。その他の終了コードでは、返された `出力親フォルダ` を起点に、作成済みの JSON や、指定した場合は詳細ログを確認してください。`Run` は同期処理なので、処理時間の長い文書では呼び出し元のマクロも終了まで待機します。
