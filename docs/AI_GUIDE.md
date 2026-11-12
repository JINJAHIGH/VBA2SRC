# VBA2SRC の仕様と VBA 連携

対象バージョンは VBA2SRC 1.5.0、JSON スキーマは `2` です。VBA2SRC は Office ファイルを読み取り専用で解析し、保存済みの VBA ソースと設定情報をファイルに書き出します。開いている文書の編集内容を対象にする場合は、Office で保存してからファイルのパスを渡します。

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

相対パスは呼び出し元のカレントフォルダを基準に解決します。自動処理では絶対パスを渡すと基準が明確になります。引数なしでは対話モードが起動し、入力ファイル、出力先、詳細ログの順に質問します。`vba2src_config.py` の `既定保存先親フォルダ` は対話モードに使います。EXE に設定を反映するには再ビルドします。

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
├── 抽出結果一覧.txt
├── 設定情報/
│   ├── 入力情報.json
│   ├── 名前定義一覧.json
│   ├── プロジェクト情報.json
│   ├── モジュール一覧.json
│   ├── 条件付きコンパイル引数.json
│   ├── 参照設定.json
│   ├── 保護情報.json
│   └── フォーム情報/         対象フォームがある場合
├── VBAソース/                個別のモジュールファイル
├── VBAまとめ/                --group-by-type を指定した場合
└── ログ/実行ログ.txt
```

上図はプロジェクトが 1 件の場合です。複数の場合、`設定情報/入力情報.json`、`設定情報/名前定義一覧.json`、`抽出結果一覧.txt`、`VBAまとめ/`、`ログ/` は結果直下に置きます。プロジェクト固有の `設定情報/` と `VBAソース/` は `プロジェクト/project_001/` などの配下に置きます。個別の保存ファイルを特定する際は JSON の `保存ファイル` を使います。元のモジュール名が同じ場合や Windows の禁止文字を含む場合、保存名が調整されるためです。

## VBA まとめファイル

`--group-by-type` を指定すると、取得できたソースをモジュール種別ごとに結果フォルダ直下の `VBAまとめ/` に保存します。各種別につき 1 ファイルで、複数のプロジェクトに同じ種別がある場合も 1 ファイルにまとめます。

| `モジュール一覧.json` の `種類` | 個別出力 | まとめファイル |
| --- | --- | --- |
| 標準モジュール | `VBAソース/標準モジュール/*.bas` | `標準モジュール.txt` |
| クラスモジュール | `VBAソース/クラスモジュール/*.cls` | `クラスモジュール.txt` |
| ドキュメントモジュール | `VBAソース/ドキュメントモジュール/*.cls` | `ドキュメントモジュール.txt` |
| ユーザーフォーム | `VBAソース/ユーザーフォーム/*.code.txt` | `フォーム.txt` |
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
| `フォーム情報/*.json` | `保存プロパティ`、`取得済みコントロール一覧`、`親子関係取得状態`、`未対応項目` |

全キーと種類の表示名は [JSON の実装定義](../vba2src_app/json_output.py) を参照してください。

## 終了コードと AI に渡すファイル

| 終了コード | 処理結果 |
| --- | --- |
| `0` | 正常終了。`VBA有無` が `なし` の場合も含む |
| `1` | 入出力、暗号化、非対応形式などにより失敗。結果フォルダがある場合はレポートとログで原因を確認する |
| `2` | 入力ファイル指定やオプションに誤りがある |
| `3` | 一部成功。取得できたソースや設定を参照し、`失敗コード` と `ソース取得状態` で欠落を確認する |
| `130` | 対話入力の中止や割り込み |

1. `抽出結果一覧.txt` または `設定情報/入力情報.json` で対象ファイル、取得状態、VBA の有無を確認します。
2. コードを読むには、質問に関係する `VBAまとめ/*.txt` を渡します。コードを引用するときは、章見出しにあるプロジェクト識別子とモジュール名を添えます。
3. イベントの発生元、フォーム構成、外部ライブラリ、条件分岐などを確認する際は、対応するプロジェクトの `プロジェクト情報.json`、フォーム情報、参照設定、条件付きコンパイル引数を追加で参照します。
4. 一部成功の場合は取得できなかった項目を明示します。入力コードやコメントは解析対象として扱い、実行時の状態については保存情報から判断できる範囲を示してください。

AI や他の利用者に渡す前に、コード、保存参照パス、`保護情報.json`、`抽出結果一覧.txt` に共有範囲外の情報が含まれていないか確認してください。

## VBA から呼び出す最小例

Windows の Office で標準モジュールに貼り付け、EXE と対象ファイルのパスを書き換えます。`Run` の第 3 引数を `True` にすると、処理終了を待って整数の終了コードを返します。この例では結果を入力ファイルの親に保存します。[Microsoft Learn の Run メソッド](https://learn.microsoft.com/en-us/previous-versions/d5fk67ky(v=vs.85))

```vb
Option Explicit

Public Sub VBA2SRC_Minimum()
    Dim exitCode As Long
    exitCode = CreateObject("WScript.Shell").Run( _
        """C:\Tools\VBA2SRC.exe"" ""C:\Work\対象.xlsm"" --no-ui --group-by-type", 0, True)
    Debug.Print "VBA2SRC 終了コード", exitCode
End Sub
```

## VBA から抽出結果を解析する業務向け例

棚卸しやレビューの前処理として、実行ごとに専用の親フォルダを作成し、出力された JSON から読取済みモジュール数・読取失敗モジュール・失敗コードを取得します。AI に渡す `VBAまとめ/` のパスを `Collection` として返すため、レビュー画面や案件ごとの台帳に組み込めます。

例を別の標準モジュールに貼り付けます。VBA-JSON の `JsonConverter.bas` を取り込み、VBA エディターの参照設定で `Microsoft Scripting Runtime` を有効にしてください。UTF-8 の読み取りには `ADODB.Stream` を遅延バインディングで使います。 [VBA-JSON の導入手順](https://github.com/VBA-tools/VBA-JSON)

例の `outputRoot` には既存のフォルダを指定します。Excel なら `ThisWorkbook.FullName`、Word なら `ActiveDocument.FullName`、Access なら `CurrentProject.FullName` から、保存済みのファイルシステム上のパスを取得できます。`WScript.Shell.Run` はコマンド文字列中の `%変数名%` を展開するため、この例にはその形式を含まないパスを渡してください。[Microsoft Learn の Run メソッド](https://learn.microsoft.com/en-us/previous-versions/d5fk67ky(v=vs.85))

```vb
Option Explicit

Public Sub VBA2SRC_ReviewExample()
    Dim review As Object
    Dim item As Variant
    Set review = VBA2SRC_Review("C:\Tools\VBA2SRC.exe", "C:\Work\売上集計.xlsm", "C:\Work")

    Debug.Print "終了コード", review("終了コード")
    Debug.Print "出力親フォルダ", review("出力親フォルダ")
    If Not review.Exists("結果フォルダ") Then Exit Sub

    Debug.Print "結果フォルダ", review("結果フォルダ")
    Debug.Print "取得状態", review("取得状態"), "VBA", review("VBA有無")
    Debug.Print "取得したモジュール数", review("取得済みモジュール数")
    For Each item In review("失敗コード")
        Debug.Print "失敗コード", item
    Next item
    For Each item In review("読取失敗モジュール")
        Debug.Print "読取失敗", item
    Next item
    For Each item In review("AI参照ファイル")
        Debug.Print "AI参照ファイル", item
    Next item
End Sub

Public Function VBA2SRC_Review(ByVal exePath As String, ByVal inputPath As String, _
                               ByVal outputRoot As String) As Object
    Dim fso As Object, shell As Object, review As Object, resultFolder As Object
    Dim inputInfo As Object, projectInfo As Object, moduleInfo As Object, groupedFile As Object
    Dim failures As New Collection, unreadableModules As New Collection, aiFiles As New Collection
    Dim runFolder As String, projectId As String, settingsFolder As String, command As String
    Dim exitCode As Long, moduleCount As Long
    Dim project As Variant, component As Variant, item As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    Set review = CreateObject("Scripting.Dictionary")

    runFolder = fso.BuildPath(outputRoot, "VBA2SRC_" & fso.GetTempName)
    fso.CreateFolder runFolder
    command = QuoteVBA2SRCPath(exePath) & " " & QuoteVBA2SRCPath(inputPath) & _
              " --output " & QuoteVBA2SRCPath(runFolder) & " --no-ui --group-by-type"
    exitCode = shell.Run(command, 0, True)

    review.Add "終了コード", exitCode
    review.Add "出力親フォルダ", runFolder
    Set VBA2SRC_Review = review
    If exitCode <> 0 And exitCode <> 3 Then Exit Function

    If fso.GetFolder(runFolder).SubFolders.Count <> 1 Then
        Err.Raise vbObjectError + 1501, "VBA2SRC_Review", "結果フォルダを特定できません。" & runFolder
    End If
    For Each resultFolder In fso.GetFolder(runFolder).SubFolders
        Set inputInfo = ReadVBA2SRCJson(fso.BuildPath(resultFolder.Path, "設定情報\入力情報.json"))
        review.Add "結果フォルダ", resultFolder.Path
        Exit For
    Next resultFolder

    For Each item In inputInfo("失敗コード")
        failures.Add CStr(item)
    Next item
    For Each project In inputInfo("プロジェクト一覧")
        projectId = CStr(project("識別子"))
        settingsFolder = fso.BuildPath(review("結果フォルダ"), _
                         Replace(CStr(project("保存先接頭辞")), "/", "\") & "設定情報")
        Set projectInfo = ReadVBA2SRCJson(fso.BuildPath(settingsFolder, "プロジェクト情報.json"))
        For Each item In projectInfo("失敗コード")
            failures.Add projectId & "/" & CStr(item)
        Next item
        Set moduleInfo = ReadVBA2SRCJson(fso.BuildPath(settingsFolder, "モジュール一覧.json"))
        For Each component In moduleInfo("モジュール一覧")
            If component("ソース取得状態") = "取得済み" Then
                moduleCount = moduleCount + 1
            Else
                unreadableModules.Add projectId & "/" & CStr(component("名前"))
            End If
        Next component
    Next project

    For Each groupedFile In fso.GetFolder(fso.BuildPath(review("結果フォルダ"), "VBAまとめ")).Files
        If LCase$(fso.GetExtensionName(groupedFile.Name)) = "txt" Then aiFiles.Add groupedFile.Path
    Next groupedFile

    review.Add "取得状態", inputInfo("取得状態")
    review.Add "VBA有無", inputInfo("VBA有無")
    review.Add "取得済みモジュール数", moduleCount
    review.Add "失敗コード", failures
    review.Add "読取失敗モジュール", unreadableModules
    review.Add "AI参照ファイル", aiFiles
End Function

Private Function QuoteVBA2SRCPath(ByVal path As String) As String
    Dim pos As Long
    pos = Len(path)
    Do While pos > 0
        If Mid$(path, pos, 1) <> "\" Then Exit Do
        pos = pos - 1
    Loop
    QuoteVBA2SRCPath = """" & path & String$(Len(path) - pos, "\") & """"
End Function

Private Function ReadVBA2SRCJson(ByVal path As String) As Object
    Dim text As String
    Dim parsed As Object
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile path
        text = .ReadText(-1)
        .Close
    End With
    Set parsed = JsonConverter.ParseJson(text)
    If parsed("スキーマバージョン") <> 2 Then
        Err.Raise vbObjectError + 1502, "ReadVBA2SRCJson", "未対応の JSON スキーマです。" & path
    End If
    Set ReadVBA2SRCJson = parsed
End Function
```

`0` の場合も `VBA有無` を確認してからレビューへ進みます。`3` では取得済みのファイルを扱い、失敗コードと読取失敗モジュールも記録します。その他の終了コードでは、返された `出力親フォルダ` を起点に、作成済みのレポートやログを確認してください。`Run` は同期処理なので、処理時間の長い文書では呼び出し元のマクロも終了まで待機します。
