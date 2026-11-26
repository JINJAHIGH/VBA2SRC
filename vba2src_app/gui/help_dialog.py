"""画面で使える操作と、コピーできるCLI例を表示する。"""

import tkinter as tk
from tkinter import ttk


class 使い方画面(tk.Toplevel):
    def __init__(自身, 親):
        super().__init__(親)
        自身.title("VBA2SRC | 使い方")
        自身.transient(親)
        自身.resizable(False, False)
        自身.通知 = tk.StringVar()

        枠 = ttk.Frame(自身, style="Workspace.TFrame", padding=(20, 16))
        枠.pack(fill="both", expand=True)
        ttk.Label(枠, text="使い方", style="Heading.TLabel").pack(anchor="w", pady=(0, 12))
        タブ = ttk.Notebook(枠)
        タブ.pack(fill="both", expand=True)
        操作 = ttk.Frame(タブ, padding=(16, 14))
        コマンド = ttk.Frame(タブ, padding=(16, 14))
        タブ.add(操作, text="画面で操作する")
        タブ.add(コマンド, text="コマンドライン")
        タブ.enable_traversal()

        for 番号, (見出し, 手順) in enumerate(
            (
                (
                    "VBAを抽出する",
                    (
                        "Excel・Word・Accessのファイルを対象欄にドロップします。",
                        "パスを入力するか、「参照」でファイルを選ぶこともできます。",
                        "出力先が空欄なら対象ファイルと同じフォルダです。「出力を実行」で開始します。",
                    ),
                ),
                (
                    "出力オプションを選ぶ",
                    (
                        "「モジュール別テキストファイル生成」で、VBAまとめに種類別のコードを保存します。",
                        "処理の記録が必要な場合は「詳細ログを保存する」をオンにします。",
                    ),
                ),
                (
                    "結果を確認する",
                    (
                        "抽出後は結果が自動で開きます。タブを切り替え、一覧から詳細を確認できます。",
                        "以前の結果は、「抽出結果を確認」で入力情報.jsonがあるフォルダを選びます。",
                    ),
                ),
            ),
            start=1,
        ):
            if 番号 > 1:
                ttk.Separator(操作).pack(fill="x", pady=12)
            ttk.Label(操作, text=f"{番号}. {見出し}", style="Section.TLabel").pack(anchor="w", pady=(0, 2))
            for 内容 in 手順:
                ttk.Label(操作, text=内容, wraplength=590, justify="left").pack(anchor="w", pady=(4, 0))

        ttk.Label(コマンド, text="画面を開かずに抽出する", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            コマンド,
            text="PowerShellで EXE のあるフォルダを開き、例のファイル名と出力先を書き換えて実行します。",
            wraplength=550,
        ).pack(anchor="w", pady=(5, 10))
        for 見出し, 例 in (
            ("対象ファイルと出力先を指定", r'.\VBA2SRC.exe "C:\Work\売上.xlsm" -o "C:\Work\結果" --no-ui'),
            ("使用できるオプションを表示", r".\VBA2SRC.exe --help"),
        ):
            ttk.Label(コマンド, text=見出し, style="Muted.TLabel").pack(anchor="w", pady=(5, 3))
            行 = ttk.Frame(コマンド)
            行.pack(fill="x", pady=(0, 7))
            欄 = ttk.Entry(行, font="TkFixedFont", width=66)
            欄.insert(0, 例)
            欄.state(["readonly"])
            欄.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ttk.Button(行, text="コピー", width=8, command=lambda 内容=例: 自身.コピー(内容)).pack(
                side="right"
            )

        ttk.Label(
            コマンド,
            text="OfficeファイルをEXEのアイコンにドロップすると、そのファイルの隣に結果を作成します。",
            style="Muted.TLabel",
            wraplength=550,
        ).pack(anchor="w", pady=(5, 0))
        操作欄 = ttk.Frame(枠, style="Workspace.TFrame")
        操作欄.pack(fill="x", pady=(12, 0))
        ttk.Label(操作欄, textvariable=自身.通知, style="WorkspaceMuted.TLabel").pack(side="left")
        ttk.Button(操作欄, text="閉じる", width=8, style="Important.TButton", command=自身.destroy).pack(
            side="right"
        )
        自身.bind("<Escape>", lambda イベント: 自身.destroy())

    def コピー(自身, 内容):
        自身.clipboard_clear()
        自身.clipboard_append(内容)
        自身.通知.set("コマンドをコピーしました。")
