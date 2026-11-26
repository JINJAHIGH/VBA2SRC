"""保存先と書体の初期値を編集する設定画面。"""

import tkinter as tk
from tkinter import filedialog, ttk

from ..preferences import GUI設定, 設定を保存
from .password_settings import 候補設定

システム既定 = "自動（推奨）"


class 設定画面(tk.Toplevel):
    def __init__(自身, 親, 設定, フォント一覧, 確定):
        super().__init__(親)
        自身.title("VBA2SRC | 設定")
        自身.transient(親)
        自身.resizable(False, False)
        自身.確定 = 確定
        自身.保存先 = tk.StringVar(value=設定.保存先)
        自身.画面書体 = tk.StringVar(value=設定.画面フォント or システム既定)
        自身.画面サイズ = tk.StringVar(value=設定.画面サイズ)
        自身.コード書体 = tk.StringVar(value=設定.コードフォント or システム既定)
        自身.コードサイズ = tk.StringVar(value=設定.コードサイズ)
        自身.エラー = tk.StringVar()

        枠 = ttk.Frame(自身, style="Workspace.TFrame", padding=(20, 16))
        枠.pack(fill="both", expand=True)
        ttk.Label(枠, text="設定", style="Heading.TLabel").pack(anchor="w", pady=(0, 10))
        自身.タブ = ttk.Notebook(枠)
        自身.タブ.pack(fill="both", expand=True)
        基本 = ttk.Frame(自身.タブ, padding=(16, 14))
        自身.タブ.add(基本, text="基本")
        自身.保護設定 = 候補設定(自身.タブ, 設定.VBA候補)
        自身.タブ.add(自身.保護設定, text="VBA保護")
        自身.タブ.enable_traversal()

        ttk.Label(基本, text="既定の出力先", style="Section.TLabel").pack(anchor="w")
        保存欄 = ttk.Frame(基本)
        保存欄.pack(fill="x", pady=(6, 4))
        ttk.Entry(保存欄, textvariable=自身.保存先, width=34).pack(
            side="left", padx=(0, 8), fill="x", expand=True
        )
        ttk.Button(保存欄, text="参照", width=8, command=自身.フォルダを選択).pack(side="right")
        ttk.Label(
            基本,
            text="空欄の場合は、対象ファイルと同じフォルダに出力します。",
            style="Muted.TLabel",
            wraplength=440,
        ).pack(anchor="w")

        ttk.Separator(基本).pack(fill="x", pady=13)
        書体欄 = ttk.Frame(基本)
        書体欄.pack(fill="x")
        書体欄.columnconfigure(1, weight=1)
        ttk.Label(書体欄, text="フォント", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(書体欄, text="サイズ（pt）", style="Muted.TLabel").grid(row=0, column=2, sticky="w")
        for 行, (名前, 書体, サイズ, 範囲) in enumerate(
            (
                ("画面", 自身.画面書体, 自身.画面サイズ, (9, 14)),
                ("VBAコード", 自身.コード書体, 自身.コードサイズ, (8, 24)),
            ),
            start=1,
        ):
            ttk.Label(書体欄, text=名前).grid(row=行, column=0, sticky="w")
            ttk.Combobox(
                書体欄,
                state="readonly",
                textvariable=書体,
                values=(システム既定, *フォント一覧),
                width=21,
            ).grid(row=行, column=1, sticky="ew", padx=12, pady=(6, 0))
            ttk.Spinbox(書体欄, textvariable=サイズ, from_=範囲[0], to=範囲[1], width=5).grid(
                row=行, column=2, sticky="w", pady=(6, 0)
            )

        自身.エラー欄 = ttk.Label(枠, textvariable=自身.エラー, style="Error.TLabel", wraplength=455)
        操作欄 = ttk.Frame(枠, style="Workspace.TFrame")
        自身.操作欄 = 操作欄
        操作欄.pack(fill="x", pady=(12, 0))
        ttk.Button(操作欄, text="すべて既定値に戻す", style="Reset.TButton", command=自身.初期化).pack(
            side="left"
        )
        ttk.Button(操作欄, text="保存", width=8, style="Accent.TButton", command=自身.保存).pack(side="right")
        ttk.Button(操作欄, text="キャンセル", width=10, command=自身.destroy).pack(side="right", padx=(0, 8))
        自身.bind("<Escape>", lambda イベント: 自身.destroy())
        自身.update_idletasks()
        自身.タブの高さを調整()
        自身.タブ.bind("<<NotebookTabChanged>>", 自身.タブの高さを調整)

    def タブの高さを調整(自身, イベント=None):
        選択 = 自身.タブ.nametowidget(自身.タブ.select())
        自身.タブ.configure(height=選択.winfo_reqheight())

    def エラーを表示(自身, メッセージ):
        自身.エラー.set(メッセージ)
        自身.エラー欄.pack(before=自身.操作欄, anchor="w", pady=(8, 0))

    def フォルダを選択(自身):
        パス = filedialog.askdirectory(
            parent=自身, title="既定の出力先を選択", initialdir=自身.保存先.get() or None
        )
        if パス:
            自身.保存先.set(パス)

    def 初期化(自身):
        自身.保存先.set("")
        自身.画面書体.set(システム既定)
        自身.コード書体.set(システム既定)
        自身.画面サイズ.set(10)
        自身.コードサイズ.set(10)
        自身.保護設定.初期化()
        自身.エラー欄.pack_forget()

    def 保存(自身):
        try:
            画面サイズ = int(自身.画面サイズ.get())
            コードサイズ = int(自身.コードサイズ.get())
            if not 9 <= 画面サイズ <= 14 or not 8 <= コードサイズ <= 24:
                raise ValueError
        except ValueError:
            自身.エラーを表示("フォントサイズは、画面は9〜14、VBAコードは8〜24の整数で入力してください。")
            return

        保存先 = 自身.保存先.get()
        if len(保存先) >= 2 and 保存先[0] == 保存先[-1] == '"':
            保存先 = 保存先[1:-1]
        画面書体 = 自身.画面書体.get()
        コード書体 = 自身.コード書体.get()
        設定 = GUI設定(
            保存先,
            "" if 画面書体 == システム既定 else 画面書体,
            画面サイズ,
            "" if コード書体 == システム既定 else コード書体,
            コードサイズ,
            自身.保護設定.設定値(),
        )
        try:
            設定を保存(設定)
        except OSError as 例外情報:
            自身.エラーを表示(f"設定を保存できませんでした。{例外情報}")
            return
        自身.destroy()
        自身.確定(設定)
