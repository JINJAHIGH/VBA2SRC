"""保存先と書体の初期値を編集する設定画面。"""

import tkinter as tk
from tkinter import filedialog, ttk

from ..preferences import GUI設定, 設定を保存
from .password_settings import 候補設定

システム既定 = "自動（推奨）"


class 設定画面(tk.Toplevel):
    def __init__(自身, 親, 設定, フォント一覧, 既定保存先, 確定):
        super().__init__(親)
        自身.title("VBA2SRC | 設定")
        自身.transient(親)
        自身.resizable(False, False)
        自身.確定 = 確定
        自身.既定保存先 = 既定保存先
        自身.保存先 = tk.StringVar(value=設定.保存先)
        自身.画面書体 = tk.StringVar(value=設定.画面フォント or システム既定)
        自身.画面サイズ = tk.StringVar(value=設定.画面サイズ)
        自身.コード書体 = tk.StringVar(value=設定.コードフォント or システム既定)
        自身.コードサイズ = tk.StringVar(value=設定.コードサイズ)
        自身.エラー = tk.StringVar()

        枠 = ttk.Frame(自身, style="Workspace.TFrame", padding=22)
        枠.pack(fill="both", expand=True)
        ttk.Label(枠, text="設定", style="Heading.TLabel").pack(anchor="w", pady=(0, 14))
        タブ = ttk.Notebook(枠)
        タブ.pack(fill="both", expand=True)
        基本 = ttk.Frame(タブ, padding=10)
        タブ.add(基本, text="基本")
        自身.保護設定 = 候補設定(タブ, 設定.VBA候補)
        タブ.add(自身.保護設定, text="VBA保護")
        タブ.enable_traversal()

        保存カード = ttk.Frame(基本, style="Card.TFrame", padding=14)
        保存カード.pack(fill="x", pady=(0, 10))
        ttk.Label(保存カード, text="既定の出力先", style="Section.TLabel").pack(anchor="w")
        保存欄 = ttk.Frame(保存カード)
        保存欄.pack(fill="x", pady=(6, 5))
        ttk.Entry(保存欄, textvariable=自身.保存先, width=45).pack(
            side="left", padx=(0, 8), fill="x", expand=True
        )
        ttk.Button(保存欄, text="参照…", command=自身.フォルダを選択).pack(side="right")
        ttk.Label(
            保存カード, text=f"空欄の場合は {既定保存先} を使用します。", style="Muted.TLabel", wraplength=480
        ).pack(anchor="w")

        書体カード = ttk.Frame(基本, style="Card.TFrame", padding=14)
        書体カード.pack(fill="x")
        書体カード.columnconfigure(1, weight=1)
        ttk.Label(書体カード, text="フォント", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Label(書体カード, text="サイズ（pt）", style="Muted.TLabel").grid(row=0, column=2, sticky="w")
        for 行, (名前, 書体, サイズ, 上限下限) in enumerate(
            (
                ("画面", 自身.画面書体, 自身.画面サイズ, (9, 14)),
                ("VBAコード", 自身.コード書体, 自身.コードサイズ, (8, 24)),
            ),
            start=1,
        ):
            ttk.Label(書体カード, text=名前).grid(row=行, column=0, sticky="w")
            ttk.Combobox(
                書体カード,
                state="readonly",
                textvariable=書体,
                values=(システム既定, *フォント一覧),
                width=24,
            ).grid(row=行, column=1, sticky="ew", padx=12, pady=5)
            ttk.Spinbox(書体カード, textvariable=サイズ, from_=上限下限[0], to=上限下限[1], width=5).grid(
                row=行, column=2, sticky="w"
            )
        ttk.Label(
            書体カード,
            text="自動では、画面はヒラギノ角ゴシックまたは BIZ UDPゴシック、コードは MS ゴシックを優先します。",
            style="Muted.TLabel",
            wraplength=445,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(7, 0))

        ttk.Label(枠, textvariable=自身.エラー, style="Error.TLabel", wraplength=475).pack(
            anchor="w", pady=(6, 0)
        )
        操作欄 = ttk.Frame(枠, style="Workspace.TFrame")
        操作欄.pack(fill="x", pady=(6, 0))
        ttk.Button(操作欄, text="既定値に戻す", command=自身.初期化).pack(side="left")
        ttk.Button(操作欄, text="保存", style="Accent.TButton", command=自身.保存).pack(side="right")
        ttk.Button(操作欄, text="キャンセル", command=自身.destroy).pack(side="right", padx=(0, 8))
        自身.bind("<Escape>", lambda イベント: 自身.destroy())

    def フォルダを選択(自身):
        パス = filedialog.askdirectory(
            parent=自身, title="既定の出力先を選択", initialdir=自身.保存先.get() or 自身.既定保存先
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
        自身.エラー.set("")

    def 保存(自身):
        try:
            画面サイズ = int(自身.画面サイズ.get())
            コードサイズ = int(自身.コードサイズ.get())
            if not 9 <= 画面サイズ <= 14 or not 8 <= コードサイズ <= 24:
                raise ValueError
        except ValueError:
            自身.エラー.set("フォントサイズは、画面は9〜14、VBAコードは8〜24の整数で入力してください。")
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
            自身.エラー.set(f"設定を保存できませんでした。{例外情報}")
            return
        自身.destroy()
        自身.確定(設定)
