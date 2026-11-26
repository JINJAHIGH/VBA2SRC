"""VBA保護パスワード候補を登録順に編集する。"""

import tkinter as tk
from tkinter import ttk

from ..candidate_config import 候補設定エラー, 既定候補を読む
from ..console import 表示値


class 候補設定(ttk.Frame):
    def __init__(自身, 親, 保存済み候補):
        super().__init__(親, padding=(16, 14))
        自身.設定案内 = tk.StringVar()
        自身.既定候補を反映()
        自身.候補 = list(保存済み候補 if 保存済み候補 is not None else 自身.既定候補)
        自身.入力値 = tk.StringVar()
        自身.案内 = tk.StringVar()

        ttk.Label(自身, text="パスワード候補", style="Section.TLabel").pack(anchor="w")
        ttk.Label(自身, text="一覧の上から順に照合します。", style="Muted.TLabel").pack(
            anchor="w", pady=(3, 8)
        )

        ttk.Label(自身, textvariable=自身.設定案内, style="Muted.TLabel", wraplength=440).pack(
            anchor="w", pady=(0, 8)
        )
        一覧枠 = ttk.Frame(自身)
        一覧枠.pack(fill="x")
        一覧枠.columnconfigure(0, weight=1)
        自身.一覧 = ttk.Treeview(一覧枠, columns=("候補",), show="", selectmode="browse", height=4)
        自身.一覧.column("候補", width=370)
        縦 = ttk.Scrollbar(一覧枠, orient="vertical", command=自身.一覧.yview)
        自身.一覧.configure(yscrollcommand=縦.set)
        自身.一覧.grid(row=0, column=0, sticky="ew")
        縦.grid(row=0, column=1, sticky="ns")
        自身.一覧.bind("<<TreeviewSelect>>", 自身.選択を反映)

        ttk.Label(自身, text="候補の入力・編集", style="Muted.TLabel").pack(anchor="w", pady=(8, 3))
        自身.入力 = ttk.Entry(自身, textvariable=自身.入力値)
        自身.入力.pack(fill="x", pady=(0, 8))

        操作 = ttk.Frame(自身)
        操作.pack(fill="x")
        for 名前, スタイル, 処理 in (
            ("追加", "Add.TButton", 自身.追加),
            ("更新", "Update.TButton", 自身.更新),
            ("削除", "Delete.TButton", 自身.削除),
            ("リセット", "Reset.TButton", 自身.初期化),
        ):
            ttk.Button(操作, text=名前, style=スタイル, width=7, command=処理).pack(side="left", padx=(0, 6))
        ttk.Label(自身, textvariable=自身.案内, style="Muted.TLabel", wraplength=440).pack(
            anchor="w", pady=(8, 0)
        )
        自身.一覧を更新()

    def 一覧を更新(自身):
        自身.一覧.delete(*自身.一覧.get_children())
        for 番号, 候補 in enumerate(自身.候補):
            自身.一覧.insert("", "end", iid=str(番号), values=(f"{番号 + 1}.   {表示値(候補)}",))
        自身.案内.set(f"{len(自身.候補)} 件。変更を反映するには「保存」を押してください。")
        自身.入力値.set("")

    def 選択を反映(自身, イベント):
        行 = 自身.一覧.selection()
        if 行:
            自身.入力値.set(自身.候補[int(行[0])])

    def 追加(自身):
        値 = 自身.入力値.get()
        if 値:
            自身.候補.append(値)
            自身.一覧を更新()
            自身.一覧.see(str(len(自身.候補) - 1))
        else:
            自身.案内.set("候補を入力してから追加してください。")

    def 更新(自身):
        行 = 自身.一覧.selection()
        値 = 自身.入力値.get()
        if 行 and 値:
            自身.候補[int(行[0])] = 値
            自身.一覧を更新()
        else:
            自身.案内.set("行を選択し、新しい値を入力してから更新してください。")

    def 削除(自身):
        行 = 自身.一覧.selection()
        if 行:
            del 自身.候補[int(行[0])]
            自身.一覧を更新()
        else:
            自身.案内.set("削除する候補の行を選択してください。")

    def 既定候補を反映(自身):
        try:
            自身.既定候補 = list(既定候補を読む())
        except 候補設定エラー as 例外情報:
            自身.既定候補 = []
            自身.設定案内.set(str(例外情報))
        else:
            自身.設定案内.set("既定候補は vba2src_config.py から読み込みます。")

    def 初期化(自身):
        自身.既定候補を反映()
        自身.候補 = list(自身.既定候補)
        自身.一覧を更新()

    def 設定値(自身):
        return None if 自身.候補 == 自身.既定候補 else list(自身.候補)
