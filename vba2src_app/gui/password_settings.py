"""保存先、書体、既知のVBA保護パスワード候補を編集する。"""

import tkinter as tk
from tkinter import ttk

import vba2src_config as config

システム既定 = "自動（推奨）"


class 候補設定(ttk.Frame):
    """候補を登録順に編集する。初期状態では候補の値を伏せて表示する。"""

    def __init__(自身, 親, 保存済み候補):
        super().__init__(親, padding=12)
        自身.既定候補 = list(config.VBA保護パスワード候補)
        自身.候補 = list(保存済み候補 if 保存済み候補 is not None else 自身.既定候補)
        自身.既定使用 = tk.BooleanVar(value=保存済み候補 is None)
        自身.表示切替 = tk.BooleanVar()
        自身.入力値 = tk.StringVar()
        自身.案内 = tk.StringVar()

        ttk.Checkbutton(自身, text="既定の候補を使う", variable=自身.既定使用, command=自身.一覧を更新).pack(
            anchor="w", pady=(0, 5)
        )
        ttk.Label(
            自身,
            text="候補は上から順に照合します。独自の候補はこのPCの利用者設定に保存します。",
            style="Muted.TLabel",
            wraplength=440,
        ).pack(anchor="w", pady=(0, 8))

        一覧枠 = ttk.Frame(自身)
        一覧枠.pack(fill="x")
        一覧枠.columnconfigure(0, weight=1)
        自身.一覧 = ttk.Treeview(一覧枠, columns=("候補",), show="headings", selectmode="browse", height=4)
        自身.一覧.heading("候補", text="パスワード候補", anchor="w")
        自身.一覧.column("候補", width=320)
        縦 = ttk.Scrollbar(一覧枠, orient="vertical", command=自身.一覧.yview)
        自身.一覧.configure(yscrollcommand=縦.set)
        自身.一覧.grid(row=0, column=0, sticky="ew")
        縦.grid(row=0, column=1, sticky="ns")
        自身.一覧.bind("<<TreeviewSelect>>", 自身.選択を反映)

        入力欄 = ttk.Frame(自身)
        入力欄.pack(fill="x", pady=(10, 5))
        自身.入力 = ttk.Entry(入力欄, textvariable=自身.入力値, show="•")
        自身.入力.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Checkbutton(入力欄, text="値を表示", variable=自身.表示切替, command=自身.表示を切替).pack(
            side="right"
        )

        操作 = ttk.Frame(自身)
        操作.pack(fill="x")
        自身.編集部品 = [自身.入力]
        for 名前, 処理 in (("追加", 自身.追加), ("更新", 自身.更新), ("削除", 自身.削除)):
            ボタン = ttk.Button(操作, text=名前, width=5, command=処理)
            ボタン.pack(side="left", padx=(0, 5))
            自身.編集部品.append(ボタン)
        ttk.Label(自身, textvariable=自身.案内, style="Muted.TLabel", wraplength=440).pack(
            anchor="w", pady=(6, 0)
        )
        自身.一覧を更新()

    def 一覧を更新(自身):
        既定 = 自身.既定使用.get()
        for 部品 in 自身.編集部品:
            部品.state(["disabled" if 既定 else "!disabled"])
        自身.一覧.delete(*自身.一覧.get_children())
        for 番号, 候補 in enumerate(自身.既定候補 if 既定 else 自身.候補):
            表示 = 候補 if 自身.表示切替.get() else "••••••"
            自身.一覧.insert("", "end", iid=str(番号), values=(f"{番号 + 1}.   {表示}",))
        自身.案内.set(
            f"{len(自身.既定候補 if 既定 else 自身.候補)} 件。"
            + (
                "チェックを外すと編集できます。"
                if 既定
                else "追加・更新・削除した後に「保存」を押してください。"
            )
        )
        自身.入力値.set("")

    def 選択を反映(自身, イベント):
        行 = 自身.一覧.selection()
        if 行 and not 自身.既定使用.get():
            自身.入力値.set(自身.候補[int(行[0])])

    def 表示を切替(自身):
        選択行 = 自身.一覧.selection()
        入力値 = 自身.入力値.get()
        自身.入力.configure(show="" if 自身.表示切替.get() else "•")
        自身.一覧を更新()
        if 選択行 and 自身.一覧.exists(選択行[0]):
            自身.一覧.selection_set(選択行[0])
        自身.入力値.set(入力値)

    def 追加(自身):
        値 = 自身.入力値.get()
        if 値:
            自身.候補.append(値)
            自身.一覧を更新()
        else:
            自身.案内.set("候補を入力してから追加してください。")

    def 更新(自身):
        行 = 自身.一覧.selection()
        値 = 自身.入力値.get()
        if 行 and 値:
            自身.候補[int(行[0])] = 値
            自身.一覧を更新()
        else:
            自身.案内.set("候補の行を選択し、新しい値を入力してから更新してください。")

    def 削除(自身):
        行 = 自身.一覧.selection()
        if 行:
            del 自身.候補[int(行[0])]
            自身.一覧を更新()
        else:
            自身.案内.set("削除する候補の行を選択してください。")

    def 初期化(自身):
        自身.既定使用.set(True)
        自身.候補 = list(自身.既定候補)
        自身.表示切替.set(False)
        自身.表示を切替()

    def 設定値(自身):
        return None if 自身.既定使用.get() else list(自身.候補)
