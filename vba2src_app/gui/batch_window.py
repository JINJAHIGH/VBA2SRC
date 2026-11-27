"""一括抽出を使う比較画面と検索画面。"""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from ..batch import (
    HTMLを保存,
    パス一覧を読む,
    フォルダを検索,
    一括抽出,
    一致箇所を検索,
    作業フォルダを作成,
    差分行,
    情報の種類,
    比較対象一覧,
)
from ..console import 表示値

差分色 = {"追加": "#e7f5ed", "削除": "#fcebed", "変更": "#fff2d9"}


def 表を作成(親, 列一覧, 高さ):
    枠 = ttk.Frame(親)
    枠.columnconfigure(0, weight=1)
    枠.rowconfigure(0, weight=1)
    表 = ttk.Treeview(
        枠, columns=tuple(range(len(列一覧))), show="headings", height=高さ, selectmode="browse"
    )
    表.grid(row=0, column=0, sticky="nsew")
    for 番号, (名前, 幅) in enumerate(列一覧):
        表.heading(番号, text=名前, anchor="w")
        表.column(番号, width=幅, minwidth=幅, stretch=番号 == len(列一覧) - 1)
    for 方向, 関数, 行, 列, 吸着 in (
        ("vertical", 表.yview, 0, 1, "ns"),
        ("horizontal", 表.xview, 1, 0, "ew"),
    ):
        バー = ttk.Scrollbar(枠, orient=方向, command=関数)
        バー.grid(row=行, column=列, sticky=吸着)
        表.configure(**{"yscrollcommand" if 行 == 0 else "xscrollcommand": バー.set})
    for 状態, 色 in 差分色.items():
        表.tag_configure(状態, background=色)
    return 枠, 表


class 一括画面:
    def __init__(自身, アプリ, 検索):
        自身.アプリ = アプリ
        自身.検索 = 検索
        自身.結果 = []
        自身.ルート = None
        自身.画面 = アプリ.見出しを作成(
            "検索" if 検索 else "比較",
            "Officeファイルから情報を抽出して検索します。"
            if 検索
            else "Officeファイルごとの抽出結果を比較します。",
        )
        自身.タブ = ttk.Notebook(自身.画面)
        自身.タブ.pack(fill="both", expand=True)
        条件 = ttk.Frame(自身.タブ, padding=14)
        自身.結果欄 = ttk.Frame(自身.タブ, padding=14)
        自身.タブ.add(条件, text="条件")
        自身.タブ.add(自身.結果欄, text="結果", state="disabled")
        自身.入力方式 = tk.StringVar(value="フォルダ指定" if 検索 else "フルパス指定")
        自身.入力フォルダ = tk.StringVar()
        自身.出力先 = tk.StringVar()
        自身.出力名 = tk.StringVar()
        自身.再帰 = tk.BooleanVar()
        自身.キーワード = tk.StringVar()
        自身.まとめ = tk.BooleanVar()
        自身.ログ = tk.BooleanVar()
        自身.種類 = tk.StringVar(value="すべて")
        自身.案内 = tk.StringVar()
        if 検索:
            切替 = ttk.Combobox(
                条件, textvariable=自身.入力方式, values=("フォルダ指定", "フルパス指定"), state="readonly"
            )
            切替.pack(anchor="w", pady=(0, 8))
            切替.bind("<<ComboboxSelected>>", 自身.入力を切替)
        自身.入力領域 = ttk.Frame(条件)
        自身.入力領域.pack(fill="x")
        自身.フォルダ欄 = ttk.Frame(自身.入力領域)
        自身.パス欄 = ttk.Frame(自身.入力領域)
        ttk.Label(自身.パス欄, text="フルパスを1行に1件ずつ入力").pack(anchor="w")
        自身.ファイル一覧 = tk.Text(自身.パス欄, height=4, wrap="none", padx=7, pady=5)
        自身.ファイル一覧.pack(fill="x", pady=4)
        ttk.Button(自身.パス欄, text="ファイルを追加", command=自身.ファイルを選ぶ).pack(anchor="e")
        if 検索:
            自身.フォルダ入力(自身.フォルダ欄, "検索フォルダ", 自身.入力フォルダ)
            ttk.Checkbutton(自身.フォルダ欄, text="サブフォルダを含める", variable=自身.再帰).pack(anchor="w")
            除外欄 = ttk.Frame(自身.フォルダ欄)
            除外欄.pack(fill="x", pady=5)
            自身.除外 = []
            for 名前 in ("除外フォルダ名", "除外ファイル名"):
                枠 = ttk.Frame(除外欄)
                枠.pack(side="left", expand=True, fill="x", padx=(0, 8))
                ttk.Label(枠, text=名前 + "（1行に1件）").pack(anchor="w")
                入力 = tk.Text(枠, height=2, width=18, padx=5, pady=4)
                入力.pack(fill="x")
                自身.除外.append(入力)
        自身.入力を切替()
        自身.フォルダ入力(条件, "出力先フォルダ", 自身.出力先)
        if 検索:
            ttk.Label(条件, text="キーワード（大文字と小文字を区別しない部分一致）").pack(anchor="w")
            ttk.Entry(条件, textvariable=自身.キーワード).pack(fill="x", pady=(4, 6))
            種類欄 = ttk.Combobox(
                条件,
                textvariable=自身.種類,
                state="readonly",
                values=(
                    "すべて",
                    "VBAコード",
                    "名前定義一覧",
                    "参照設定",
                    "条件付きコンパイル引数",
                    "フォーム情報",
                    "保護情報",
                    "入力情報",
                    "プロジェクト情報",
                    "モジュール一覧",
                    "SHA256",
                    "実行ログ",
                ),
            )
            種類欄.pack(anchor="w")
        else:
            ttk.Label(条件, text="比較結果フォルダ名（空欄で日時を使用）").pack(anchor="w")
            ttk.Entry(条件, textvariable=自身.出力名).pack(fill="x", pady=(4, 6))
        オプション = ttk.Frame(条件)
        オプション.pack(fill="x", pady=6)
        ttk.Checkbutton(オプション, text="モジュール別テキストファイル生成", variable=自身.まとめ).pack(
            side="left"
        )
        ttk.Checkbutton(オプション, text="詳細ログを保存する", variable=自身.ログ).pack(side="left", padx=12)
        実行 = ttk.Button(
            条件, text="検索を実行" if 検索 else "比較を実行", style="Accent.TButton", command=自身.開始
        )
        実行.pack(anchor="e")
        アプリ.操作部品.append(実行)
        ttk.Label(自身.結果欄, textvariable=自身.案内, wraplength=750, style="Muted.TLabel").pack(
            anchor="w", pady=(0, 6)
        )
        自身.結果画面を作成()

    def フォルダ入力(自身, 親, 名前, 変数):
        ttk.Label(親, text=名前, style="Section.TLabel").pack(anchor="w")
        行 = ttk.Frame(親)
        行.pack(fill="x", pady=(4, 8))
        ttk.Entry(行, textvariable=変数).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(行, text="参照", command=lambda: 自身.フォルダを選ぶ(変数)).pack(side="right")

    def フォルダを選ぶ(自身, 変数):
        パス = filedialog.askdirectory(parent=自身.画面)
        if パス:
            変数.set(パス)

    def ファイルを選ぶ(自身):
        from .window import 対応形式

        パス一覧 = filedialog.askopenfilenames(parent=自身.画面, filetypes=対応形式)
        if パス一覧:
            既存 = 自身.ファイル一覧.get("1.0", "end-1c")
            自身.ファイル一覧.insert(
                "end", ("\n" if 既存 and not 既存.endswith("\n") else "") + "\n".join(パス一覧)
            )

    def 入力を切替(自身, イベント=None):
        for 枠 in (自身.フォルダ欄, 自身.パス欄):
            枠.pack_forget()
        選択 = 自身.フォルダ欄 if 自身.入力方式.get() == "フォルダ指定" else 自身.パス欄
        選択.pack(fill="x", pady=(0, 8))

    def 開始(自身):
        出力先 = 自身.出力先.get().strip().strip('"')
        キーワード = 自身.キーワード.get()
        if not 出力先 or (自身.検索 and not キーワード.strip()):
            自身.アプリ.状態を表示("出力先と、検索の場合はキーワードを入力してください。", 状態="失敗")
            return
        本文 = 自身.ファイル一覧.get("1.0", "end")
        フォルダ方式 = 自身.入力方式.get() == "フォルダ指定"
        入力フォルダ = 自身.入力フォルダ.get().strip().strip('"')
        再帰, 名前, まとめ, 種類 = (
            自身.再帰.get(),
            自身.出力名.get().strip(),
            自身.まとめ.get(),
            自身.種類.get(),
        )
        除外 = [入力.get("1.0", "end").splitlines() for 入力 in 自身.除外] if 自身.検索 else []
        検索 = 自身.検索
        ログ = 自身.ログ.get()

        def 処理():
            進捗 = 自身.アプリ.進捗を通知
            if フォルダ方式 and not 入力フォルダ:
                raise ValueError("検索フォルダを指定してください。")
            パス一覧 = (
                list(フォルダを検索(Path(入力フォルダ), 再帰, *除外))
                if フォルダ方式
                else パス一覧を読む(本文)
            )
            if len(パス一覧) < (1 if 検索 else 2):
                raise ValueError("対象ファイルが不足しています。比較は2件以上選択してください。")
            フォルダ = 作業フォルダを作成(Path(出力先).resolve(), 名前, "検索作業" if 検索 else "比較結果")
            結果 = 一括抽出(パス一覧, フォルダ, 進捗, まとめ=まとめ, 詳細ログ=ログ)
            一致 = 一致箇所を検索(結果, キーワード, 種類, 進捗) if 検索 else []
            return フォルダ, 結果, 一致

        自身.アプリ.処理を開始("対象ファイルを確認しています…", 処理, 自身.完了)

    def 結果画面を作成(自身):
        タブ = ttk.Notebook(自身.結果欄)
        タブ.pack(fill="both", expand=True)
        内容 = ttk.Frame(タブ, padding=8)
        タブ.add(内容, text="一致箇所" if 自身.検索 else "差分")
        枠, 自身.処理一覧 = 表を作成(タブ, [("ファイル", 260), ("状態", 90), ("理由", 240)], 8)
        タブ.add(枠, text="ファイル別の処理結果")
        if 自身.検索:
            枠, 自身.一致表 = 表を作成(
                内容, [("ファイル", 170), ("種類", 110), ("場所", 200), ("抜粋", 240)], 6
            )
            枠.pack(fill="both", expand=True)
            自身.一致表.bind("<<TreeviewSelect>>", 自身.一致詳細を表示)
            自身.詳細 = tk.Text(内容, height=5, wrap="word", state="disabled", padx=8, pady=6)
            自身.詳細.pack(fill="x", pady=(5, 0))
            return
        選択欄 = ttk.Frame(内容)
        選択欄.pack(fill="x", pady=(0, 5))
        自身.比較欄 = []
        for 名前 in ("基準", "比較対象"):
            ttk.Label(選択欄, text=名前).pack(side="left", padx=(0, 5))
            欄 = ttk.Combobox(選択欄, state="readonly", width=16)
            欄.pack(side="left", expand=True, fill="x", padx=(0, 8))
            欄.bind("<<ComboboxSelected>>", 自身.比較を表示)
            自身.比較欄.append(欄)
        保存 = ttk.Button(選択欄, text="HTML出力", style="Accent.TButton", command=自身.HTML設定を開く)
        保存.pack(side="right")
        自身.アプリ.操作部品.append(保存)
        枠, 自身.差分一覧 = 表を作成(内容, [("成果物", 360), ("状態", 90)], 3)
        枠.pack(fill="x")
        自身.差分一覧.bind("<<TreeviewSelect>>", 自身.差分を表示)
        枠, 自身.差分表 = 表を作成(
            内容, [("状態", 55), ("項目・行", 95), ("基準", 250), ("行", 45), ("比較対象", 250)], 5
        )
        枠.pack(fill="both", expand=True, pady=(5, 0))

        左右枠 = ttk.Frame(内容)
        左右枠.pack(fill="x", pady=(6, 0))
        自身.左右詳細 = []
        for 列, 名前 in enumerate(("基準の値・コード", "比較対象の値・コード")):
            左右枠.columnconfigure(列, weight=1, uniform="差分")
            ttk.Label(左右枠, text=名前).grid(row=0, column=列, sticky="w", padx=3)
            詳細 = tk.Text(左右枠, width=20, height=3, wrap="word", font="TkFixedFont", state="disabled")
            詳細.grid(row=1, column=列, sticky="ew", padx=3)
            自身.左右詳細.append(詳細)
        自身.差分表.bind("<<TreeviewSelect>>", 自身.比較行を表示)

    def 比較行を表示(自身, イベント=None):
        選択 = 自身.差分表.selection()
        行 = 自身.行一覧[int(選択[0])] if 選択 else None
        for 詳細, 列 in zip(自身.左右詳細, (2, 4), strict=True):
            詳細.configure(state="normal")
            詳細.delete("1.0", "end")
            if 行:
                詳細.insert("end", 行[列])
            詳細.configure(state="disabled")

    def 完了(自身, 応答):
        自身.ルート, 自身.結果, 自身.一致 = 応答
        自身.処理一覧.delete(*自身.処理一覧.get_children())
        for 項目 in 自身.結果:
            自身.処理一覧.insert("", "end", values=(str(項目.パス), 項目.状態, 項目.診断))
        自身.タブ.tab(自身.結果欄, state="normal")
        自身.タブ.select(自身.結果欄)
        if 自身.検索:
            自身.案内.set(f"{len(自身.一致)} 件一致  ・  {自身.ルート}")
            自身.一致表.delete(*自身.一致表.get_children())
            for 番号, (パス, 種類, 場所, 値) in enumerate(自身.一致):
                自身.一致表.insert(
                    "", "end", iid=str(番号), values=(Path(パス).name, 種類, 場所, 表示値(値[:240]))
                )
            自身.詳細を設定("")
        else:
            for 番号, 欄 in enumerate(自身.比較欄):
                欄.configure(values=[f"{位置}. {項目.パス.name}" for 位置, 項目 in enumerate(自身.結果, 1)])
                欄.current(番号)
            自身.比較を表示()
        失敗数 = sum(項目.状態 != "成功" for 項目 in 自身.結果)
        自身.アプリ.状態を表示(
            f"処理完了  {len(自身.結果)} / {len(自身.結果)} 件  ・  確認が必要なファイル {失敗数} 件"
        )

    def 詳細を設定(自身, 内容):
        自身.詳細.configure(state="normal")
        自身.詳細.delete("1.0", "end")
        自身.詳細.insert("end", 内容)
        自身.詳細.configure(state="disabled")

    def 一致詳細を表示(自身, イベント):
        選択 = 自身.一致表.selection()
        if 選択:
            自身.詳細を設定("\n".join(自身.一致[int(選択[0])]))

    def 比較を表示(自身, イベント=None):
        基準, 対象 = (自身.結果[欄.current()] for 欄 in 自身.比較欄)
        自身.差分一覧.delete(*自身.差分一覧.get_children())
        自身.差分表.delete(*自身.差分表.get_children())
        自身.行一覧 = []
        自身.比較行を表示()
        if "失敗" in (基準.状態, 対象.状態):
            自身.案内.set("比較できません。「ファイル別の処理結果」で理由を確認してください。")
            return
        自身.対象差分 = 比較対象一覧(基準.比較内容, 対象.比較内容)
        自身.案内.set(f"差分のある成果物 {len(自身.対象差分)} 件  ・  {自身.ルート}")
        for 番号, (パス, 状態) in enumerate(自身.対象差分):
            自身.差分一覧.insert("", "end", iid=str(番号), values=(パス.as_posix(), 状態), tags=(状態,))

    def 差分を表示(自身, イベント):
        選択 = 自身.差分一覧.selection()
        if not 選択:
            return
        パス = 自身.対象差分[int(選択[0])][0]
        基準, 対象 = (自身.結果[欄.current()].比較内容 for 欄 in 自身.比較欄)
        自身.差分表.delete(*自身.差分表.get_children())
        自身.行一覧 = list(差分行(基準.get(パス), 対象.get(パス)))
        自身.比較行を表示()
        for 番号, 行 in enumerate(自身.行一覧):
            自身.差分表.insert("", "end", iid=str(番号), values=tuple(表示値(値) for 値 in 行), tags=(行[0],))

    def HTML設定を開く(自身):
        基準番号, 対象番号 = (欄.current() for 欄 in 自身.比較欄)
        基準 = 自身.結果[基準番号]
        画面 = tk.Toplevel(自身.画面)
        画面.title("HTML出力範囲")
        画面.transient(自身.画面.winfo_toplevel())
        枠 = ttk.Frame(画面, padding=16)
        枠.pack(fill="both", expand=True)
        ttk.Label(枠, text="比較するOfficeファイル（複数選択可）").pack(anchor="w")
        選択欄 = tk.Listbox(枠, selectmode="extended", exportselection=False, height=5, width=48)
        選択欄.pack(fill="x", pady=(4, 8))
        対象一覧 = [(番号, 項目) for 番号, 項目 in enumerate(自身.結果) if 番号 != 基準番号]
        for 位置, (番号, 項目) in enumerate(対象一覧):
            選択欄.insert("end", f"{番号 + 1}. {項目.パス.name}")
            if 番号 == 対象番号:
                選択欄.selection_set(位置)
        if not 選択欄.curselection() and 対象一覧:
            選択欄.selection_set(0)
        成果物一覧 = sorted({パス for 項目 in 自身.結果 for パス in 項目.比較内容})
        種類 = ttk.Combobox(
            枠, values=["すべて", *sorted({情報の種類(パス) for パス in 成果物一覧})], state="readonly"
        )
        ttk.Label(枠, text="情報の種類").pack(anchor="w")
        種類.pack(fill="x", pady=(4, 8))
        種類.current(0)
        ttk.Label(枠, text="成果物の指定（空欄ですべて、複数選択可）").pack(anchor="w")
        成果物欄 = tk.Listbox(枠, selectmode="extended", exportselection=False, height=6, width=48)
        成果物欄.pack(fill="x", pady=(4, 8))
        for パス in 成果物一覧:
            成果物欄.insert("end", パス.as_posix())

        def 保存():
            対象 = [対象一覧[番号][1] for 番号 in 選択欄.curselection()]
            if not 対象:
                return
            パス = filedialog.asksaveasfilename(
                parent=画面,
                initialdir=自身.ルート,
                initialfile="比較結果.html",
                defaultextension=".html",
                filetypes=[("HTML", "*.html")],
            )
            if not パス:
                return
            分類 = 種類.get()
            成果物 = {成果物一覧[番号] for 番号 in 成果物欄.curselection()} or None
            画面.destroy()
            自身.アプリ.処理を開始(
                "HTMLを出力しています…",
                lambda: HTMLを保存(Path(パス), 基準, 対象, 分類, 成果物),
                lambda _: 自身.アプリ.状態を表示(f"HTMLを保存しました。{パス}"),
            )

        ttk.Button(枠, text="保存", style="Accent.TButton", command=保存).pack(anchor="e")
        画面.grab_set()
