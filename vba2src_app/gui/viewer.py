"""保存された設定・コードを用途別に表示する結果ビュー。"""

import json
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from ..console import 表示値
from .results import 保存パスを取得
from .style import 本文色, 状態のスタイル, 罫線色


def 明細文字列(データ):
    行一覧 = []
    for 名前, 値 in データ.items():
        if 値 is None:
            値 = "—"
        elif isinstance(値, (list, dict)):
            値 = json.dumps(値, ensure_ascii=False, indent=2)
            行一覧.append(f"{名前}\n{値}")
            continue
        elif isinstance(値, bool):
            値 = "はい" if 値 else "いいえ"
        行一覧.append(f"{名前}    {値}")
    return "\n".join(行一覧)


class 一覧画面(ttk.Frame):
    """選択行の詳細を下に表示する、読み取り専用の一覧。"""

    def __init__(自身, 親, 列設定, 詳細名, *, 等幅=False):
        super().__init__(親, padding=(14, 12))
        自身.columnconfigure(0, weight=1)
        自身.rowconfigure(1, weight=5, uniform="一覧と詳細")
        自身.rowconfigure(3, weight=5, uniform="一覧と詳細")

        自身.状態 = tk.StringVar()
        ttk.Label(自身, textvariable=自身.状態, style="Muted.TLabel").grid(row=0, sticky="w", pady=(0, 8))

        一覧枠 = ttk.Frame(自身)
        一覧枠.grid(row=1, sticky="nsew")
        一覧枠.columnconfigure(0, weight=1)
        一覧枠.rowconfigure(0, weight=1)
        列名 = [名前 for 名前, 幅 in 列設定]
        自身.一覧 = ttk.Treeview(一覧枠, columns=列名, show="headings", selectmode="browse", height=5)
        縦 = ttk.Scrollbar(一覧枠, orient="vertical", command=自身.一覧.yview)
        横 = ttk.Scrollbar(一覧枠, orient="horizontal", command=自身.一覧.xview)
        自身.一覧.configure(yscrollcommand=縦.set, xscrollcommand=横.set)
        自身.一覧.grid(row=0, column=0, sticky="nsew")
        縦.grid(row=0, column=1, sticky="ns")
        横.grid(row=1, column=0, sticky="ew")
        for 名前, 幅 in 列設定:
            自身.一覧.heading(名前, text=名前, anchor="w")
            自身.一覧.column(名前, width=幅, minwidth=幅, stretch=名前 == 列名[-1])

        ttk.Label(自身, text=詳細名, style="Section.TLabel").grid(row=2, sticky="w", pady=(12, 6))
        詳細枠 = ttk.Frame(自身)
        詳細枠.grid(row=3, sticky="nsew")
        詳細枠.columnconfigure(0, weight=1)
        詳細枠.rowconfigure(0, weight=1)
        自身.詳細 = tk.Text(
            詳細枠,
            wrap="none" if 等幅 else "word",
            height=4,
            padx=12,
            spacing1=2,
            spacing3=1,
            pady=9,
            relief="flat",
            background="white",
            foreground=本文色,
            highlightthickness=1,
            highlightbackground=罫線色,
            font="TkFixedFont" if 等幅 else "TkDefaultFont",
            state="disabled",
        )
        自身.詳細.grid(row=0, column=0, sticky="nsew")
        詳細縦 = ttk.Scrollbar(詳細枠, orient="vertical", command=自身.詳細.yview)
        詳細縦.grid(row=0, column=1, sticky="ns")
        自身.詳細.configure(yscrollcommand=詳細縦.set)
        if 等幅:
            詳細横 = ttk.Scrollbar(詳細枠, orient="horizontal", command=自身.詳細.xview)
            詳細横.grid(row=1, column=0, sticky="ew")
            自身.詳細.configure(xscrollcommand=詳細横.set)
        自身.一覧.bind("<<TreeviewSelect>>", 自身.選択を表示)
        自身.詳細を取得 = None

    def 詳細を表示(自身, 内容):
        自身.詳細.configure(state="normal")
        自身.詳細.delete("1.0", "end")
        自身.詳細.insert("1.0", 内容)
        自身.詳細.configure(state="disabled")
        自身.詳細.yview_moveto(0)
        自身.詳細.xview_moveto(0)

    def 行を設定(自身, 行一覧, 状態, 詳細を取得):
        自身.詳細を取得 = 詳細を取得
        選択項目 = 自身.一覧.get_children()
        自身.一覧.delete(*選択項目)
        for 番号, 項目 in enumerate(行一覧):
            値一覧 = tuple("—" if 値 is None or 値 == "" else 表示値(値) for 値 in 項目)
            自身.一覧.insert("", "end", iid=str(番号), values=値一覧)
        自身.状態.set(f"{状態}  ・  {len(行一覧)} 件")
        自身.詳細を表示("一覧から行を選択してください。" if 行一覧 else "表示できる項目はありません。")

    def 選択を表示(自身, イベント):
        選択 = 自身.一覧.selection()
        if 選択:
            try:
                内容 = 自身.詳細を取得(int(選択[0]))
            except (OSError, ValueError, RuntimeError) as 例外情報:
                内容 = f"保存ファイルを読み込めませんでした。\n{例外情報}"
            自身.詳細を表示(内容)


class 結果ビュー(ttk.Frame):
    """結果フォルダ全体の情報と、VBAプロジェクトごとの設定を表示する。"""

    def __init__(自身, 親):
        super().__init__(親)
        自身.結果 = None
        自身.タブ = ttk.Notebook(自身)
        自身.タブ.pack(fill="both", expand=True)

        自身.概要 = ttk.Frame(自身.タブ, padding=(14, 12))
        自身.タブ.add(自身.概要, text="概要")
        自身.指標 = ttk.Frame(自身.概要)
        自身.指標.pack(fill="x", pady=(0, 10))
        自身.詳細一覧 = ttk.Treeview(自身.概要, columns=("項目", "値"), show="headings", selectmode="none")
        自身.詳細一覧.heading("項目", text="項目", anchor="w")
        自身.詳細一覧.heading("値", text="値", anchor="w")
        自身.詳細一覧.column("項目", width=185, stretch=False)
        自身.詳細一覧.column("値", width=360)
        縦 = ttk.Scrollbar(自身.概要, orient="vertical", command=自身.詳細一覧.yview)
        自身.詳細一覧.configure(yscrollcommand=縦.set)
        縦.pack(side="right", fill="y")
        自身.詳細一覧.pack(fill="both", expand=True)

        一覧設定 = [
            ("VBAソース", [("モジュール名", 190), ("種類", 180), ("取得状態", 105)], "ソースコード", True),
            ("参照設定", [("表示名", 225), ("保存パス", 370)], "参照の詳細", False),
            ("条件付きコンパイル", [("定数名", 240), ("設定値", 140)], "条件付きコンパイルの設定", False),
            (
                "名前定義",
                [("名前", 160), ("適用範囲", 100), ("シート", 115), ("参照範囲", 135)],
                "名前定義の詳細",
                False,
            ),
            ("フォーム", [("ユーザーフォーム名", 250), ("取得状態", 200)], "フォームの詳細", False),
            ("保存ファイル", [("相対パス", 450)], "ファイルの内容", True),
        ]
        自身.画面一覧 = {}
        for 名前, 列設定, 詳細名, 等幅 in 一覧設定:
            画面 = 一覧画面(自身.タブ, 列設定, 詳細名, 等幅=等幅)
            自身.タブ.add(画面, text=名前)
            自身.画面一覧[名前] = 画面
        自身.タブ.enable_traversal()

    def 結果を表示(自身, 結果):
        自身.結果 = 結果
        名前定義 = 結果.名前定義.get("名前定義一覧", [])
        自身.画面一覧["名前定義"].行を設定(
            [
                (行.get("名前"), 行.get("適用範囲"), 行.get("適用先シート名"), 行.get("参照先の保存式"))
                for 行 in 名前定義
            ],
            結果.名前定義.get("取得状態", "情報なし"),
            lambda 番号: 明細文字列(名前定義[番号]),
        )
        自身.画面一覧["保存ファイル"].行を設定(
            [(パス.as_posix(),) for パス in 結果.保存ファイル],
            "結果フォルダ内の出力",
            lambda 番号: 自身.ファイルを読む(結果.保存ファイル[番号]),
        )
        自身.プロジェクトを表示(0)
        自身.タブ.select(自身.概要)

    def ファイルを読む(自身, 相対パス):
        if not 相対パス:
            return "保存ファイルはありません。"
        パス = 保存パスを取得(自身.結果.フォルダ, Path(相対パス))
        # サイズの大きなまとめファイルでも、先頭を表示して画面の応答性を保つ。
        上限 = 262_144
        with パス.open(encoding="utf-8-sig") as ストリーム:
            内容 = ストリーム.read(上限 + 1)
        return 内容 if len(内容) <= 上限 else 内容[:上限] + "\n\n… 続きは保存ファイルを参照してください。"

    def プロジェクトを表示(自身, 番号):
        結果 = 自身.結果
        プロジェクト = 結果.プロジェクト一覧[番号] if 結果.プロジェクト一覧 else None
        情報 = プロジェクト.情報 if プロジェクト else {}
        モジュール情報 = プロジェクト.モジュール if プロジェクト else {}
        参照情報 = プロジェクト.参照 if プロジェクト else {}
        コンパイル = プロジェクト.コンパイル if プロジェクト else {}
        保護 = プロジェクト.保護 if プロジェクト else {}

        モジュール = モジュール情報.get("モジュール一覧", [])
        参照 = 参照情報.get("参照設定一覧", [])
        定数 = コンパイル.get("設定一覧", [])
        フォーム = 情報.get("フォーム一覧", [])

        自身.画面一覧["VBAソース"].行を設定(
            [(行.get("名前"), 行.get("種類"), 行.get("ソース取得状態")) for 行 in モジュール],
            情報.get("モジュール取得状態", "情報なし"),
            lambda 番号: 自身.ファイルを読む(モジュール[番号].get("保存ファイル")),
        )
        自身.画面一覧["参照設定"].行を設定(
            [
                (行.get("表示名") or 行.get("名前"), 行.get("保存パス") or 行.get("外部プロジェクト相対パス"))
                for 行 in 参照
            ],
            参照情報.get("取得状態", "情報なし"),
            lambda 番号: 明細文字列(参照[番号]),
        )
        コンパイル画面 = 自身.画面一覧["条件付きコンパイル"]
        コンパイル画面.行を設定(
            [(行.get("名前"), 行.get("設定値")) for 行 in 定数],
            コンパイル.get("取得状態", "情報なし"),
            lambda 番号: 明細文字列(
                {
                    **定数[番号],
                    "元の値": コンパイル.get("元の値"),
                    "Unicode原文": コンパイル.get("Unicode原文"),
                }
            ),
        )
        if not 定数 and コンパイル.get("元の値"):
            コンパイル画面.詳細を表示(明細文字列(コンパイル))
        自身.画面一覧["フォーム"].行を設定(
            [(行.get("名前"), 行.get("取得状態")) for 行 in フォーム],
            "取得済みフォーム",
            lambda 番号: 自身.ファイルを読む(フォーム[番号].get("保存ファイル")),
        )
        自身.概要を表示(結果, プロジェクト, 情報, 保護, len(モジュール))

    def 概要を表示(自身, 結果, プロジェクト, 情報, 保護, モジュール数):
        for ウィジェット in 自身.指標.winfo_children():
            ウィジェット.destroy()
        入力 = 結果.入力
        指標一覧 = [
            ("取得状態", 入力.get("取得状態", "不明")),
            ("VBA", 入力.get("VBA有無", "不明")),
            ("モジュール", f"{モジュール数} 件"),
            ("名前定義", f"{len(結果.名前定義.get('名前定義一覧', []))} 件"),
        ]
        for 列, (名前, 値) in enumerate(指標一覧):
            自身.指標.columnconfigure(列, weight=1, uniform="指標")
            カード = ttk.Frame(自身.指標, style="Card.TFrame", padding=(12, 9))
            カード.grid(row=0, column=列, padx=(0, 6), sticky="ew")
            ttk.Label(カード, text=名前, style="Muted.TLabel").pack(anchor="w")
            状態スタイル = 状態のスタイル.get(値) if 名前 == "取得状態" else None
            スタイル = f"{状態スタイル}.Metric.TLabel" if 状態スタイル else "Metric.TLabel"
            ttk.Label(カード, text=値, style=スタイル).pack(anchor="w", pady=(2, 0))

        シート一覧 = [
            f"{行.get('表示名') or '—'} ({行.get('コードネーム') or '—'})"
            for 行 in 情報.get("シート一覧", [])
        ]
        診断 = 入力.get("失敗コード", []) + 情報.get("失敗コード", []) + 情報.get("診断コード", [])
        VBA保護状態 = 保護.get("VBA保護状態")
        概要 = [
            ("対象ファイル", 入力.get("元ファイル名")),
            ("ファイル形式", 入力.get("ファイル形式")),
            ("プロジェクト名", 情報.get("名前") or (プロジェクト.識別子 if プロジェクト else None)),
            ("説明", 情報.get("説明")),
            ("コードページ", 情報.get("コードページ")),
            ("VBA保護状態", VBA保護状態),
        ]
        if VBA保護状態 == "あり":
            概要.append(("VBAパスワード", 保護.get("一致したパスワード") or "不明"))
        概要.extend(
            (
                ("名前定義の取得状態", 結果.名前定義.get("取得状態")),
                ("シート / コードネーム", "  /  ".join(シート一覧)),
                ("失敗・診断コード", "  /  ".join(sorted(set(診断)))),
            )
        )
        自身.詳細一覧.delete(*自身.詳細一覧.get_children())
        for 名前, 値 in 概要:
            自身.詳細一覧.insert("", "end", values=(名前, "—" if 値 is None or 値 == "" else 表示値(値)))
