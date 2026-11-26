"""Azureテーマの読み込みと、VBA2SRCで使う配色・書体・ボタン。"""

import math
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import font, ttk

背景色 = "#F4F6FA"
本文色 = "#17243A"
補足色 = "#52637A"
強調色 = "#0758AE"
罫線色 = "#D7DFEB"
エラー色 = "#A4262C"
成功色 = "#176341"
注意色 = "#825300"
状態のスタイル = {"成功": "Success", "失敗": "Error", "一部成功": "Partial"}


def フォント一覧を取得(ウィンドウ):
    return sorted({名前 for 名前 in font.families(ウィンドウ) if not 名前.startswith("@")}, key=str.casefold)


def フォントを適用(ウィンドウ, 設定, フォント一覧, 画面標準字体):
    使用可能 = {unicodedata.normalize("NFKC", 名前).casefold(): 名前 for 名前 in フォント一覧}
    画面候補 = (設定.画面フォント, "BIZ UDPゴシック", "BIZ UDPGothic")
    コード候補 = (設定.コードフォント, "MS ゴシック", "MS Gothic", "Consolas", "Menlo")
    for 候補一覧, 基準フォント, 設定サイズ, 既定字体 in (
        (画面候補, "TkDefaultFont", 設定.画面サイズ, 画面標準字体),
        (コード候補, "TkFixedFont", 設定.コードサイズ, None),
    ):
        for 名前 in 候補一覧:
            選択 = 使用可能.get(unicodedata.normalize("NFKC", 名前).casefold())
            if 選択:
                font.nametofont(基準フォント).configure(family=選択, size=設定サイズ)
                break
        else:
            フォント = font.nametofont(基準フォント)
            フォント.configure(family=既定字体 or フォント.actual("family"), size=設定サイズ)

    字体 = font.nametofont("TkDefaultFont").actual("family")
    サイズ = 設定.画面サイズ
    スタイル = ttk.Style(ウィンドウ)
    for 名前, 大きさ, 太さ in (
        (".", サイズ, "normal"),
        ("Brand.TLabel", サイズ + 24, "bold"),
        ("Heading.TLabel", サイズ + 13, "bold"),
        ("Section.TLabel", サイズ + 1, "bold"),
        ("Metric.TLabel", サイズ + 7, "bold"),
        ("Treeview", サイズ, "normal"),
        ("Treeview.Heading", サイズ, "bold"),
        ("TButton", サイズ, "normal"),
        ("Important.TButton", サイズ, "bold"),
        ("Add.TButton", サイズ, "bold"),
        ("Update.TButton", サイズ, "bold"),
        ("Delete.TButton", サイズ, "bold"),
        ("Reset.TButton", サイズ, "bold"),
        ("Status.TLabel", サイズ, "bold"),
        ("DropTitle.TLabel", サイズ + 1, "bold"),
        ("Accent.TButton", サイズ, "bold"),
        ("Selected.TButton", サイズ, "bold"),
        ("TNotebook.Tab", サイズ, "bold"),
    ):
        スタイル.configure(名前, font=(字体, 大きさ, 太さ))
    スタイル.configure("Treeview", rowheight=サイズ + 18)


def 丸角の画像を作成(ウィンドウ, 塗り色, 枠色, *, 余白=0):
    """角を固定したまま伸ばせる小さなボタン背景を作る。"""
    大きさ, 半径 = 21, 6
    画像 = tk.PhotoImage(master=ウィンドウ, width=大きさ + 余白 * 2, height=大きさ)
    for 行 in range(大きさ):
        角まで = max(半径 - 行, 行 - (大きさ - 1 - 半径), 0)
        左端 = 半径 - int(math.sqrt(半径**2 - 角まで**2))
        右端 = 大きさ - 左端
        画像.put(枠色, to=(余白 + 左端, 行, 余白 + 右端, 行 + 1))
        if 0 < 行 < 大きさ - 1 and 右端 - 左端 > 2:
            画像.put(塗り色, to=(余白 + 左端 + 1, 行, 余白 + 右端 - 1, 行 + 1))
    return 画像


def 外観を設定(ウィンドウ):
    テーマ = Path(__file__).with_name("assets") / "azure" / "azure.tcl"
    ウィンドウ.tk.call("source", str(テーマ))
    ウィンドウ.tk.call("set_theme", "light")
    # AzureのオプションDBより、各ttkスタイルのフォント指定を優先する。
    ウィンドウ.option_clear()
    スタイル = ttk.Style(ウィンドウ)
    スタイル.configure(".", foreground=本文色)
    スタイル.configure("Workspace.TFrame", background=背景色)
    スタイル.configure("Brand.TLabel", foreground=本文色, background=背景色)
    スタイル.configure("Heading.TLabel", foreground=本文色, background=背景色)
    スタイル.configure("Section.TLabel", foreground=本文色)
    スタイル.configure("Muted.TLabel", foreground=補足色)
    スタイル.configure("WorkspaceMuted.TLabel", foreground=補足色, background=背景色)
    スタイル.configure("Error.TLabel", foreground=エラー色, background=背景色)
    for 名前, 色 in (("Success", 成功色), ("Error", エラー色), ("Partial", 注意色)):
        スタイル.configure(f"{名前}.Status.TLabel", foreground=色, background=背景色)
        スタイル.configure(f"{名前}.Metric.TLabel", foreground=色)
    ドロップ背景 = "#EDF4FD"
    スタイル.configure("Drop.TFrame", background=ドロップ背景)
    スタイル.configure("DropTitle.TLabel", foreground=本文色, background=ドロップ背景)
    スタイル.configure("DropHint.TLabel", foreground=補足色, background=ドロップ背景)
    スタイル.configure("Drop.TEntry", padding=(10, 9))
    スタイル.configure("Metric.TLabel", foreground=強調色)
    スタイル.configure("Treeview.Heading", padding=(5, 4))
    スタイル.map("Treeview", background=[("selected", "#DCEBFC")], foreground=[("selected", 本文色)])

    # Azureのコントロールを使い、ボタンとタブの操作状態はアプリ内で揃える。
    画像一覧 = []
    ボタン設定 = (
        ("TButton", "RegularAction", 本文色, "#FFFFFF", 罫線色, "#EFF4FA", "#E4EDF9"),
        ("Accent.TButton", "PrimaryAction", "#FFFFFF", "#0965C7", "#0965C7", "#0758AE", "#064B96"),
        ("Selected.TButton", "SelectedAction", 強調色, "#E8F1FD", "#C8DCF5", "#DCEBFC", "#D2E4FA"),
        ("Add.TButton", "AddAction", 成功色, "#E7F5ED", "#B8DDC8", "#D6EFDF", "#C1E6CF"),
        ("Update.TButton", "UpdateAction", 強調色, "#E8F1FD", "#C8DCF5", "#DCEBFC", "#D2E4FA"),
        ("Delete.TButton", "DeleteAction", エラー色, "#FCEBED", "#EFC7CC", "#F9DCE0", "#F1CFD4"),
        ("Reset.TButton", "ResetAction", 注意色, "#FFF2D9", "#ECD3A0", "#FAE8BF", "#F0DCA9"),
    )
    for 名前, 要素名, 文字色, 通常色, 枠色, ホバー色, 押下色 in ボタン設定:
        通常, ホバー, 押下, 無効, フォーカス = [
            丸角の画像を作成(ウィンドウ, 塗り色, 輪郭色)
            for 塗り色, 輪郭色 in (
                (通常色, 枠色),
                (ホバー色, 枠色),
                (押下色, 枠色),
                ("#F0F2F5", "#DFE4EC"),
                (通常色, "#6EA6E6"),
            )
        ]
        画像一覧.extend((通常, ホバー, 押下, 無効, フォーカス))
        スタイル.element_create(
            要素名,
            "image",
            通常,
            ("disabled", 無効),
            ("pressed", 押下),
            ("focus", フォーカス),
            ("active", ホバー),
            border=7,
            sticky="nsew",
        )
        スタイル.layout(
            名前,
            [
                (
                    要素名,
                    {
                        "sticky": "nsew",
                        "children": [
                            (
                                "Button.padding",
                                {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]},
                            )
                        ],
                    },
                )
            ],
        )
        スタイル.configure(名前, padding=(14, 5), anchor="center")
        スタイル.map(名前, foreground=[("disabled", "#5D6A7C"), ("!disabled", 文字色)])

    タブ通常, タブ選択, タブホバー = [
        丸角の画像を作成(ウィンドウ, 色, 枠, 余白=3)
        for 色, 枠 in (("#F4F6FA", "#E0E6EF"), ("#E8F1FD", "#C8DCF5"), ("#EAF0F8", "#D7DFEB"))
    ]
    画像一覧.extend((タブ通常, タブ選択, タブホバー))
    スタイル.element_create(
        "ResultTab",
        "image",
        タブ通常,
        ("selected", タブ選択),
        ("active", タブホバー),
        border=(10, 7, 10, 7),
        sticky="nsew",
    )
    スタイル.layout(
        "TNotebook.Tab",
        [
            (
                "ResultTab",
                {
                    "sticky": "nsew",
                    "children": [
                        (
                            "Notebook.padding",
                            {"sticky": "nsew", "children": [("Notebook.label", {"sticky": "nsew"})]},
                        )
                    ],
                },
            ),
        ],
    )
    スタイル.configure("TNotebook", tabmargins=(1, 0, 1, 8))
    スタイル.configure("TNotebook.Tab", padding=(14, 6))
    スタイル.map("TNotebook.Tab", foreground=[("selected", 強調色)])
    return 画像一覧
