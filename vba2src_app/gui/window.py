"""抽出の入力、保存済み結果の閲覧、バックグラウンド処理を統括する。"""

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from queue import SimpleQueue
from threading import Thread
from tkinter import filedialog, font, ttk

from tkinterdnd2 import COPY, DND_FILES, REFUSE_DROP

from ..preferences import 設定を読む
from ..run import 抽出を実行
from .batch_window import 一括画面
from .help_dialog import 使い方画面
from .results import 結果を読む
from .settings_dialog import 設定画面
from .style import フォントを適用, フォント一覧を取得, 外観を設定, 強調色, 状態のスタイル
from .viewer import 結果ビュー

対応形式 = (
    ("Excel", "*.xlsx *.xlsm *.xls *.xla *.xlsb *.xltm *.xlam *.xltx"),
    ("Word", "*.docx *.docm *.dotx *.dotm *.doc *.dot"),
    ("Access", "*.accdb *.mdb"),
)


def パス入力値(入力欄):
    """Windowsの「パスのコピー」で付く外側の引用符を除去する。"""
    値 = 入力欄.get()
    if len(値) >= 2 and 値[0] == 値[-1] == '"':
        return 値[1:-1]
    return 値


class メイン画面:
    def __init__(自身, ウィンドウ, *, ドロップ有効):
        自身.ウィンドウ = ウィンドウ
        自身.ドロップ有効 = ドロップ有効
        自身.処理中 = False
        自身.ヘルプ = None
        自身.設定ウィンドウ = None
        自身.設定 = 設定を読む()
        自身.待ち行列 = SimpleQueue()
        自身.進捗待ち行列 = SimpleQueue()
        自身.対象 = tk.StringVar()
        自身.保存先 = tk.StringVar(value=自身.設定.保存先)
        自身.ログ設定 = tk.BooleanVar()
        自身.まとめ設定 = tk.BooleanVar()
        自身.結果パス = tk.StringVar()
        自身.状態 = tk.StringVar(value="対象ファイルを選択すると抽出できます。")
        自身.操作部品 = []

        ウィンドウ.title("VBA2SRC")
        ウィンドウ.geometry("1140x840")
        ウィンドウ.protocol("WM_DELETE_WINDOW", 自身.閉じる)
        自身.テーマ画像 = 外観を設定(ウィンドウ)
        自身.フォント一覧 = フォント一覧を取得(ウィンドウ)
        自身.画面標準字体 = font.nametofont("TkDefaultFont").actual("family")
        自身.フォントを反映()

        本体 = ttk.Frame(ウィンドウ, style="Workspace.TFrame", padding=22)
        本体.pack(fill="both", expand=True)
        本体.columnconfigure(1, weight=1)
        本体.rowconfigure(1, weight=1)

        ヘッダー = ttk.Frame(本体, style="Workspace.TFrame")
        ヘッダー.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        ttk.Label(ヘッダー, text="VBA2SRC", style="Brand.TLabel").pack(side="left")
        ttk.Button(ヘッダー, text="使い方", command=自身.ヘルプを開く).pack(side="right")
        設定ボタン = ttk.Button(ヘッダー, text="設定", command=自身.設定を開く)
        設定ボタン.pack(side="right", padx=(0, 8))
        自身.操作部品.append(設定ボタン)

        ナビ = ttk.Frame(本体, style="Card.TFrame", padding=(12, 16))
        ナビ.grid(row=1, column=0, sticky="ns", padx=(0, 16))
        自身.ナビボタン = {}
        自身.パネル = ttk.Frame(本体, style="Workspace.TFrame")
        自身.パネル.grid(row=1, column=1, sticky="nsew")
        自身.画面一覧 = {}
        for 名前, 構築 in (
            ("VBAを抽出", 自身.入力画面を作成),
            ("抽出結果を確認", 自身.結果画面を作成),
            ("比較", lambda: 一括画面(自身, False).画面),
            ("検索", lambda: 一括画面(自身, True).画面),
        ):
            ボタン = ttk.Button(
                ナビ, text=名前, width=15, command=lambda 画面名=名前: 自身.画面を切替(画面名)
            )
            ボタン.pack(fill="x", pady=(0, 8))
            自身.ナビボタン[名前] = ボタン
            自身.操作部品.append(ボタン)
            自身.画面一覧[名前] = 構築()

        フッター = ttk.Frame(本体, style="Workspace.TFrame")
        フッター.grid(row=2, columnspan=2, sticky="ew", pady=(14, 0))
        自身.状態ラベル = ttk.Label(
            フッター, textvariable=自身.状態, style="WorkspaceMuted.TLabel", wraplength=790
        )
        自身.状態ラベル.pack(side="left", fill="x", expand=True)
        自身.進捗 = ttk.Progressbar(フッター, mode="indeterminate", length=105)
        自身.画面を切替("VBAを抽出")

    def 見出しを作成(自身, 見出し, 説明):
        画面 = ttk.Frame(自身.パネル, style="Workspace.TFrame")
        ttk.Label(画面, text=見出し, style="Heading.TLabel").pack(anchor="w")
        ttk.Label(画面, text=説明, style="WorkspaceMuted.TLabel").pack(anchor="w", pady=(4, 16))
        return 画面

    def パス欄を作成(自身, 親, 変数, 選択処理, *, 入力スタイル="TEntry"):
        枠スタイル = "Drop.TFrame" if 入力スタイル == "Drop.TEntry" else "TFrame"
        枠 = ttk.Frame(親, style=枠スタイル)
        枠.pack(fill="x", pady=(5, 10))
        入力欄 = ttk.Entry(枠, textvariable=変数, style=入力スタイル)
        入力欄.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ボタン = ttk.Button(枠, text="参照", width=8, command=選択処理)
        ボタン.pack(side="right")
        自身.操作部品.extend((入力欄, ボタン))
        return 入力欄

    def 入力画面を作成(自身):
        画面 = 自身.見出しを作成("VBAを抽出", "Excel・Word・Access から、VBAソースと設定を保存します。")
        入力カード = ttk.Frame(画面, style="Card.TFrame", padding=18)
        入力カード.pack(fill="x", pady=(0, 12))
        対象枠 = ttk.Frame(入力カード, style="Drop.TFrame", padding=(12, 10))
        対象枠.pack(fill="x")
        ttk.Label(対象枠, text="対象ファイル", style="DropTitle.TLabel").pack(anchor="w")
        自身.対象欄 = 自身.パス欄を作成(対象枠, 自身.対象, 自身.対象を選択, 入力スタイル="Drop.TEntry")
        案内 = (
            "Officeファイルをここにドラッグ＆ドロップできます。"
            if 自身.ドロップ有効
            else "パスを入力するか、参照ボタンでOfficeファイルを選択してください。"
        )
        ttk.Label(対象枠, text=案内, style="DropHint.TLabel").pack(anchor="w")
        if 自身.ドロップ有効:
            アイコン = tk.Canvas(自身.対象欄, width=23, height=22, background="white", highlightthickness=0)
            アイコン.create_line(11.5, 2, 11.5, 13, fill=強調色, width=2, capstyle="round")
            アイコン.create_line(7.5, 9, 11.5, 13, 15.5, 9, fill=強調色, width=2, joinstyle="round")
            アイコン.create_line(
                3.5, 14, 3.5, 19, 19.5, 19, 19.5, 14, fill=強調色, width=2, joinstyle="round"
            )
            アイコン.place(relx=1, x=-10, rely=0.5, anchor="e")
            アイコン.bind("<Button-1>", lambda イベント: 自身.対象欄.focus_set())
            for 部品 in (対象枠, 自身.対象欄, アイコン):
                部品.drop_target_register(DND_FILES)
                部品.dnd_bind("<<Drop>>", 自身.対象をドロップ)

        ttk.Label(入力カード, text="出力先フォルダ", style="Section.TLabel").pack(anchor="w", pady=(14, 0))
        自身.パス欄を作成(入力カード, 自身.保存先, 自身.保存先を選択)
        ttk.Label(
            入力カード,
            text="空欄の場合は、対象ファイルと同じフォルダに出力します。",
            style="Muted.TLabel",
        ).pack(anchor="w")

        設定カード = ttk.Frame(画面, style="Card.TFrame", padding=18)
        設定カード.pack(fill="x")
        ttk.Label(設定カード, text="出力オプション", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        for 文言, 変数 in (
            ("モジュール別テキストファイル生成", 自身.まとめ設定),
            ("詳細ログを保存する", 自身.ログ設定),
        ):
            チェック = ttk.Checkbutton(設定カード, text=文言, variable=変数, style="Switch.TCheckbutton")
            チェック.pack(anchor="w", pady=4)
            自身.操作部品.append(チェック)

        実行ボタン = ttk.Button(
            画面, text="出力を実行", style="Accent.TButton", width=17, command=自身.抽出を開始
        )
        実行ボタン.pack(anchor="e", pady=(16, 0))
        自身.操作部品.append(実行ボタン)
        自身.対象欄.bind("<Return>", lambda イベント: 自身.抽出を開始())
        return 画面

    def 結果画面を作成(自身):
        画面 = 自身.見出しを作成(
            "抽出結果を確認", "抽出済みフォルダを開き、コードやVBAプロジェクトの設定を確認します。"
        )
        パスカード = ttk.Frame(画面, style="Card.TFrame", padding=(14, 10))
        パスカード.pack(fill="x", pady=(0, 12))
        ttk.Label(パスカード, text="結果フォルダ", style="Section.TLabel").pack(anchor="w")
        入力欄 = 自身.パス欄を作成(パスカード, 自身.結果パス, 自身.結果を選択)
        入力欄.bind("<Return>", lambda イベント: 自身.結果を開く())
        操作欄 = ttk.Frame(パスカード)
        操作欄.pack(fill="x")
        ttk.Label(操作欄, text="プロジェクト", style="Muted.TLabel").pack(side="left", padx=(0, 8))
        自身.プロジェクト欄 = ttk.Combobox(操作欄, state="disabled", width=23)
        自身.プロジェクト欄.pack(side="left", fill="x", expand=True)
        自身.プロジェクト欄.bind("<<ComboboxSelected>>", 自身.プロジェクトを切替)
        開く = ttk.Button(操作欄, text="読み込む", command=自身.結果を開く, style="Accent.TButton")
        開く.pack(side="right", padx=(8, 0))
        自身.フォルダボタン = ttk.Button(
            操作欄,
            text="フォルダを表示",
            command=自身.フォルダを表示,
            style="Important.TButton",
            state="disabled",
        )
        自身.フォルダボタン.pack(side="right", padx=(12, 0))
        自身.操作部品.append(開く)
        自身.結果待ち = ttk.Frame(画面, style="Card.TFrame", padding=22)
        自身.結果待ち.pack(fill="both", expand=True)
        ttk.Label(自身.結果待ち, text="表示する結果を選択", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            自身.結果待ち,
            text="保存済みの「入力情報.json」を含むフォルダを選択してください。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(6, 0))
        自身.ビュー = 結果ビュー(画面)
        return 画面

    def 画面を切替(自身, 名前):
        for 画面名, 画面 in 自身.画面一覧.items():
            ボタン = 自身.ナビボタン[画面名]
            ボタン.configure(style="Selected.TButton" if 名前 == 画面名 else "TButton")
            if 名前 == 画面名:
                画面.pack(fill="both", expand=True)
            else:
                画面.pack_forget()

    def 対象を選択(自身):
        全形式 = " ".join(拡張子 for 種類, 拡張子 in 対応形式)
        パス = filedialog.askopenfilename(
            parent=自身.ウィンドウ,
            title="対象のOfficeファイルを選択",
            initialdir=自身.設定.保存先 or None,
            filetypes=(("Officeファイル", 全形式), *対応形式, ("すべてのファイル", "*")),
        )
        if パス:
            自身.対象.set(パス)
            自身.状態を表示("抽出する準備ができました。")

    def 対象をドロップ(自身, イベント):
        if 自身.処理中:
            return REFUSE_DROP
        try:
            パス一覧 = 自身.ウィンドウ.tk.splitlist(イベント.data)
            if len(パス一覧) != 1:
                自身.状態を表示("対象ファイルを1件だけドロップしてください。", 状態="失敗")
                return REFUSE_DROP
            if not パス一覧[0] or not Path(パス一覧[0]).is_file():
                自身.状態を表示("ドロップされたファイルを確認できません。", 状態="失敗")
                return REFUSE_DROP
        except (OSError, ValueError, tk.TclError):
            自身.状態を表示("ドロップされたファイルを確認できません。", 状態="失敗")
            return REFUSE_DROP
        自身.対象.set(パス一覧[0])
        自身.状態を表示("抽出する準備ができました。")
        return COPY

    def 保存先を選択(自身):
        パス = filedialog.askdirectory(
            parent=自身.ウィンドウ,
            title="出力先フォルダを選択",
            initialdir=パス入力値(自身.保存先) or None,
        )
        if パス:
            自身.保存先.set(パス)

    def 結果を選択(自身):
        パス = filedialog.askdirectory(
            parent=自身.ウィンドウ, title="保存済みの結果フォルダを選択", initialdir=自身.設定.保存先 or None
        )
        if パス:
            自身.結果パス.set(パス)
            自身.結果を開く()

    def 状態を表示(自身, メッセージ, *, 状態=None):
        スタイル = 状態のスタイル.get(状態)
        自身.状態ラベル.configure(style=f"{スタイル}.Status.TLabel" if スタイル else "WorkspaceMuted.TLabel")
        自身.状態.set(メッセージ)

    def 抽出を開始(自身):
        if 自身.処理中:
            return
        対象 = パス入力値(自身.対象)
        if not 対象:
            自身.状態を表示("対象ファイルを指定してください。", 状態="失敗")
            自身.対象欄.focus_set()
            return
        保存先 = パス入力値(自身.保存先) or None
        詳細ログ, まとめ出力 = 自身.ログ設定.get(), 自身.まとめ設定.get()

        def 処理():
            実行 = 抽出を実行(対象, 保存先, 詳細ログ=詳細ログ, 種類別出力=まとめ出力)
            結果 = None
            if 実行.フォルダ and (実行.フォルダ / "入力情報.json").is_file():
                結果 = 結果を読む(実行.フォルダ)
            return 実行, 結果

        自身.処理を開始("VBAソースと設定を抽出しています…", 処理, 自身.抽出を完了)

    def 抽出を完了(自身, 応答):
        実行, 結果 = 応答
        if 結果:
            自身.結果を表示(結果)
        if 実行.メッセージ:
            案内 = f"  保存先：{実行.フォルダ}" if 実行.フォルダ else ""
            自身.状態を表示(実行.メッセージ + 案内, 状態="失敗")
        else:
            案内 = "  ・  概要の「事象の詳細」を確認してください。" if 実行.終了コード == 3 else ""
            if 実行.設定の注意:
                案内 += f"  ・  {実行.設定の注意}"
            自身.状態を表示(
                f"抽出結果：{実行.結果.状態}  ・  {実行.経過秒:.2f} 秒{案内}",
                状態="一部成功" if 実行.設定の注意 and 実行.終了コード == 0 else 実行.結果.状態,
            )

    def 結果を開く(自身):
        if 自身.処理中:
            return
        パス = パス入力値(自身.結果パス)
        if not パス:
            自身.状態を表示("結果フォルダを指定してください。", 状態="失敗")
            return
        自身.結果表示をクリア()
        自身.処理を開始("結果フォルダを読み込んでいます…", lambda: 結果を読む(パス), 自身.読込を完了)

    def 結果表示をクリア(自身):
        自身.ビュー.結果 = None
        自身.ビュー.pack_forget()
        自身.結果待ち.pack(fill="both", expand=True)
        自身.フォルダボタン.state(["disabled"])
        自身.プロジェクト欄.set("")
        自身.プロジェクト欄.configure(values=(), state="disabled")

    def 読込を完了(自身, 結果):
        自身.結果を表示(結果)
        自身.状態を表示(f"{結果.入力.get('元ファイル名', 結果.フォルダ.name)} の結果を表示しています。")

    def 結果を表示(自身, 結果):
        自身.ビュー.結果を表示(結果)
        自身.結果パス.set(str(結果.フォルダ))
        項目 = [
            f"{プロジェクト.情報.get('名前') or 'VBAProject'}  ({プロジェクト.識別子})"
            for プロジェクト in 結果.プロジェクト一覧
        ]
        自身.プロジェクト欄.configure(values=項目, state="readonly" if 項目 else "disabled")
        if 項目:
            自身.プロジェクト欄.current(0)
        else:
            自身.プロジェクト欄.set("VBAプロジェクトなし")
        自身.フォルダボタン.state(["!disabled"])
        自身.結果待ち.pack_forget()
        自身.ビュー.pack(fill="both", expand=True)
        自身.画面を切替("抽出結果を確認")

    def プロジェクトを切替(自身, イベント):
        番号 = 自身.プロジェクト欄.current()
        if 番号 >= 0:
            自身.ビュー.プロジェクトを表示(番号)

    def フォルダを表示(自身):
        フォルダ = 自身.ビュー.結果.フォルダ
        try:
            if os.name == "nt":
                os.startfile(フォルダ)
            else:
                コマンド = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([コマンド, str(フォルダ)])
        except OSError as 例外情報:
            自身.状態を表示(f"フォルダを表示できませんでした。{例外情報}", 状態="失敗")

    def 処理を開始(自身, メッセージ, 処理, 完了処理):
        自身.処理中 = True
        自身.状態を表示(メッセージ)
        for 部品 in 自身.操作部品:
            部品.state(["disabled"])
        自身.進捗.pack(side="right", padx=(12, 0))
        自身.進捗.start(16)

        def 実行():
            try:
                自身.待ち行列.put((処理(), None))
            except Exception as 例外情報:
                自身.待ち行列.put((None, 例外情報))

        Thread(target=実行, daemon=True).start()
        自身.ウィンドウ.after(60, 自身.処理を確認, 完了処理)

    def 進捗を通知(自身, メッセージ):
        自身.進捗待ち行列.put(メッセージ)

    def 処理を確認(自身, 完了処理):
        最新進捗 = None
        while not 自身.進捗待ち行列.empty():
            最新進捗 = 自身.進捗待ち行列.get_nowait()
        if 最新進捗:
            自身.状態を表示(最新進捗)
        if 自身.待ち行列.empty():
            自身.ウィンドウ.after(60, 自身.処理を確認, 完了処理)
            return
        応答, 例外情報 = 自身.待ち行列.get_nowait()
        自身.処理中 = False
        自身.進捗.stop()
        自身.進捗.pack_forget()
        for 部品 in 自身.操作部品:
            部品.state(["!disabled"])
        try:
            if 例外情報:
                raise 例外情報
            完了処理(応答)
        except (
            OSError,
            ValueError,
            TypeError,
            KeyError,
            AttributeError,
            RuntimeError,
            tk.TclError,
        ) as 例外情報:
            自身.状態を表示(f"抽出結果を表示できませんでした。{例外情報}", 状態="失敗")

    def 閉じる(自身):
        if 自身.処理中:
            自身.状態を表示("処理が完了してからウィンドウを閉じてください。")
        else:
            自身.ウィンドウ.destroy()

    def 設定を開く(自身):
        if 自身.設定ウィンドウ and 自身.設定ウィンドウ.winfo_exists():
            自身.設定ウィンドウ.lift()
            return
        自身.設定ウィンドウ = 設定画面(自身.ウィンドウ, 自身.設定, 自身.フォント一覧, 自身.設定を反映)

    def 設定を反映(自身, 設定):
        if パス入力値(自身.保存先) == 自身.設定.保存先:
            自身.保存先.set(設定.保存先)
        自身.設定 = 設定
        自身.フォントを反映()
        自身.状態を表示("設定を保存しました。")

    def フォントを反映(自身):
        フォントを適用(自身.ウィンドウ, 自身.設定, 自身.フォント一覧, 自身.画面標準字体)
        最小幅 = max(1000, 1000 + (自身.設定.画面サイズ - 10) * 76)
        幅 = min(最小幅, 自身.ウィンドウ.winfo_screenwidth() - 28)
        高さ = min(740, 自身.ウィンドウ.winfo_screenheight() - 56)
        自身.ウィンドウ.minsize(幅, 高さ)

    def ヘルプを開く(自身):
        if 自身.ヘルプ and 自身.ヘルプ.winfo_exists():
            自身.ヘルプ.lift()
            return
        自身.ヘルプ = 使い方画面(自身.ウィンドウ)
