"""通常の抽出結果を使った比較、検索、HTMLレポート。"""

import json
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from html import escape
from pathlib import Path

from .output import テキストを保存
from .paths import 安全な保存名
from .run import 抽出を実行

対応拡張子 = frozenset(
    ".xlsx .xlsm .xls .xla .xlsb .xltm .xlam .xltx .docx .docm .dotx .dotm .doc .dot .accdb .mdb".split()
)
状態名 = {"insert": "追加", "delete": "削除", "replace": "変更", "equal": "同一"}


@dataclass
class 抽出項目:
    パス: Path
    フォルダ: Path | None
    状態: str
    診断: str
    比較内容: dict
    検索内容: dict


def パス一覧を読む(本文):
    パス一覧 = [Path(行.strip().strip('"')) for 行 in 本文.splitlines() if 行.strip()]
    if not パス一覧 or any(not パス.is_absolute() for パス in パス一覧):
        raise ValueError("Officeファイルのフルパスを1行に1件ずつ入力してください。")
    return list(dict.fromkeys(パス一覧))


def フォルダを検索(親, 再帰, 除外フォルダ, 除外ファイル):
    if not 親.is_dir():
        raise ValueError("検索するフォルダを指定してください。")
    除外フォルダ = {名前.strip().casefold() for 名前 in 除外フォルダ}
    除外ファイル = {名前.strip().casefold() for 名前 in 除外ファイル}
    for 現在地, フォルダ一覧, ファイル一覧 in 親.walk():
        フォルダ一覧[:] = sorted(
            名前 for 名前 in フォルダ一覧 if 再帰 and 名前.casefold() not in 除外フォルダ
        )
        yield from (
            現在地 / 名前
            for 名前 in sorted(ファイル一覧)
            if 名前.casefold() not in 除外ファイル and Path(名前).suffix.lower() in 対応拡張子
        )


def 作業フォルダを作成(親, 名前, モード):
    if 名前 and (名前 != 安全な保存名(名前) or 名前 in (".", "..")):
        raise ValueError("フォルダ名にはWindowsで使用できる名前を入力してください。")
    名前 = 名前 or f"{モード}_{datetime.now():%Y%m%d_%H%M%S}"
    親.mkdir(parents=True, exist_ok=True)
    番号 = 0
    while True:
        フォルダ = 親 / (f"{名前}_{番号:03d}" if 番号 else 名前)
        try:
            フォルダ.mkdir()
            return フォルダ
        except FileExistsError:
            番号 += 1


def JSON項目を列挙(値, 場所=""):
    if isinstance(値, dict):
        for 名前, 内容 in 値.items():
            yield from JSON項目を列挙(内容, f"{場所} / {名前}" if 場所 else 名前)
    elif isinstance(値, list) and 値:
        識別子の件数 = {}
        for 番号, 内容 in enumerate(値, 1):
            if isinstance(内容, dict):
                識別子 = " / ".join(
                    str(内容[キー])
                    for キー in ("名前", "識別子", "スコープ", "シート名")
                    if 内容.get(キー) is not None
                )
            else:
                識別子 = ""
            識別子 = 識別子 or str(番号)
            件数 = 識別子の件数.get(識別子, 0) + 1
            識別子の件数[識別子] = 件数
            yield from JSON項目を列挙(内容, f"{場所} / {識別子}" + (f" ({件数})" if 件数 > 1 else ""))
    else:
        yield 場所, 値 if isinstance(値, str) else json.dumps(値, ensure_ascii=False)


def 出力内容を読む(フォルダ):
    比較, 検索 = {}, {}
    if フォルダ is None:
        return 比較, 検索
    for パス in sorted(フォルダ.rglob("*")):
        if not パス.is_file() or パス.suffix.lower() not in {".bas", ".cls", ".vba", ".txt", ".json"}:
            continue
        相対 = パス.relative_to(フォルダ)
        本文 = パス.read_text(encoding="utf-8-sig")
        if パス.suffix.lower() == ".json":
            項目 = dict(JSON項目を列挙(json.loads(本文)))
            検索[相対] = 項目
            比較[相対] = {
                場所: 値
                for 場所, 値 in 項目.items()
                if 場所 not in {"元ファイル名", "保存用ファイル名"}
                and 場所.rsplit(" / ", 1)[-1] != "保存先接頭辞"
            }
        else:
            検索[相対] = 比較[相対] = 本文
    return 比較, 検索


def 一括抽出(パス一覧, フォルダ, 進捗, *, まとめ=False, 詳細ログ=False):
    結果一覧 = []
    for 番号, パス in enumerate(パス一覧, 1):
        進捗(f"抽出中  {番号 - 1} / {len(パス一覧)} 件  ・  {パス.name}")
        実行 = 抽出を実行(パス, フォルダ, 種類別出力=まとめ, 詳細ログ=詳細ログ)
        診断 = 実行.メッセージ
        if 実行.結果:
            コード = set(実行.結果.失敗一覧)
            for プロジェクト in 実行.結果.プロジェクト一覧:
                コード.update(プロジェクト.失敗一覧)
            診断 = 診断 or " / ".join(sorted(コード))
        状態 = 実行.結果.状態 if 実行.結果 and not 実行.メッセージ else "失敗"
        try:
            比較, 検索 = 出力内容を読む(実行.フォルダ)
        except (OSError, ValueError) as 例外:
            比較, 検索 = {}, {}
            状態, 診断 = "失敗", f"出力を読み込めませんでした。{例外}"
        結果一覧.append(抽出項目(パス, 実行.フォルダ, 状態, 診断, 比較, 検索))
    return 結果一覧


def 比較対象一覧(基準, 対象):
    return [
        (パス, "変更" if パス in 基準 and パス in 対象 else "追加" if パス in 対象 else "削除")
        for パス in sorted(基準.keys() | 対象.keys())
        if 基準.get(パス) != 対象.get(パス)
    ]


def 差分行(左, 右):
    if isinstance(左, dict) or isinstance(右, dict):
        左, 右 = 左 or {}, 右 or {}
        for キー in sorted(左.keys() | 右.keys()):
            if 左.get(キー) != 右.get(キー):
                状態 = "変更" if キー in 左 and キー in 右 else "追加" if キー in 右 else "削除"
                yield 状態, キー, 左.get(キー, ""), "", 右.get(キー, "")
        return
    左, 右 = (左 or "").splitlines(), (右 or "").splitlines()
    for 種類, 開始左, 終了左, 開始右, 終了右 in SequenceMatcher(None, 左, 右).get_opcodes():
        for 番号 in range(max(終了左 - 開始左, 終了右 - 開始右)):
            左番号, 右番号 = 開始左 + 番号, 開始右 + 番号
            yield (
                状態名[種類],
                左番号 + 1 if 左番号 < 終了左 else "",
                左[左番号] if 左番号 < 終了左 else "",
                右番号 + 1 if 右番号 < 終了右 else "",
                右[右番号] if 右番号 < 終了右 else "",
            )


def 情報の種類(パス):
    if "VBAソース" in パス.parts or "VBAまとめ" in パス.parts:
        return "フォーム情報" if パス.suffix == ".json" else "VBAコード"
    return パス.stem


def 一致箇所を検索(結果, キーワード, 種類, 進捗):
    一致 = []
    検索語 = キーワード.casefold()
    for 番号, 項目 in enumerate(結果, 1):
        進捗(f"検索中  {番号 - 1} / {len(結果)} 件  ・  {項目.パス.name}")
        for パス, 内容 in 項目.検索内容.items():
            分類 = 情報の種類(パス)
            if 種類 != "すべて" and 分類 != 種類:
                continue
            行一覧 = 内容.items() if isinstance(内容, dict) else enumerate(内容.splitlines(), 1)
            for 場所, 値 in 行一覧:
                検索値 = f"{場所} {値}" if isinstance(内容, dict) else 値
                if 検索語 in 検索値.casefold():
                    一致.append((str(項目.パス), 分類, f"{パス.as_posix()} / {場所}", 値))
    return 一致


def HTMLを保存(パス, 基準, 対象一覧, 種類, 成果物一覧=None):
    節 = []
    for 対象 in 対象一覧:
        左, 右 = 基準.比較内容, 対象.比較内容
        節.append(f"<section><h2>{escape(str(基準.パス))} ⇄ {escape(str(対象.パス))}</h2>")
        for 項目 in (基準, 対象):
            節.append(f"<p>{escape(項目.パス.name)}　{escape(項目.状態)}　{escape(項目.診断)}</p>")
        if "失敗" in (基準.状態, 対象.状態):
            節.append("<p>抽出に失敗したファイルとの比較は行えません。</p></section>")
            continue
        件数 = 0
        for 相対, 状態 in 比較対象一覧(左, 右):
            if 種類 != "すべて" and 情報の種類(相対) != 種類:
                continue
            if 成果物一覧 is not None and 相対 not in 成果物一覧:
                continue
            件数 += 1
            節.append(f"<details open><summary>{escape(状態)}　{escape(相対.as_posix())}</summary>")
            節.append(
                "<table><thead><tr><th>差分</th><th>位置</th><th>基準</th><th>位置</th><th>比較対象</th></tr></thead><tbody>"
            )
            for 行 in 差分行(左.get(相対), 右.get(相対)):
                色 = {"追加": "add", "削除": "del", "変更": "edit"}.get(行[0], "")
                節.append(
                    f'<tr class="{色}">' + "".join(f"<td>{escape(str(値))}</td>" for 値 in 行) + "</tr>"
                )
            節.append("</tbody></table></details>")
        if not 件数:
            節.append("<p>選択された範囲に差分はありません。</p>")
        節.append("</section>")
    テキストを保存(
        パス,
        """<!doctype html><html lang="ja"><meta charset="utf-8">
<title>VBA2SRC 比較結果</title><style>
body{font:16px system-ui,sans-serif;color:#17243a;background:#f4f6fa;margin:2rem}
section,details{background:white;border:1px solid #d7dfeb;border-radius:8px;padding:1rem;margin:1rem 0}
h2{font-size:1rem;overflow-wrap:anywhere}summary{font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:1rem;table-layout:fixed}
td,th{text-align:left;vertical-align:top;padding:.3rem;border-bottom:1px solid #d7dfeb;overflow-wrap:anywhere}
td{white-space:pre-wrap;font-family:monospace}th:nth-child(1),th:nth-child(2),th:nth-child(4){width:8%}
.add{background:#e7f5ed}.del{background:#fcebed}.edit{background:#fff2d9}
</style><h1>VBA2SRC 比較結果</h1>"""
        + "\n".join(節)
        + "</html>\n",
    )
