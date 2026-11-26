"""出力済みフォルダの新旧レイアウトを、同じ表示モデルとして読み込む。"""

import json
from dataclasses import dataclass
from pathlib import Path

from ..json_output import スキーマバージョン


@dataclass
class プロジェクト表示:
    識別子: str
    情報: dict
    モジュール: dict
    参照: dict
    コンパイル: dict
    保護: dict


@dataclass
class 結果表示:
    フォルダ: Path
    入力: dict
    名前定義: dict
    プロジェクト一覧: list[プロジェクト表示]
    保存ファイル: list[Path]


def 保存パスを取得(フォルダ, 相対パス):
    """JSONに記録されたパスを、選択した結果フォルダの範囲内で解決する。"""
    パス = (フォルダ / 相対パス).resolve()
    if not パス.is_relative_to(フォルダ):
        raise ValueError("結果フォルダの外を指す保存パスがあります。")
    return パス


def JSONを読む(フォルダ, 相対フォルダ, ファイル名, *, 必須=False):
    """直下の配置を優先し、旧版では設定情報フォルダを参照する。"""
    for 相対パス in (相対フォルダ / ファイル名, 相対フォルダ / "設定情報" / ファイル名):
        パス = 保存パスを取得(フォルダ, 相対パス)
        try:
            with パス.open(encoding="utf-8-sig") as ストリーム:
                データ = json.load(ストリーム)
        except FileNotFoundError:
            continue
        if not isinstance(データ, dict) or データ.get("スキーマバージョン") != スキーマバージョン:
            raise ValueError(f"{ファイル名}の形式に対応していません。")
        return データ
    if 必須:
        raise ValueError(f"{ファイル名}が見つかりません。抽出結果のフォルダを指定してください。")
    return {}


def 結果を読む(フォルダ):
    """入力情報を起点に、全プロジェクトと関連する出力を取得する。"""
    フォルダ = Path(フォルダ).resolve()
    入力 = JSONを読む(フォルダ, Path(), "入力情報.json", 必須=True)
    名前定義 = JSONを読む(フォルダ, Path(), "名前定義一覧.json")
    プロジェクト一覧 = []
    for 項目 in 入力.get("プロジェクト一覧", []):
        相対フォルダ = Path(項目["保存先接頭辞"])
        プロジェクト一覧.append(
            プロジェクト表示(
                項目["識別子"],
                JSONを読む(フォルダ, 相対フォルダ, "プロジェクト情報.json"),
                JSONを読む(フォルダ, 相対フォルダ, "モジュール一覧.json"),
                JSONを読む(フォルダ, 相対フォルダ, "参照設定.json"),
                JSONを読む(フォルダ, 相対フォルダ, "条件付きコンパイル引数.json"),
                JSONを読む(フォルダ, 相対フォルダ, "保護情報.json"),
            )
        )
    保存ファイル = sorted(
        パス.relative_to(フォルダ)
        for パス in フォルダ.rglob("*")
        if パス.is_file() and パス.suffix.lower() in {".txt", ".json", ".bas", ".cls", ".vba"}
    )
    return 結果表示(フォルダ, 入力, 名前定義, プロジェクト一覧, 保存ファイル)
