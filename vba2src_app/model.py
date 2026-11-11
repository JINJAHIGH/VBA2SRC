"""解析と出力が共有する内部モデル。取得不能と空の設定を区別する。"""

from dataclasses import dataclass
from dataclasses import field as フィールド


# ── 1. プロジェクト内の取得情報 ──────────────────────────
@dataclass
class モジュール:
    """ソース本文、Attribute、分類根拠、保存ファイルの対応を保持する。"""

    名前: str
    種類: str
    抽出元ストリーム: str
    コード: str | None = None
    属性情報: list[str] = フィールド(default_factory=list)
    保存パス: str | None = None
    メタデータ: dict = フィールド(default_factory=dict)


@dataclass
class プロジェクト:
    """一つのVBAプロジェクトの取得結果と、項目別の取得状態を保持する。"""

    識別子: str
    抽出元識別子: str
    名前: str | None = None
    説明: str | None = None
    コードページ: int | None = None
    文字コード: str | None = None
    モジュール一覧: list[モジュール] = フィールド(default_factory=list)
    条件付きコンパイル引数: dict = フィールド(
        default_factory=lambda: {"status": "読取失敗", "raw": None, "items": []}
    )
    参照一覧: list[dict] = フィールド(default_factory=list)
    参照取得状態: str = "読取失敗"
    保護情報: dict = フィールド(default_factory=dict)
    フォーム一覧: list[dict] = フィールド(default_factory=list)
    診断一覧: list[str] = フィールド(default_factory=list)
    失敗一覧: list[str] = フィールド(default_factory=list)
    シート一覧: list[dict] = フィールド(default_factory=list)
    シート取得状態: str = "未対応"
    保存先接頭辞: str = ""
    モジュール取得状態: str = "読取失敗"
    保存モジュール数: int | None = None


# ── 2. 入力全体の状態と終了コード ──────────────────────────
@dataclass
class 抽出結果:
    """入力全体の状態と、相互に独立したプロジェクトを保持する。"""

    ファイル名: str
    形式: str = "不明"
    VBA有無: str = "判定不能"
    プロジェクト一覧: list[プロジェクト] = フィールド(default_factory=list)
    失敗一覧: list[str] = フィールド(default_factory=list)
    致命的失敗: bool = False
    名前定義情報: dict = フィールド(
        default_factory=lambda: {"status": "判定不能", "items": [], "diagnostics": [], "part": None}
    )

    @property
    def 終了コード(自身):
        """有用な部分回収の有無と失敗状態から、0・1・3を決定する。"""
        if 自身.致命的失敗:
            return 1
        if 自身.失敗一覧 or any(p.失敗一覧 for p in 自身.プロジェクト一覧):
            取得済み = any(
                p.名前 is not None
                or p.コードページ is not None
                or any(m.コード is not None for m in p.モジュール一覧)
                or p.保護情報.get("flags") is not None
                for p in 自身.プロジェクト一覧
            )
            # 名前定義だけの失敗では、確認済みのVBAなし判定も有用な回収結果とする。
            取得済み = 取得済み or 自身.VBA有無 == "なし" or bool(自身.名前定義情報["items"])
            return 3 if 取得済み else 1
        return 0

    @property
    def 状態(自身):
        """終了コードをレポート用の固定表示へ変換する。"""
        return {0: "成功", 1: "失敗", 3: "一部成功"}[自身.終了コード]
