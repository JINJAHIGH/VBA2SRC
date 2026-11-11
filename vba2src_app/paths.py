"""保存先の規則と、Windows互換の名前・衝突回避を管理する。"""

import re
import sys
from datetime import datetime
from pathlib import Path


# ── 1. 保存先の決定 ──────────────────────────
def ツールフォルダを取得():
    """Python起動ファイルまたは実EXEの格納先を返す。一時展開先は使わない。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def 既定親フォルダを取得(設定値):
    """対話入力の初期値を設定から決め、無効な設定はツール格納先へ戻す。"""
    if 設定値:
        パス = Path(設定値).expanduser()
        if not パス.is_absolute():
            パス = ツールフォルダを取得() / パス
        if パス.is_dir():
            return パス.resolve()
    return ツールフォルダを取得()


def 保存先の親を決定(対象, 指定先=None):
    """明示された保存先を優先し、省略時は対象ファイルの親を返す。"""
    # 存在確認・作成は後段で行う。共有先の障害時に別の場所へ退避しない。
    return Path(指定先).resolve() if 指定先 is not None else Path(対象).resolve().parent


# ── 2. 保存名と衝突回避 ──────────────────────────
def 安全な保存名(値, 上限=100):
    """Windowsの禁止文字・予約名・長さを調整して単一のパス要素へ変換する。"""
    名前 = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', "_", 値).rstrip(" .")
    if not 名前:
        名前 = "_"
    if 名前.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL"} | {
        f"{種類}{i}" for 種類 in ("COM", "LPT") for i in "123456789¹²³"
    }:
        名前 = "_" + 名前
    # 補助平面の文字もUTF-16コード単位で数え、Windowsの要素長を制限する。
    while len(名前.encode("utf-16-le")) // 2 > 上限:
        名前 = 名前[:-1]
    return 名前.rstrip(" .") or "_"


class 保存名割当器:
    """同じ出力先内の名前をWindows互換の比較規則で予約する。"""

    def __init__(自身):
        自身.使用済み名 = set()

    def 割当(自身, 名前, 拡張子=""):
        """大小文字を区別しない衝突を連番で解消する。呼び出し側で割当順を固定する。"""
        基準名 = 安全な保存名(名前)
        候補 = 基準名 + 拡張子
        連番 = 0
        while 候補.casefold() in 自身.使用済み名:
            連番 += 1
            候補 = f"{基準名}_{連番:03d}{拡張子}"
        自身.使用済み名.add(候補.casefold())
        return 候補


def 結果フォルダを作成(親, ファイル名, 現在時刻=None):
    """ローカル時刻の名前を原子的に確保し、既存結果を上書きしない。"""
    親 = Path(親)
    親.mkdir(parents=True, exist_ok=True)
    基準名 = f"{安全な保存名(ファイル名)}_{(現在時刻 or datetime.now()):%Y%m%d_%H%M%S}"
    連番 = 0
    while True:
        パス = 親 / (基準名 if 連番 == 0 else f"{基準名}_{連番:03d}")
        try:
            パス.mkdir()  # mkdirの成功を予約とし、競合した既存フォルダへ混在させない。
            return パス
        except FileExistsError:
            連番 += 1
