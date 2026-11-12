# ── 1. Windows配布用の依存収集 ───────────────────────────
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, copy_metadata

if sys.platform != "win32":
    raise SystemExit("Build the Windows executable on Windows.")

同梱バイナリ = []
# Python本体とabi3t互換DLLを同梱し、開発PCのPATHに依存させない。
for ディレクトリ in (Path(sys.base_prefix), Path(sys.base_prefix) / "abi3t-compat"):
    if ディレクトリ.is_dir():
        for DLLファイル in sorted(ディレクトリ.glob("python3*.dll")):
            if not any(Path(抽出元).name.lower() == DLLファイル.name.lower() for 抽出元, 未使用 in 同梱バイナリ):
                同梱バイナリ.append((str(DLLファイル), "."))
同梱データ = collect_data_files("oletools")
for 名前 in ("oletools", "olefile", "msoffcrypto-tool", "cryptography", "pyOpenVBA"):
    同梱データ += copy_metadata(名前)
同梱データ += [("LICENSE", "."), ("vba2src_app/title.txt", "vba2src_app")]
ライセンス通知 = Path(SPECPATH) / "build/THIRD_PARTY_NOTICES.txt"
if ライセンス通知.is_file():
    同梱データ += [(str(ライセンス通知), ".")]
# ── 2. ウィンドウ版の単一EXE ──────────────────────────────
# コンソールは起動時の引数に応じてruntime.pyで接続する。
解析定義 = Analysis(["vba2src.py"], pathex=[SPECPATH], binaries=同梱バイナリ, datas=同梱データ,
                    hiddenimports=[],
                    hookspath=[], runtime_hooks=[], excludes=["tkinter"], noarchive=False)
Pythonアーカイブ = PYZ(解析定義.pure)
# リダイレクトされた標準入出力でも、日本語のパスをUTF-8で扱う。
実行オプション = [("X utf8", None, "OPTION")]
実行ファイル = EXE(Pythonアーカイブ, 解析定義.scripts, 解析定義.binaries, 解析定義.datas, 実行オプション, name="VBA2SRC",
          debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
          console=False, disable_windowed_traceback=True)
