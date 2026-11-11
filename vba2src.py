"""Python実行と配布EXEで共通の起動点。"""

from vba2src_app.cli import メイン

if __name__ == "__main__":
    raise SystemExit(メイン())
