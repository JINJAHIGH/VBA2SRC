"""Python実行と配布EXEで共通の起動点。"""

import sys

if sys.version_info < (3, 13):
    raise SystemExit("VBA2SRC の実行には Python 3.13 以上が必要です。")

from vba2src_app.cli import メイン  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(メイン())
