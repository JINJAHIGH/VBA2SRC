"""引数なし起動で使うTkinter画面の入口。"""

import tkinter as tk

from ..runtime import 安全に表示
from .window import メイン画面


def GUIを起動():
    try:
        ウィンドウ = tk.Tk()
    except tk.TclError as 例外情報:
        安全に表示(f"GUIを起動できませんでした。{例外情報}", True)
        return 1
    メイン画面(ウィンドウ)
    ウィンドウ.mainloop()
    return 0
