"""Windows配布EXEを、呼び出した端末またはリダイレクト先につなぐ。"""

import ctypes
import msvcrt
import os
import sys
from ctypes import wintypes
from pathlib import Path

無効ハンドル = ctypes.c_void_p(-1).value


class プロセス一覧項目(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("cntUsage", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("th32DefaultHeapID", ctypes.c_size_t),
        ("th32ModuleID", wintypes.DWORD),
        ("cntThreads", wintypes.DWORD),
        ("th32ParentProcessID", wintypes.DWORD),
        ("pcPriClassBase", wintypes.LONG),
        ("dwFlags", wintypes.DWORD),
        ("szExeFile", wintypes.WCHAR * 260),
    ]


def 配布EXEの呼出元を取得(Windows):
    """単一EXEの展開プロセスを経由して起動した場合に、その呼び出し元を特定する。"""
    Windows.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
    Windows.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    Windows.Process32FirstW.argtypes = Windows.Process32NextW.argtypes = (
        wintypes.HANDLE,
        ctypes.POINTER(プロセス一覧項目),
    )
    スナップショット = Windows.CreateToolhelp32Snapshot(0x00000002, 0)
    if スナップショット in (None, 無効ハンドル):
        return None
    try:
        親ID = os.getppid()
        EXE名 = Path(sys.executable).name.casefold()
        項目 = プロセス一覧項目(dwSize=ctypes.sizeof(プロセス一覧項目))
        取得済み = Windows.Process32FirstW(スナップショット, ctypes.byref(項目))
        while 取得済み:
            if 項目.th32ProcessID == 親ID:
                return 項目.th32ParentProcessID if 項目.szExeFile.casefold() == EXE名 else None
            取得済み = Windows.Process32NextW(スナップショット, ctypes.byref(項目))
    finally:
        Windows.CloseHandle(スナップショット)
    return None


def 標準入出力を接続():
    Windows = ctypes.WinDLL("kernel32", use_last_error=True)
    Windows.GetStdHandle.argtypes = (wintypes.DWORD,)
    Windows.GetStdHandle.restype = wintypes.HANDLE
    Windows.GetFileType.argtypes = Windows.CloseHandle.argtypes = (wintypes.HANDLE,)
    Windows.AttachConsole.argtypes = (wintypes.DWORD,)
    Windows.GetCurrentProcess.restype = wintypes.HANDLE
    Windows.DuplicateHandle.argtypes = (
        wintypes.HANDLE,
        wintypes.HANDLE,
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.HANDLE),
        wintypes.DWORD,
        wintypes.BOOL,
        wintypes.DWORD,
    )

    # 明示されたファイル・パイプへの出力を、端末に接続する前に確保する。
    ストリーム一覧 = (("stdout", -11, "w"), ("stderr", -12, "w"), ("stdin", -10, "r"))
    現在プロセス = Windows.GetCurrentProcess()
    for 名前, ハンドル番号, モード in ストリーム一覧:
        if getattr(sys, 名前) is not None:
            continue
        ハンドル = Windows.GetStdHandle(wintypes.DWORD(ハンドル番号))
        if ハンドル in (None, 無効ハンドル) or Windows.GetFileType(ハンドル) not in (1, 3):
            continue
        複製 = wintypes.HANDLE()
        if not Windows.DuplicateHandle(現在プロセス, ハンドル, 現在プロセス, ctypes.byref(複製), 0, False, 2):
            continue
        ファイル記述子 = None
        try:
            ファイル記述子 = msvcrt.open_osfhandle(複製.value, os.O_RDONLY if モード == "r" else os.O_WRONLY)
            setattr(sys, 名前, os.fdopen(ファイル記述子, モード, encoding="utf-8", buffering=1))
        except (OSError, ValueError):
            if ファイル記述子 is None:
                Windows.CloseHandle(複製)
            else:
                os.close(ファイル記述子)

    if not Windows.GetConsoleCP() and not Windows.AttachConsole(wintypes.DWORD(-1)):
        呼出元 = 配布EXEの呼出元を取得(Windows)
        if 呼出元 is None or not Windows.AttachConsole(呼出元):
            return
    for 名前, 未使用, モード in ストリーム一覧:
        if getattr(sys, 名前) is None:
            try:
                # Win32のコンソールデバイス名を指定し、日本語をUnicodeで渡す。
                デバイス = "CONIN$" if モード == "r" else "CONOUT$"
                setattr(sys, 名前, open(デバイス, モード, encoding="utf-8", buffering=1))
            except OSError:
                pass
