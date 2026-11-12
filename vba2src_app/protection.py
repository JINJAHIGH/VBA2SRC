"""MS-OVBAの保護フラグと検証情報を解析し、指定された候補だけを照合する。"""

import hashlib
import hmac
import re


# ── 1. 保存検証情報の復元 ──────────────────────────
def 難読化を復元(値):
    """MS-OVBAの保存キーと連鎖状態を用いてCMG・DPB・GCの内容を復元する。"""
    if not re.fullmatch(r"[0-9a-fA-F]+", 値) or len(値) % 2:
        raise ValueError("PROTECTION_HEX")
    データ = bytes.fromhex(値)
    if len(データ) < 8 or データ[0] ^ データ[1] != 2:
        raise ValueError("PROTECTION_VERSION")
    キー = データ[0] ^ データ[2]
    # 復元は保存キーで行う。PROJECT IDとの不一致は別に診断し、回収可能なフラグを残す。
    直前平文バイト, 直前暗号バイト, 前々暗号バイト = キー, データ[2], データ[1]
    復元データ = bytearray()
    for 暗号バイト in データ[3:]:
        平文データ = 暗号バイト ^ ((前々暗号バイト + 直前平文バイト) & 255)
        復元データ.append(平文データ)
        直前平文バイト, 前々暗号バイト, 直前暗号バイト = 平文データ, 直前暗号バイト, 暗号バイト
    無視データ = (データ[0] & 6) // 2
    長さデータ = 復元データ[無視データ : 無視データ + 4]
    if len(長さデータ) != 4:
        raise ValueError("PROTECTION_LENGTH")
    長さ = int.from_bytes(長さデータ, "little")
    ペイロード = bytes(復元データ[無視データ + 4 :])
    if 長さ != len(ペイロード):
        raise ValueError("PROTECTION_LENGTH")
    return ペイロード


def ハッシュ情報を復元(データ):
    """符号化された塩とSHA-1検証値を復元し、長さと終端を検証する。"""
    if len(データ) != 29 or データ[-1] != 0:
        raise ValueError("PASSWORD_HASH_STRUCTURE")
    # MS-OVBAのビット図に合わせ、上位ビットから検証値を復元する。
    フラグ群 = int.from_bytes(データ[1:4], "big")
    符号化データ = データ[4:28]
    復元データ = bytes(b if フラグ群 & (1 << (23 - i)) else 0 for i, b in enumerate(符号化データ))
    if any(b == 0 for b in 符号化データ):
        raise ValueError("PASSWORD_HASH_NULL_ENCODING")
    return 復元データ[:4], 復元データ[4:]


# ── 2. 保護判定と候補照合 ──────────────────────────
def 保護情報を解析(プロパティ, 文字コード, 候補一覧):
    """保護とパスワード設定を別々に判定し、候補の設定順で照合する。"""
    結果 = {
        "protection_state": "判定不能",
        "flags": None,
        "visibility": None,
        "password_set": None,
        "password_method": "判定不能",
        "password": "不明",
        "candidate_status": "判定不能",
        "fields": {},
        "diagnostics": [],
    }
    復元データ = {}
    for フィールド in ("CMG", "DPB", "GC"):
        値一覧 = プロパティ.get(フィールド, [])
        # 生の検証値と復元バイト列を結果やログへ出さない。
        結果["fields"][フィールド] = {"present": bool(値一覧), "count": len(値一覧), "decoded_length": None}
        if len(値一覧) != 1:
            結果["diagnostics"].append(f"{フィールド}_MISSING_OR_DUPLICATE")
            continue
        try:
            識別子一覧 = プロパティ.get("ID", [])
            復元データ[フィールド] = 難読化を復元(値一覧[0])
            結果["fields"][フィールド]["decoded_length"] = len(復元データ[フィールド])
            if len(識別子一覧) == 1:
                符号化データ = bytes.fromhex(値一覧[0])
                if 符号化データ[0] ^ 符号化データ[2] != sum(識別子一覧[0].encode("ascii")) % 256:
                    結果["diagnostics"].append(f"{フィールド}_PROJECT_KEY_INCONSISTENT")
        except (ValueError, UnicodeError):
            結果["diagnostics"].append(f"{フィールド}_INVALID")
    保護フラグデータ = 復元データ.get("CMG")
    if 保護フラグデータ is not None:
        if len(保護フラグデータ) == 4:
            フラグ群 = int.from_bytes(保護フラグデータ, "little")
            結果["flags"] = {
                "raw": フラグ群,
                "user": bool(フラグ群 & 1),
                "host": bool(フラグ群 & 2),
                "vbe": bool(フラグ群 & 4),
            }
            結果["protection_state"] = "あり" if フラグ群 & 7 else "なし"
            if フラグ群 & ~7:
                結果["diagnostics"].append("CMG_RESERVED_BITS")
        else:
            結果["diagnostics"].append("CMG_LENGTH")
    表示状態データ = 復元データ.get("GC")
    if 表示状態データ in (b"\x00", b"\xff"):
        結果["visibility"] = 表示状態データ == b"\xff"
        if 表示状態データ == b"\x00" and 結果["flags"] and not 結果["flags"]["vbe"]:
            結果["diagnostics"].append("GC_CMG_INCONSISTENT")
    elif 表示状態データ is not None:
        結果["diagnostics"].append("GC_INVALID_STATE")

    パスワードデータ = 復元データ.get("DPB")
    検証情報 = None
    if パスワードデータ == b"\x00":
        結果.update(password_set=False, password_method="なし", password="なし", candidate_status="対象外")
    elif パスワードデータ is not None:
        try:
            if len(パスワードデータ) == 29 and パスワードデータ[0] == 255:
                検証情報 = ("hash", ハッシュ情報を復元(パスワードデータ))
                結果["password_method"] = "MS-OVBA SHA-1"
                識別子一覧 = プロパティ.get("ID", [])
                if 識別子一覧 != ["{00000000-0000-0000-0000-000000000000}"]:
                    結果["diagnostics"].append("HASH_PROJECT_ID_INCONSISTENT")
            elif パスワードデータ and パスワードデータ[-1] == 0 and b"\0" not in パスワードデータ[:-1]:
                検証情報 = ("legacy", パスワードデータ[:-1])
                結果["password_method"] = "旧文字列方式"
            else:
                raise ValueError("PASSWORD_STRUCTURE")
            結果["password_set"] = True
        except ValueError:
            結果["diagnostics"].append("DPB_INVALID_STRUCTURE")

    if 結果["password_set"] is not False:
        if not 候補一覧 and 検証情報 is not None:
            結果["candidate_status"] = "スキップ（候補未設定）"
        elif 検証情報 is not None and 文字コード:
            符号化不能あり = False
            一致候補 = None
            for 候補 in 候補一覧:
                try:
                    符号化データ = 候補.encode(文字コード, errors="strict")
                except (UnicodeError, LookupError, AttributeError):
                    符号化不能あり = True
                    continue
                if b"\0" in 符号化データ:
                    符号化不能あり = True
                    continue
                if 検証情報[0] == "hash":
                    キー, 期待値 = 検証情報[1]
                    実際値 = hashlib.sha1(符号化データ + キー).digest()
                else:
                    実際値, 期待値 = 符号化データ, 検証情報[1]
                if hmac.compare_digest(実際値, 期待値) and 一致候補 is None:
                    一致候補 = 候補
            if 一致候補 is not None:
                結果.update(password=一致候補, candidate_status="一致")
            else:
                結果["candidate_status"] = "判定不能" if 符号化不能あり else "不一致"
            if 符号化不能あり:
                結果["diagnostics"].append("CANDIDATE_ENCODING_UNAVAILABLE")
    結果["diagnostics"].sort()
    return 結果
