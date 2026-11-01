# -*- coding: utf-8 -*-
"""
Build index.html (MLB-style grade lookup site) from "Grade review.xlsx".

Access is enforced by encryption rather than by the UI: each record's key is
PBKDF2(學號 + "|" + 查詢碼), so a student can only decrypt their own record,
and only with their own code. The published page contains nothing else --
there is deliberately no all-class blob on the public site.

The teacher's own view of the whole class is produced locally instead, as
成績總表.csv next to this script.

Codes are kept in secrets.json so they stay stable across rebuilds.
secrets.json is gitignored and must never be published.

Usage:  python build.py
"""

import hashlib
import hmac
import json
import os
import re
import secrets
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Grade review.xlsx")
TEMPLATE = os.path.join(HERE, "template.html")
OUTPUT = os.path.join(HERE, "index.html")
SECRETS = os.path.join(HERE, "secrets.json")
HANDOUT = os.path.join(HERE, "查詢碼名單.csv")
SUMMARY = os.path.join(HERE, "成績總表.csv")

STUDENT_ITERS = 60000      # mirrored in template.html via DATA.iters
DK_LEN = 64                # 32-byte cipher key + 32-byte lookup token

# No 0/O/1/I/L — unambiguous when read off a screen or printed
ALPHABET = "23456789ABCDEFGHJKMNPQRSTUVWXYZ"
CODE_LEN = 8               # 31^8 ~ 8.5e11 combinations

COL_CLASS, COL_ID, COL_NAME, COL_GROUP = 0, 1, 2, 3
FIRST_ITEM_COL = 4
DEDUCTION_HEADER = "點名扣分"
TOTAL_HEADER = "學期總分"


# ---------------------------------------------------------------- crypto ----

def norm_code(text):
    """Codes are compared case-insensitively with separators ignored."""
    return re.sub(r"[^0-9A-Z]", "", str(text).upper())


def derive(password: str, salt: bytes, iters: int) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt,
                               iters, dklen=DK_LEN)


def keystream(key: bytes, length: int) -> bytes:
    out = bytearray()
    counter = 0
    while len(out) < length:
        out += hashlib.sha256(key + counter.to_bytes(4, "big")).digest()
        counter += 1
    return bytes(out[:length])


def encrypt(key: bytes, plaintext: bytes) -> tuple:
    ct = bytes(a ^ b for a, b in zip(plaintext, keystream(key, len(plaintext))))
    mac = hmac.new(key, ct, hashlib.sha256).digest()[:16]
    return ct, mac


def make_code(length=CODE_LEN):
    return "".join(secrets.choice(ALPHABET) for _ in range(length))


def group4(text):
    return "-".join(text[i:i + 4] for i in range(0, len(text), 4))


def total_of(stu, items):
    """(目前總分, 已評分權重) — 未評分項目以 0 分計入，點名扣分直接扣除。"""
    total = graded_w = 0.0
    for score, item in zip(stu["scores"], items):
        if score is not None:
            total += score * item["weight"] / 100.0
            graded_w += item["weight"]
    if stu["deduction"] is not None:
        total -= abs(stu["deduction"])
    return total, graded_w


# ------------------------------------------------------------- xlsx read ----

def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.replace("　", " ").strip()
        return value or None
    return value


def parse_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("%", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def read_workbook():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("Grade review.xlsx 第一個工作表是空的。")

    header = rows[0]
    items, deduction_col = [], None

    for idx in range(FIRST_ITEM_COL, len(header)):
        raw = header[idx]
        if raw is None:
            continue
        flat = re.sub(r"\s+", " ", str(raw).replace("　", " ").strip())
        if not flat or TOTAL_HEADER in flat:
            continue
        if DEDUCTION_HEADER in flat:
            deduction_col = idx
            continue
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*%", flat)
        if not match:
            continue
        label = re.sub(r"[（(]?\s*[0-9]+(?:\.[0-9]+)?\s*%\s*[）)]?", "", flat).strip(" .、·-")
        items.append({"col": idx, "label": label or flat,
                      "weight": float(match.group(1))})

    if not items:
        sys.exit("找不到任何帶有百分比的成績欄位，請確認標題列格式（例如「口頭簡報（50%）」）。")

    students, seen = [], {}
    for row_no, row in enumerate(rows[1:], start=2):
        raw_id = row[COL_ID] if len(row) > COL_ID else None
        if raw_id is None:
            continue
        if isinstance(raw_id, float) and raw_id.is_integer():
            sid = str(int(raw_id))
        elif isinstance(raw_id, int) and not isinstance(raw_id, bool):
            sid = str(raw_id)
        else:
            sid = re.sub(r"\s+", "", str(raw_id))
        if not sid:
            continue
        if sid in seen:
            print(f"  ! 第 {row_no} 列學號 {sid} 重複，已略過（保留第 {seen[sid]} 列）。")
            continue
        seen[sid] = row_no

        name = clean(row[COL_NAME] if len(row) > COL_NAME else None) or ""
        students.append({
            "id": sid,
            "name": re.sub(r"\s*\*+$", "", name).strip() or name,
            "marked": bool(re.search(r"\*+$", name)),
            "class": clean(row[COL_CLASS] if len(row) > COL_CLASS else None) or "",
            "group": clean(row[COL_GROUP] if len(row) > COL_GROUP else None),
            "scores": [parse_number(row[it["col"]]) if len(row) > it["col"] else None
                       for it in items],
            "deduction": (parse_number(row[deduction_col])
                          if deduction_col is not None and len(row) > deduction_col
                          else None),
        })

    return items, students


# ---------------------------------------------------------------- secrets ---

def load_secrets():
    if not os.path.exists(SECRETS):
        return {"codes": {}}
    with open(SECRETS, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    data.setdefault("codes", {})
    data.pop("admin_password", None)   # 網站已不再有全班名冊，主密碼不再需要
    return data


def save_secrets(data):
    with open(SECRETS, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------- build ----

def main():
    for path, what in ((XLSX, "成績檔"), (TEMPLATE, "版型檔")):
        if not os.path.exists(path):
            sys.exit(f"找不到{what}：{path}")

    items, students = read_workbook()
    store = load_secrets()

    new_codes = []
    for stu in students:
        if stu["id"] not in store["codes"]:
            store["codes"][stu["id"]] = make_code()
            new_codes.append(stu["id"])
    save_secrets(store)

    print(f"成績項目（共 {len(items)} 項，權重合計 {sum(i['weight'] for i in items):g}%）：")
    for it in items:
        print(f"  - {it['label']}：{it['weight']:g}%")

    salt = secrets.token_bytes(16)

    records = {}
    for stu in students:
        code = store["codes"][stu["id"]]
        payload = {
            "id": stu["id"], "name": stu["name"], "marked": stu["marked"],
            "cls": stu["class"], "group": stu["group"],
            "scores": stu["scores"], "deduction": stu["deduction"],
        }
        dk = derive(stu["id"] + "|" + norm_code(code), salt, STUDENT_ITERS)
        key, token = dk[:32], dk[32:]
        blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        ct, mac = encrypt(key, blob)
        records[token.hex()] = {"c": ct.hex(), "m": mac.hex()}

    data = {
        "salt": salt.hex(), "iters": STUDENT_ITERS,
        "items": [{"label": i["label"], "weight": i["weight"]} for i in items],
        "records": records,
    }

    with open(TEMPLATE, "r", encoding="utf-8") as fh:
        html = fh.read()
    if "__GRADE_DATA__" not in html:
        sys.exit("template.html 裡找不到 __GRADE_DATA__ 佔位字串。")
    html = html.replace("__GRADE_DATA__",
                        json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    with open(OUTPUT, "w", encoding="utf-8") as fh:
        fh.write(html)

    def csv_cell(value):
        return '"' + str("" if value is None else value).replace('"', '""') + '"'

    # Distribution list for Moodle (never published — gitignored)
    with open(HANDOUT, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("班級,學號,姓名,查詢碼\n")
        for stu in students:
            fh.write(",".join(csv_cell(v) for v in (
                stu["class"], stu["id"], stu["name"],
                group4(store["codes"][stu["id"]]))) + "\n")

    # The teacher's own full-class view, generated locally instead of published
    header = ["班級", "學號", "姓名", "組"]
    header += [f'{i["label"]}({i["weight"]:g}%)' for i in items]
    header += ["點名扣分", "已評分權重%", "目前總分"]
    with open(SUMMARY, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(",".join(csv_cell(h) for h in header) + "\n")
        for stu in sorted(students, key=lambda s: -total_of(s, items)[0]):
            total, graded_w = total_of(stu, items)
            deduct = abs(stu["deduction"]) if stu["deduction"] is not None else None
            row = [stu["class"], stu["id"], stu["name"], stu["group"]]
            row += ["" if v is None else f"{v:g}" for v in stu["scores"]]
            row += ["" if not deduct else f"{deduct:g}",
                    f"{graded_w:g}", f"{round(total, 2):g}"]
            fh.write(",".join(csv_cell(v) for v in row) + "\n")

    graded = sum(1 for s in students if any(v is not None for v in s["scores"]))
    print(f"\n學生人數：{len(students)}（其中 {graded} 位已有成績資料）")
    if new_codes:
        print(f"新配發查詢碼：{len(new_codes)} 筆 —— {', '.join(new_codes[:8])}"
              + (" …" if len(new_codes) > 8 else ""))
    else:
        print("查詢碼：沿用 secrets.json 中既有的碼（學生手上的碼仍然有效）")
    print(f"\n已輸出：{OUTPUT}")
    print(f"全班成績總表（本機用，勿上傳）：{SUMMARY}")
    print(f"查詢碼名單（勿上傳）：{HANDOUT}")
    print(f"查詢碼保管檔（勿上傳）：{SECRETS}")


if __name__ == "__main__":
    main()
