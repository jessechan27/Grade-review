# -*- coding: utf-8 -*-
"""
Build index.html (MLB-style grade lookup site) from "Grade review.xlsx".

Each student's record is encrypted with a key derived from their own student ID,
so the published HTML does not expose anyone else's grades in plain text.

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

PBKDF2_ITERS = 60000          # mirrored in template.html
DK_LEN = 64                   # 32 bytes cipher key + 32 bytes lookup token

COL_CLASS, COL_ID, COL_NAME, COL_GROUP = 0, 1, 2, 3
FIRST_ITEM_COL = 4

DEDUCTION_HEADER = "點名扣分"
TOTAL_HEADER = "學期總分"


# ---------------------------------------------------------------- crypto ----

def derive(student_id: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", student_id.encode("utf-8"), salt,
                               PBKDF2_ITERS, dklen=DK_LEN)


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


# ------------------------------------------------------------- xlsx read ----

def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.replace("　", " ").strip()
        return value or None
    return value


def parse_number(value):
    """Return a float, or None when the cell is empty / non-numeric."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "")
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
        text = str(raw).replace("　", " ").strip()
        if not text:
            continue
        flat = re.sub(r"\s+", " ", text)
        if TOTAL_HEADER in flat:
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
        sid = clean(row[COL_ID] if len(row) > COL_ID else None)
        if sid is None:
            continue
        sid = re.sub(r"\s+", "", str(sid))
        if isinstance(row[COL_ID], float) and row[COL_ID].is_integer():
            sid = str(int(row[COL_ID]))
        if sid in seen:
            print(f"  ! 第 {row_no} 列學號 {sid} 重複，已略過（保留第 {seen[sid]} 列）。")
            continue
        seen[sid] = row_no

        name = clean(row[COL_NAME] if len(row) > COL_NAME else None) or ""
        display_name = re.sub(r"\s*\*+$", "", name).strip()

        students.append({
            "id": sid,
            "name": display_name or name,
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


# ----------------------------------------------------------------- build ----

def main():
    if not os.path.exists(XLSX):
        sys.exit(f"找不到 {XLSX}")
    if not os.path.exists(TEMPLATE):
        sys.exit(f"找不到 {TEMPLATE}")

    items, students = read_workbook()
    salt = secrets.token_bytes(16)

    print(f"成績項目（共 {len(items)} 項，權重合計 "
          f"{sum(i['weight'] for i in items):g}%）：")
    for it in items:
        print(f"  - {it['label']}：{it['weight']:g}%")

    records = {}
    for stu in students:
        dk = derive(stu["id"], salt)
        key, token = dk[:32], dk[32:]
        payload = json.dumps({
            "id": stu["id"],
            "name": stu["name"],
            "marked": stu["marked"],
            "cls": stu["class"],
            "group": stu["group"],
            "scores": stu["scores"],
            "deduction": stu["deduction"],
        }, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        ct, mac = encrypt(key, payload)
        records[token.hex()] = {"c": ct.hex(), "m": mac.hex()}

    graded = sum(1 for s in students if any(v is not None for v in s["scores"]))
    data = {
        "salt": salt.hex(),
        "iters": PBKDF2_ITERS,
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

    print(f"\n學生人數：{len(students)}（其中 {graded} 位已有成績資料）")
    print(f"已輸出：{OUTPUT}")


if __name__ == "__main__":
    main()
