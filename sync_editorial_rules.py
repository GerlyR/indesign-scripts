# -*- coding: utf-8 -*-
"""
Convert editorial_rules.xlsx -> editorial_rules.txt
Runs automatically before spellcheck; skips if xlsx is older than txt.
"""
import os
import sys


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(script_dir, "editorial_rules.xlsx")
    txt_path = os.path.join(script_dir, "editorial_rules.txt")

    if not os.path.exists(xlsx_path):
        return

    # Only convert if xlsx is newer than txt
    if os.path.exists(txt_path):
        if os.path.getmtime(xlsx_path) <= os.path.getmtime(txt_path):
            return

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("openpyxl not installed, skipping xlsx sync", file=sys.stderr)
        return

    try:
        wb = load_workbook(xlsx_path, read_only=True)
    except Exception as e:
        print(f"Cannot read xlsx (file locked?): {e}", file=sys.stderr)
        return

    try:
        ws = wb.active

        lines = [
            "# Auto-generated from editorial_rules.xlsx\n",
            "# DO NOT EDIT — edit the xlsx file instead\n",
        ]

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue

            col_a = str(row[0]).strip() if row[0] is not None else ""
            col_b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            col_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""

            if not col_b:
                continue

            rule_type = col_a.upper()

            if rule_type == "GREP" and col_c:
                lines.append(f"GREP\t{col_b}\t{col_c}\n")
            elif rule_type in ("TEXT", "\u0422\u0415\u041A\u0421\u0422", "") and col_b and col_c:
                lines.append(f"{col_b}\t{col_c}\n")
    finally:
        wb.close()

    try:
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.writelines(lines)
    except OSError as e:
        print(f"Cannot write {txt_path}: {e}", file=sys.stderr)
        return

    print(f"Synced {len(lines) - 2} rules from xlsx", file=sys.stderr)


if __name__ == '__main__':
    main()
