# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

import pandas as pd


def main():
    if len(sys.argv) < 2:
        print("Usage: python rebuild_reports.py results/<timestamp>")
        return

    run_dir = Path(sys.argv[1])
    jsonl = run_dir / "checkpoint.jsonl"
    if not jsonl.exists():
        print(f"checkpoint.jsonl not found: {jsonl}")
        return

    rows = []
    with jsonl.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))

    if not rows:
        print("No rows in checkpoint")
        return

    # Категоризация
    by_status = {}
    for r in rows:
        st = r.get("final_status", "unknown")
        by_status.setdefault(st, []).append(r.get("email"))

    # XLSX
    def save(name, emails):
        pd.DataFrame({"Email": sorted(set(emails))}).to_excel(run_dir / name, index=False)

    save("01_valid_emails_REBUILT.xlsx", by_status.get("valid", []))
    save("02_invalid_emails_REBUILT.xlsx", by_status.get("invalid", []))
    save("03_temporary_emails_REBUILT.xlsx", by_status.get("temporary", []))
    save("04_unknown_emails_REBUILT.xlsx", by_status.get("unknown", []))
    save("05_error_emails_REBUILT.xlsx", by_status.get("error", []))

    # Полный отчёт
    pd.DataFrame({
        "email": [r.get("email") for r in rows],
        "final_status": [r.get("final_status") for r in rows],
        "confidence": [r.get("confidence") for r in rows],
        "filter_reason": [r.get("filter_reason") for r in rows],
    }).to_excel(run_dir / "10_report_REBUILT.xlsx", index=False)

    print("Rebuilt reports saved to:", run_dir)


if __name__ == "__main__":
    main()
