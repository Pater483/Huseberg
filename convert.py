#!/usr/bin/env python3
"""
convert.py – Läser bookings.xlsx och skriver data.json
Körs automatiskt av GitHub Actions vid varje push.
"""

import json
import sys
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl saknas. Kör: pip install openpyxl")
    sys.exit(1)


EXCEL_FILE = "bookings.xlsx"
JSON_FILE  = "data.json"

# Kolumnnamn i Excel → nycklar i JSON (case-insensitive)
COL_MAP = {
    "hus":        "house",
    "namn":       "name",
    "antal":      "antal",
    "från":       "from",
    "fran":       "from",
    "till":       "to",
    "kommentar":  "kommentar",
}


def to_date_str(value):
    """Konverterar Excel-datum eller sträng till YYYY-MM-DD."""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        if isinstance(value, datetime):
            value = value.date()
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s


def main():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    except FileNotFoundError:
        print(f"ERROR: Hittar inte {EXCEL_FILE}")
        sys.exit(1)

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        print("WARNING: Excel-filen är tom.")
        bookings = []
    else:
        # Hitta rubrikrad (första icke-tomma rad)
        header_row = None
        data_start = 0
        for i, row in enumerate(rows):
            if any(cell is not None for cell in row):
                header_row = [str(c).strip().lower() if c else "" for c in row]
                data_start = i + 1
                break

        if header_row is None:
            print("WARNING: Ingen rubrikrad hittades.")
            bookings = []
        else:
            bookings = []
            for row in rows[data_start:]:
                # Hoppa över helt tomma rader
                if not any(cell is not None for cell in row):
                    continue

                entry = {}
                for col_i, header in enumerate(header_row):
                    key = COL_MAP.get(header)
                    if key is None:
                        continue
                    val = row[col_i] if col_i < len(row) else None
                    if key in ("from", "to"):
                        val = to_date_str(val)
                    elif val is not None:
                        val = str(val).strip()
                    entry[key] = val if val is not None else ""

                # Kräv minst hus + from + to
                if entry.get("house") and entry.get("from") and entry.get("to"):
                    bookings.append(entry)

    output = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "bookings": bookings,
    }

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"✓ {len(bookings)} bokningar skrivna till {JSON_FILE}")


if __name__ == "__main__":
    main()
