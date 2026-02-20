#!/usr/bin/env python3
"""
Extract raw data from BAGRUT XLSX files for Western and Eastern Galilee regions.
Outputs CSV files with all raw rows so data can be verified against HTML dashboards.

XLSX structure:
- Sheet 1 (סע' 1): National-level subject data
- Sheet 2 (סע' 2): Authority-level subject data
- Sheet 3 (סע' 3): School-level subject data (key sheet for dashboard verification)
  Columns: ממוצע ציון סופי, מספר נבחנים, י"ל, מחזור סיום, מקצוע, רשות מקומית, שם מוסד, סמל מוסד, פיקוח, מגזר, מחוז תקשוב
- Sheet 4 (סע' 4): School-level eligibility data
"""
import openpyxl
import csv
import os

# Authorities for each region (from the HTML dashboards)
WEST_GALILEE_AUTHORITIES = {
    'עכו', 'נהריה', 'מעלות-תרשיחא', 'בית ג\'ן', 'חורפיש',
    'כפר ורדים', 'ירכא', 'אבו סנאן', 'כפר יאסיף', 'כסרא-סמיע',
    'ג\'וליס', 'מעיליא', 'מטה אשר', 'מעלה יוסף', 'ג\'דיידה-מכר',
}

EAST_GALILEE_AUTHORITIES = {
    'חצור הגלילית', 'קרית שמונה', 'קצרין', 'גולן', 'צפת',
    'הגליל העליון', 'מג\'דל שמס', 'בוקעאתה', 'מסעדה', 'גוש חלב',
}

# Column indices for sheet 3 (סע' 3) - school-level subject data
S3_COLS = {
    'score': 0,       # ממוצע ציון סופי
    'examinees': 1,   # מספר נבחנים
    'units': 2,       # י"ל (יחידות לימוד)
    'year': 3,        # מחזור סיום
    'subject': 4,     # מקצוע
    'authority': 5,   # רשות מקומית
    'school': 6,      # שם מוסד
    'school_code': 7, # סמל מוסד
    'supervision': 8, # פיקוח
    'sector': 9,      # מגזר
    'district': 10,   # מחוז תקשוב
}

# Column indices for sheet 2 (סע' 2) - authority-level subject data
# Headers: ממוצע ציון סופי, מספר נבחנים, י"ל, מחזור סיום, שם מקצוע, סמל מקצוע, שם רשות, מחוז חינוך
S2_COLS = {
    'score': 0,         # ממוצע ציון סופי
    'examinees': 1,     # מספר נבחנים
    'units': 2,         # י"ל
    'year': 3,          # מחזור סיום
    'subject': 4,       # שם מקצוע
    'subject_code': 5,  # סמל מקצוע
    'authority': 6,     # שם רשות
    'district': 7,      # מחוז חינוך
}


def strip(val):
    """Strip whitespace from a value if it's a string."""
    if isinstance(val, str):
        return val.strip()
    return val


def read_sheet_data(filepath, sheet_name):
    """Read data from a specific sheet, finding the header row automatically."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb[sheet_name]

    rows = []
    header_idx = None

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        row_list = list(row)
        # Look for header row by checking for known column names
        if header_idx is None:
            row_text = ' '.join(str(cell) for cell in row_list if cell is not None)
            if 'ממוצע ציון סופי' in row_text and 'מספר נבחנים' in row_text:
                header_idx = i
                headers = [str(h).strip() if h is not None else f'col_{j}' for j, h in enumerate(row_list)]
                continue
        elif header_idx is not None:
            # Only include rows that have actual data (first cell is not None)
            if row_list[0] is not None:
                rows.append(row_list)

    wb.close()
    return headers if header_idx is not None else [], rows


def save_csv(filepath, headers, rows):
    """Save rows to a CSV file."""
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([str(cell).strip() if cell is not None else '' for cell in row])


def extract_physics_school_data(base_dir, output_dir):
    """Extract Physics 5-unit school-level data from sheet 3 for both regions."""
    print("\n" + "=" * 80)
    print("EXTRACTING PHYSICS 5 YL SCHOOL-LEVEL DATA (Sheet 3)")
    print("=" * 80)

    all_physics_rows = []

    # BAGRUT.xlsx has years 2021-2023
    for fname, label in [('BAGRUT.xlsx', '2021-2023'), ('BAGRUT תשפד.xlsx', '2024')]:
        fpath = os.path.join(base_dir, fname)
        if not os.path.exists(fpath):
            print(f"  File not found: {fpath}")
            continue

        print(f"\n  Reading {fname} ({label})...")
        headers, rows = read_sheet_data(fpath, "סע' 3")
        print(f"  Total rows in sheet: {len(rows)}")

        # Filter for Physics 5-unit
        physics_rows = []
        for row in rows:
            subject = strip(row[S3_COLS['subject']]) if S3_COLS['subject'] < len(row) else ''
            units = row[S3_COLS['units']] if S3_COLS['units'] < len(row) else None
            if 'פיסיקה' in str(subject) and units == 5:
                physics_rows.append(row)

        print(f"  Physics 5YL rows: {len(physics_rows)}")
        all_physics_rows.extend(physics_rows)

    # Now filter by region
    west_rows = []
    east_rows = []

    for row in all_physics_rows:
        auth = strip(row[S3_COLS['authority']]) if S3_COLS['authority'] < len(row) else ''
        if auth in WEST_GALILEE_AUTHORITIES:
            west_rows.append(row)
        elif auth in EAST_GALILEE_AUTHORITIES:
            east_rows.append(row)

    # Sort by school name, then year
    west_rows.sort(key=lambda r: (strip(r[S3_COLS['school']]) or '', r[S3_COLS['year']] or 0))
    east_rows.sort(key=lambda r: (strip(r[S3_COLS['school']]) or '', r[S3_COLS['year']] or 0))

    headers_out = ['שם מוסד', 'רשות מקומית', 'פיקוח', 'מגזר', 'מחזור סיום', 'מספר נבחנים', 'ממוצע ציון סופי', 'סמל מוסד', 'מחוז תקשוב']

    def format_row(row):
        return [
            strip(row[S3_COLS['school']]),
            strip(row[S3_COLS['authority']]),
            strip(row[S3_COLS['supervision']]),
            strip(row[S3_COLS['sector']]),
            row[S3_COLS['year']],
            row[S3_COLS['examinees']],
            row[S3_COLS['score']],
            row[S3_COLS['school_code']],
            strip(row[S3_COLS['district']]),
        ]

    # Save West Galilee
    west_formatted = [format_row(r) for r in west_rows]
    west_path = os.path.join(output_dir, 'west_galilee_physics_5yl_raw.csv')
    save_csv(west_path, headers_out, west_formatted)
    print(f"\n  West Galilee Physics 5YL: {len(west_rows)} rows -> {west_path}")

    # Save East Galilee
    east_formatted = [format_row(r) for r in east_rows]
    east_path = os.path.join(output_dir, 'east_galilee_physics_5yl_raw.csv')
    save_csv(east_path, headers_out, east_formatted)
    print(f"  East Galilee Physics 5YL: {len(east_rows)} rows -> {east_path}")

    # Print the data for quick reference
    for region, rows_data in [('WEST GALILEE', west_formatted), ('EAST GALILEE', east_formatted)]:
        print(f"\n  --- {region} Physics 5YL Raw Data ---")
        print(f"  {'School':<30} {'Authority':<20} {'Supervision':<15} {'Sector':<10} {'Year':<6} {'Examinees':<10} {'Avg Score':<10}")
        print(f"  {'-'*30} {'-'*20} {'-'*15} {'-'*10} {'-'*6} {'-'*10} {'-'*10}")
        for r in rows_data:
            print(f"  {str(r[0]):<30} {str(r[1]):<20} {str(r[2]):<15} {str(r[3]):<10} {str(r[4]):<6} {str(r[5]):<10} {str(r[6]):<10}")

    return west_formatted, east_formatted


def extract_all_subjects_school_data(base_dir, output_dir):
    """Extract ALL subject school-level data from sheet 3 for both regions."""
    print("\n" + "=" * 80)
    print("EXTRACTING ALL SUBJECTS SCHOOL-LEVEL DATA (Sheet 3)")
    print("=" * 80)

    all_rows = []

    for fname, label in [('BAGRUT.xlsx', '2021-2023'), ('BAGRUT תשפד.xlsx', '2024')]:
        fpath = os.path.join(base_dir, fname)
        if not os.path.exists(fpath):
            continue

        print(f"\n  Reading {fname} ({label})...")
        headers, rows = read_sheet_data(fpath, "סע' 3")

        for row in rows:
            auth = strip(row[S3_COLS['authority']]) if S3_COLS['authority'] < len(row) else ''
            if auth in WEST_GALILEE_AUTHORITIES or auth in EAST_GALILEE_AUTHORITIES:
                region = 'west' if auth in WEST_GALILEE_AUTHORITIES else 'east'
                all_rows.append((region, row))

    headers_out = ['Region', 'שם מוסד', 'רשות מקומית', 'פיקוח', 'מגזר', 'מחזור סיום', 'מקצוע', 'י"ל', 'מספר נבחנים', 'ממוצע ציון סופי', 'סמל מוסד']

    west_all = []
    east_all = []

    for region, row in all_rows:
        formatted = [
            region,
            strip(row[S3_COLS['school']]),
            strip(row[S3_COLS['authority']]),
            strip(row[S3_COLS['supervision']]),
            strip(row[S3_COLS['sector']]),
            row[S3_COLS['year']],
            strip(row[S3_COLS['subject']]),
            row[S3_COLS['units']],
            row[S3_COLS['examinees']],
            row[S3_COLS['score']],
            row[S3_COLS['school_code']],
        ]
        if region == 'west':
            west_all.append(formatted)
        else:
            east_all.append(formatted)

    # Save all subjects
    west_path = os.path.join(output_dir, 'west_galilee_all_subjects_raw.csv')
    save_csv(west_path, headers_out, west_all)
    print(f"\n  West Galilee all subjects: {len(west_all)} rows -> {west_path}")

    east_path = os.path.join(output_dir, 'east_galilee_all_subjects_raw.csv')
    save_csv(east_path, headers_out, east_all)
    print(f"  East Galilee all subjects: {len(east_all)} rows -> {east_path}")


def extract_authority_level_data(base_dir, output_dir):
    """Extract authority-level subject data from sheet 2 for physics 5YL."""
    print("\n" + "=" * 80)
    print("EXTRACTING AUTHORITY-LEVEL PHYSICS 5YL DATA (Sheet 2)")
    print("=" * 80)

    all_rows = []

    for fname, label in [('BAGRUT.xlsx', '2021-2023'), ('BAGRUT תשפד.xlsx', '2024')]:
        fpath = os.path.join(base_dir, fname)
        if not os.path.exists(fpath):
            continue

        print(f"\n  Reading {fname} ({label})...")
        headers, rows = read_sheet_data(fpath, "סע' 2")
        print(f"  Total rows in sheet: {len(rows)}")

        for row in rows:
            subject = strip(row[S2_COLS['subject']]) if S2_COLS['subject'] < len(row) else ''
            units = row[S2_COLS['units']] if S2_COLS['units'] < len(row) else None
            auth = strip(row[S2_COLS['authority']]) if S2_COLS['authority'] < len(row) else ''

            if 'פיסיקה' in str(subject) and units == 5:
                if auth in WEST_GALILEE_AUTHORITIES or auth in EAST_GALILEE_AUTHORITIES:
                    region = 'west' if auth in WEST_GALILEE_AUTHORITIES else 'east'
                    all_rows.append((region, row))

    headers_out = ['Region', 'רשות מקומית', 'מחזור סיום', 'מספר נבחנים', 'ממוצע ציון סופי', 'מחוז חינוך']

    formatted_rows = []
    for region, row in all_rows:
        formatted_rows.append([
            region,
            strip(row[S2_COLS['authority']]),
            row[S2_COLS['year']],
            row[S2_COLS['examinees']],
            row[S2_COLS['score']],
            strip(row[S2_COLS['district']]) if S2_COLS['district'] < len(row) else '',
        ])

    csv_path = os.path.join(output_dir, 'both_galilee_authority_physics_5yl_raw.csv')
    save_csv(csv_path, headers_out, formatted_rows)
    print(f"\n  Authority-level physics 5YL: {len(formatted_rows)} rows -> {csv_path}")

    # Print data
    print(f"\n  {'Region':<8} {'Authority':<20} {'Year':<6} {'Examinees':<10} {'Avg Score':<10}")
    print(f"  {'-'*8} {'-'*20} {'-'*6} {'-'*10} {'-'*10}")
    for r in formatted_rows:
        print(f"  {str(r[0]):<8} {str(r[1]):<20} {str(r[2]):<6} {str(r[3]):<10} {str(r[4]):<10}")


def extract_eligibility_data(base_dir, output_dir):
    """Extract eligibility data from sheet 4 for both regions."""
    print("\n" + "=" * 80)
    print("EXTRACTING ELIGIBILITY DATA (Sheet 4)")
    print("=" * 80)

    for fname, label in [('BAGRUT.xlsx', '2021-2023'), ('BAGRUT תשפד.xlsx', '2024')]:
        fpath = os.path.join(base_dir, fname)
        if not os.path.exists(fpath):
            continue

        print(f"\n  Reading {fname} ({label})...")
        headers, rows = read_sheet_data(fpath, "סע' 4")
        print(f"  Total rows: {len(rows)}, Headers: {headers}")


def main():
    base_dir = '/home/user/BagrutWestGallile'
    output_dir = os.path.join(base_dir, 'raw_data')
    os.makedirs(output_dir, exist_ok=True)

    # Clear old files
    for f in os.listdir(output_dir):
        os.remove(os.path.join(output_dir, f))

    # 1. Physics 5YL school-level data (main dashboard data)
    extract_physics_school_data(base_dir, output_dir)

    # 2. All subjects school-level data (for full reference)
    extract_all_subjects_school_data(base_dir, output_dir)

    # 3. Authority-level physics data
    extract_authority_level_data(base_dir, output_dir)

    print(f"\n{'='*80}")
    print(f"All CSV files saved to: {output_dir}")
    print(f"{'='*80}")

    for f in sorted(os.listdir(output_dir)):
        fpath = os.path.join(output_dir, f)
        size = os.path.getsize(fpath)
        print(f"  {f} ({size:,} bytes)")


if __name__ == '__main__':
    main()
