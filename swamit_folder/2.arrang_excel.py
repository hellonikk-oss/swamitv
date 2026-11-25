import pandas as pd
import os
import json
import re
from datetime import datetime

# Automatically locate the 'file' folder (targets swamitv/file)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DIR = os.path.join(BASE_DIR, 'file')
os.makedirs(FILE_DIR, exist_ok=True)


def parse_and_format_timestamp(raw_value):
    """
    Parses a timestamp string or epoch float/int and formats it as:
    'Completed on DD-MM-YYYY at H:MM AM/PM' (e.g., 'Completed on 25-11-2026 at 5:00 PM').
    """
    if not raw_value:
        return None

    # If it's a file mtime (float/int)
    if isinstance(raw_value, (int, float)):
        dt = datetime.fromtimestamp(raw_value)
        date_str = dt.strftime('%d-%m-%Y')
        hour_str = dt.strftime('%I').lstrip('0')
        ampm = dt.strftime('%p')
        return f"Completed on {date_str} at {hour_str}:00 {ampm}"

    if isinstance(raw_value, str):
        val = raw_value.strip()
        if val.startswith("Completed on "):
            return val

        # Format: "25 Nov 2026, 05:00 PM IST" or "25 Nov 2026, 05:00 PM"
        m1 = re.search(r'(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4}),?\s+(\d{1,2}):(\d{2})\s*(AM|PM)?', val, re.IGNORECASE)
        if m1:
            day, month_str, year, hour, minute, ampm = m1.groups()
            ampm_part = ampm or ""
            dt = datetime.strptime(f"{day} {month_str[:3]} {year} {hour}:{minute} {ampm_part}".strip(),
                                   "%d %b %Y %I:%M %p" if ampm else "%d %b %Y %H:%M")
            date_str = dt.strftime('%d-%m-%Y')
            hour_str = dt.strftime('%I').lstrip('0')
            minute_ampm = dt.strftime('%M %p')
            return f"Completed on {date_str} at {hour_str}:{minute_ampm}"

        # Format: "25-11-2026 5:00 PM" or "25-11-2026 at 5:00 PM"
        m2 = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s*(?:at)?\s*(\d{1,2}):(\d{2})\s*(AM|PM)?', val, re.IGNORECASE)
        if m2:
            day, month, year, hour, minute, ampm = m2.groups()
            ampm_part = ampm or ""
            dt = datetime.strptime(f"{int(day):02d}-{int(month):02d}-{year} {hour}:{minute} {ampm_part}".strip(),
                                   "%d-%m-%Y %I:%M %p" if ampm else "%d-%m-%Y %H:%M")
            date_str = dt.strftime('%d-%m-%Y')
            hour_str = dt.strftime('%I').lstrip('0')
            minute_ampm = dt.strftime('%M %p')
            return f"Completed on {date_str} at {hour_str}:{minute_ampm}"

    return str(raw_value)


def get_column_headers(current_file=None, old_file=None):
    """
    Determines the header names for 'Completed' and 'Completed_old'.
    Priority:
    1. refresh_time.json (saved from 1.block_dis_excel_download.py)
    2. File modification timestamps of current_file and old_file.
    """
    time_file = os.path.join(FILE_DIR, 'refresh_time.json')
    current_time_raw = None
    old_time_raw = None

    if os.path.exists(time_file):
        try:
            with open(time_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                current_time_raw = data.get('current_time')
                old_time_raw = data.get('old_time')
        except Exception:
            pass

    # Fallbacks from file modification times if missing
    if not current_time_raw and current_file and os.path.exists(current_file):
        current_time_raw = os.path.getmtime(current_file)

    if not old_time_raw and old_file and os.path.exists(old_file):
        old_time_raw = os.path.getmtime(old_file)

    current_header = parse_and_format_timestamp(current_time_raw) or 'Completed'
    old_header = parse_and_format_timestamp(old_time_raw) or 'Completed_old'

    return current_header, old_header


def check_files_exist(file_paths):
    """Verify all required input files exist before processing."""
    missing = [p for p in file_paths if not os.path.exists(p)]
    if missing:
        print("Skipping step — missing required file(s):")
        for m in missing:
            print(f"  - {m}")
        print(f"Files currently inside '{FILE_DIR}': {os.listdir(FILE_DIR)}")
        return False
    return True


def sort_with_total_at_bottom(df, name_col, sort_cols=None, ascending=True):
    """Separates the 'Total' row, sorts the remaining rows, and places 'Total' at the bottom."""
    is_total = df[name_col].astype(str).str.strip().str.lower() == 'total'
    total_row = df[is_total]
    main_df = df[~is_total]

    if sort_cols:
        main_df = main_df.sort_values(by=sort_cols, ascending=ascending)

    return pd.concat([main_df, total_row], ignore_index=True)


def download_Final_district():
    district_file = os.path.join(FILE_DIR, 'district.xls')
    district_old_file = os.path.join(FILE_DIR, 'district_old.xls')
    output_file = os.path.join(FILE_DIR, 'final_district.xlsx')

    if not check_files_exist([district_file, district_old_file]):
        return

    if os.path.exists(output_file):
        os.remove(output_file)

    current_header, old_header = get_column_headers(district_file, district_old_file)

    df1 = pd.read_excel(district_file, sheet_name='Data')
    df2 = pd.read_excel(district_old_file, sheet_name='Data', usecols=['District Name', 'Completed'])

    df3 = pd.merge(df1, df2, how='outer', on='District Name', suffixes=('', '_old'))
    df3['Progress'] = df3['Completed'] - df3['Completed_old']

    target_idx = df3.columns.get_loc('Completion %')
    df3.insert(target_idx, 'Completed_old', df3.pop('Completed_old'))
    df3.insert(target_idx + 1, 'Progress', df3.pop('Progress'))

    # Sort by Completion % descending while keeping 'Total' at the very last row
    df3 = sort_with_total_at_bottom(df3, 'District Name', sort_cols=['Completion %'], ascending=[False])

    # Rename 'Completed' and 'Completed_old' to their timestamp headers
    df3 = df3.rename(columns={
        'Completed': current_header,
        'Completed_old': old_header
    })

    df3.to_excel(output_file, index=False, sheet_name='Data')
    print(f"Saved: {output_file}")


def download_Final_block():
    block_file = os.path.join(FILE_DIR, 'block.xls')
    block_old_file = os.path.join(FILE_DIR, 'block_old.xls')
    typedate_file = os.path.join(FILE_DIR, 'svamitva_registry_tehsilwise.xlsx')
    output_file = os.path.join(FILE_DIR, 'final_block.xlsx')

    if not check_files_exist([block_file, block_old_file, typedate_file]):
        return

    if os.path.exists(output_file):
        os.remove(output_file)

    current_header, old_header = get_column_headers(block_file, block_old_file)

    df1 = pd.read_excel(block_file, sheet_name='Data')
    df2 = pd.read_excel(block_old_file, sheet_name='Data', usecols=['Tehsil Name', 'Completed'])

    df3 = pd.merge(df1, df2, how='outer', on='Tehsil Name', suffixes=('', '_old'))
    df3['Progress'] = df3['Completed'] - df3['Completed_old']

    target_idx = df3.columns.get_loc('Completion %')
    df3.insert(target_idx, 'Completed_old', df3.pop('Completed_old'))
    df3.insert(target_idx + 1, 'Progress', df3.pop('Progress'))

    # Keep 'Total' at the bottom of the Data sheet
    df3 = sort_with_total_at_bottom(df3, 'Tehsil Name')

    # Rename 'Completed' and 'Completed_old' to their timestamp headers
    df3 = df3.rename(columns={
        'Completed': current_header,
        'Completed_old': old_header
    })

    df3.to_excel(output_file, index=False, sheet_name='Data')

    df_type = pd.read_excel(typedate_file, sheet_name='detail', usecols=['Tehsil Name', 'type'])
    df4 = pd.merge(df3, df_type, how='outer', on='Tehsil Name', suffixes=('', '_old'))

    # Keep 'Total' row even when filtering by '1urban' and 'rural'
    is_total = df4['Tehsil Name'].astype(str).str.strip().str.lower() == 'total'
    df4 = df4[df4['type'].isin(['1urban', 'rural']) | is_total]

    # Sort by type and Completion % while keeping 'Total' at the very last row
    df4 = sort_with_total_at_bottom(df4, 'Tehsil Name', sort_cols=['type', 'Completion %'], ascending=[True, True])

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df4.to_excel(writer, index=False, sheet_name='sheet1')
    print(f"Saved: {output_file}")


if __name__ == '__main__':
    download_Final_block()
    download_Final_district()
