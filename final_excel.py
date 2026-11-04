import os
import sys
import pandas as pd

# Base directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_DIR = os.path.join(BASE_DIR, 'file')


def find_existing_file(filename):
    """
    Locates a file in FILE_DIR, checking for both .xls and .xlsx extensions.
    """
    path = os.path.join(FILE_DIR, filename)
    if os.path.exists(path):
        return path
    if filename.endswith('.xls'):
        alt = os.path.join(FILE_DIR, filename + 'x')
        if os.path.exists(alt):
            return alt
    elif filename.endswith('.xlsx'):
        alt = os.path.join(FILE_DIR, filename[:-1])
        if os.path.exists(alt):
            return alt
    return path


def load_excel_file(file_path):
    """
    Robustly loads an Excel file.
    Downloaded files often have .xls extension but are actually OpenXML (xlsx/zip) or HTML tables.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    # Try openpyxl first as government portals frequently serve .xlsx with .xls extension
    engines = ['openpyxl', 'xlrd', None]
    last_error = None

    for engine in engines:
        try:
            # Try reading "Data" sheet
            try:
                return pd.read_excel(file_path, sheet_name="Data", engine=engine)
            except Exception:
                # Fallback to default sheet
                return pd.read_excel(file_path, engine=engine)
        except Exception as e:
            last_error = e
            continue

    # Try HTML table fallback if it was saved as HTML
    try:
        tables = pd.read_html(file_path)
        if tables:
            return tables[0]
    except Exception:
        pass

    raise ValueError(f"Could not read excel file '{file_path}'. Error: {last_error}")


def process_report(current_filename, old_filename, output_filename, id_column):
    """
    Processes current and old report files, computes Progress, and generates the final output excel.
    """
    os.makedirs(FILE_DIR, exist_ok=True)

    current_path = find_existing_file(current_filename)
    old_path = find_existing_file(old_filename)
    output_path = os.path.join(FILE_DIR, output_filename)

    if not os.path.exists(current_path):
        print(f"[!] Warning: Current file '{current_path}' not found. Skipping.")
        return False

    print(f"Reading current report: {current_path}")
    df_1 = load_excel_file(current_path)
    df_1.columns = [str(c).strip() for c in df_1.columns]

    # Resolve ID column (District Name or Tehsil Name)
    match_col = id_column
    if match_col not in df_1.columns:
        candidates = [c for c in df_1.columns if id_column.lower() in c.lower()]
        if candidates:
            match_col = candidates[0]
        else:
            print(f"[!] Warning: '{id_column}' not found in columns: {list(df_1.columns)}")
            return False

    # Check and merge old report if available
    if os.path.exists(old_path):
        print(f"Reading previous report: {old_path}")
        try:
            df_2 = load_excel_file(old_path)
            df_2.columns = [str(c).strip() for c in df_2.columns]

            old_match_col = match_col if match_col in df_2.columns else None
            if not old_match_col:
                old_candidates = [c for c in df_2.columns if id_column.lower() in c.lower()]
                if old_candidates:
                    old_match_col = old_candidates[0]

            if old_match_col and 'Completed' in df_2.columns:
                df_1[match_col] = df_1[match_col].astype(str).str.strip()
                df_2[old_match_col] = df_2[old_match_col].astype(str).str.strip()

                df3 = pd.merge(
                    df_1,
                    df_2[[old_match_col, 'Completed']],
                    left_on=match_col,
                    right_on=old_match_col,
                    how='left',
                    suffixes=('', '_old')
                )
                if old_match_col != match_col and old_match_col in df3.columns:
                    df3.drop(columns=[old_match_col], inplace=True)
            else:
                df3 = df_1.copy()
                df3['Completed_old'] = 0
        except Exception as e:
            print(f"[!] Warning: Could not read previous report '{old_path}': {e}. Using 0 for previous completed.")
            df3 = df_1.copy()
            df3['Completed_old'] = 0
    else:
        print(f"Note: Previous file '{old_path}' not found. Using 0 for previous completed values.")
        df3 = df_1.copy()
        df3['Completed_old'] = 0

    # Ensure Completed and Completed_old are numeric
    df3['Completed_old'] = pd.to_numeric(df3.get('Completed_old', 0), errors='coerce').fillna(0).astype(int)
    completed_curr = pd.to_numeric(df3.get('Completed', 0), errors='coerce').fillna(0).astype(int)
    df3['Progress'] = completed_curr - df3['Completed_old']

    # Position Completed_old and Progress right before 'Completion %' or after 'Completed'
    completed_old_series = df3.pop('Completed_old')
    progress_series = df3.pop('Progress')

    if 'Completion %' in df3.columns:
        target_idx = df3.columns.get_loc('Completion %')
        df3.insert(target_idx, 'Completed_old', completed_old_series)
        df3.insert(target_idx + 1, 'Progress', progress_series)
    elif 'Completed' in df3.columns:
        target_idx = df3.columns.get_loc('Completed') + 1
        df3.insert(target_idx, 'Completed_old', completed_old_series)
        df3.insert(target_idx + 1, 'Progress', progress_series)
    else:
        df3['Completed_old'] = completed_old_series
        df3['Progress'] = progress_series

    # Remove existing output file if present
    if os.path.exists(output_path):
        try:
            print(f"Removing old {output_filename}...")
            os.remove(output_path)
        except PermissionError:
            print(f"[X] Error: Could not delete '{output_filename}'. Please close it in Excel and re-run.")
            return False

    # Save to Excel
    try:
        df3.to_excel(output_path, sheet_name="Data", index=False)
        print(f"[✓] Successfully generated: {output_path}")
        return True
    except PermissionError:
        print(f"[X] Error: Permission denied writing '{output_path}'. Please close the file if open in Excel.")
        return False
    except Exception as e:
        print(f"[X] Error saving '{output_path}': {e}")
        return False


def download_district_excel():
    print("\n--- Processing District Report ---")
    return process_report(
        current_filename="district.xls",
        old_filename="district_old.xls",
        output_filename="final_district.xlsx",
        id_column="District Name"
    )


def download_block_excel():
    print("\n--- Processing Block / Tehsil Report ---")
    return process_report(
        current_filename="block.xls",
        old_filename="block_old.xls",
        output_filename="final_block.xlsx",
        id_column="Tehsil Name"
    )


if __name__ == "__main__":
    download_district_excel()
    download_block_excel()
