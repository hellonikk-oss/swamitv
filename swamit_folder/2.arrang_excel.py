

import pandas as pd
import os

# Automatically locate the 'file' folder (targets swamitv/file)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DIR = os.path.join(BASE_DIR, 'file')
os.makedirs(FILE_DIR, exist_ok=True)


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

    df1 = pd.read_excel(district_file, sheet_name='Data')
    df2 = pd.read_excel(district_old_file, sheet_name='Data', usecols=['District Name', 'Completed'])

    df3 = pd.merge(df1, df2, how='outer', on='District Name', suffixes=('', '_old'))
    df3['Progress'] = df3['complet'] - df3['Completed_old']

    target_idx = df3.columns.get_loc('Completion %')
    df3.insert(target_idx, 'Completed_old', df3.pop('Completed_old'))
    df3.insert(target_idx + 1, 'Progress', df3.pop('Progress'))

    # Sort by Completion % while keeping 'Total' at the very last row
    df3 = sort_with_total_at_bottom(df3, 'District Name', sort_cols=['Completion %'], ascending=[True])

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

    df1 = pd.read_excel(block_file, sheet_name='Data')
    df2 = pd.read_excel(block_old_file, sheet_name='Data', usecols=['Tehsil Name', 'Completed'])

    df3 = pd.merge(df1, df2, how='outer', on='Tehsil Name', suffixes=('', '_old'))
    df3['Progress'] = df3['Completed'] - df3['Completed_old']

    target_idx = df3.columns.get_loc('Completion %')
    df3.insert(target_idx, 'Completed_old', df3.pop('Completed_old'))
    df3.insert(target_idx + 1, 'Progress', df3.pop('Progress'))

    # Keep 'Total' at the bottom of the Data sheet
    df3 = sort_with_total_at_bottom(df3, 'Tehsil Name')
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
    