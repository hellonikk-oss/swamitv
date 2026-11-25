


import pandas as pd
import os


def download_Final_district():
    # Check if the file exists before trying to remove it
    if os.path.exists('./file/final_district.xlsx'):
        os.remove('./file/final_district.xlsx')

    # Load all of df1
    df1 = pd.read_excel('./file/district.xls', sheet_name='Data')

    # Load ONLY 'District Name' (for merging) and 'Completed' from df2
    df2 = pd.read_excel('./file/district_old.xls', sheet_name='Data', usecols=['District Name', 'Completed'])

    # Merge them. Because df2 only has 'Completed', only that column gets the '_old' suffix
    df3 = pd.merge(df1, df2, how='outer', on="District Name", suffixes=('', '_old'))

    # 1. Calculate the progress (Completed - Completed_old)
    df3['Progress'] = df3['Completed'] - df3['Completed_old']

    # 2. Find the index position of 'Completion %'
    target_idx = df3.columns.get_loc('Completion %')

    # 3. Move 'Completed_old' before 'Completion %'
    completed_old_col = df3.pop('Completed_old')
    df3.insert(target_idx, 'Completed_old', completed_old_col)

    # 4. Move the new 'Progress' column to sit next to them
    progress_col = df3.pop('Progress')
    df3.insert(target_idx + 1, 'Progress', progress_col)

    # Save the final output
    df3.to_excel('./file/final_district.xlsx', index=False, sheet_name='Data')



def download_Final_block():
    # Check if the file exists before trying to remove it
    if os.path.exists('./file/final_block.xlsx'):
        os.remove('./file/final_block.xlsx')

    # Load all of df1
    df1 = pd.read_excel('./file/block.xls', sheet_name='Data')

    # Load ONLY 'Block Name' (for merging) and 'Completed' from df2
    df2 = pd.read_excel('./file/block_old.xls', sheet_name='Data', usecols=['Tehsil Name', 'Completed'])

    # Merge them. Because df2 only has 'Completed', only that column gets the '_old' suffix
    df3 = pd.merge(df1, df2, how='outer', on="Tehsil Name", suffixes=('', '_old'))

    # 1. Calculate the progress (Completed - Completed_old)
    df3['Progress'] = df3['Completed'] - df3['Completed_old']

    # 2. Find the index position of 'Completion %'
    target_idx = df3.columns.get_loc('Completion %')

    # 3. Move 'Completed_old' before 'Completion %'
    completed_old_col = df3.pop('Completed_old')
    df3.insert(target_idx, 'Completed_old', completed_old_col)

    # 4. Move the new 'Progress' column to sit next to them
    progress_col = df3.pop('Progress')
    df3.insert(target_idx + 1, 'Progress', progress_col)
    df3.to_excel('./file/final_block.xlsx', index=False, sheet_name='Data')

    # Save the final output
   
    source_file = './file/final_block.xlsx'
    typedate_file= './file/svamitva_registry_report.xlsx'
    df2= pd.read_excel(typedate_file, sheet_name='detail', usecols=['Tehsil Name', 'type'])
    df4 = pd.merge(df3, df2, how='outer', on="Tehsil Name", suffixes=('', '_old'))
    df4 = df4[df4['type'].isin(['1urban', 'rural'])]
    df4 = df4.sort_values(by=['type', 'Completion %'], ascending=[True, True]).reset_index(drop=True)
    with pd.ExcelWriter(source_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
      df4.to_excel(writer, index=False, sheet_name='sheet1')
        

download_Final_block()
download_Final_district()