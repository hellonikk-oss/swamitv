import openpyxl
import pandas as pd
import os

def finalreport():
    source_file = './file/final_block.xlsx'
    dest_file = './file/svamitva_registry_report.xlsx'
    source_wb = openpyxl.load_workbook(source_file)
    dest_wb = openpyxl.load_workbook(dest_file)
      
    source_sheet = source_wb['sheet1']  # Assuming the source sheet is named 'sheet1'
    if 'Data' in dest_wb.sheetnames:
        dest_sheet = dest_wb['Data']
        # This deletes all data from row 1 down to the very last row used
        if dest_sheet.max_row > 0:
            dest_sheet.delete_rows(1, dest_sheet.max_row)
    else:
        dest_sheet = dest_wb.create_sheet('Data')

    for row_idx, row_data in enumerate(source_sheet.iter_rows(values_only=True), start=1):
        for col_idx, cell_value in enumerate(row_data, start=1):
            dest_sheet.cell(row=row_idx, column=col_idx, value=cell_value)

    # Save the changes
    dest_wb.save('./file/svamitva_registry_report.xlsx')
    source_wb.close()
    dest_wb.close()
    
    print("Old data cleared safely, and new data successfully copied!")

finalreport()
