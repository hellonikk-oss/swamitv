import openpyxl
import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DIR = os.path.join(BASE_DIR, 'file')

def finalreport_block():
    source_file = os.path.join(FILE_DIR, 'final_block.xlsx')
    dest_file = os.path.join(FILE_DIR, 'svamitva_registry_tehsilwise.xlsx')
    source_wb = openpyxl.load_workbook(source_file)
    dest_wb = openpyxl.load_workbook(dest_file)
      
    source_sheet = source_wb['sheet1'] if 'sheet1' in source_wb.sheetnames else source_wb.active
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
    dest_wb.save(dest_file)
    source_wb.close()
    dest_wb.close()
    
    print("Block data cleared safely, and new data successfully copied!")

def finalreport_district():
    source_file = os.path.join(FILE_DIR, 'final_district.xlsx')
    dest_file = os.path.join(FILE_DIR, 'Svamitva_Registry_District.xlsx')
    source_wb = openpyxl.load_workbook(source_file)
    dest_wb = openpyxl.load_workbook(dest_file)
      
    source_sheet = source_wb['Data'] if 'Data' in source_wb.sheetnames else source_wb.active
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
    dest_wb.save(dest_file)
    source_wb.close()
    dest_wb.close()
    
    print("District data cleared safely, and new data successfully copied!")

finalreport_block()
finalreport_district()
