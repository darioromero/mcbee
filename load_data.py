import os
import pandas as pd
import numpy as np
import xlrd
import re
from datetime import datetime

path = 'data/'

if __name__ == '__main__':

    xl_files = os.listdir(path)
    print(len(xl_files), xl_files)

            # column names for dataframe
    col_names = ['Well', 'Date', 'Gas - Mcf', 'Cond. - bpd', 'Water - bpd', 
                      'Comment1', 'Comment2', 'Comment3', 'Artif Lift']
    xl_wells = pd.DataFrame(columns=col_names)

    # read all excel files in path
    for xl_file in xl_files:

        path = os.path.join("data/", xl_file)
        xl_book = xlrd.open_workbook(filename=path, ragged_rows=False, verbosity=0)

        sheets = xl_book.sheet_names()

        # read every sheet in current workbook
        for sheet in sheets:

            xl_sheet = xl_book.sheet_by_name(sheet)
            print(f"{xl_file}/{sheet} has {xl_sheet.nrows} rows; {xl_sheet.ncols} cols")
            rowx = 0; coly = 0
            try:
                res = xl_sheet.cell_value(rowx, coly)
                # extract well name and artificial lift method
                pattern = r'(.*)\sDaily Prod\.(\s*[-=w\/]+\s*)?(.*)'
                match = re.search(pattern, res)
                if match:
                    well_name = match.group(1).strip()
                    artlift_mech = match.group(3).strip()
                    print(f"[{well_name}] | [{artlift_mech}]")
                
                new_row = {col_names[0]:well_name, col_names[8]:artlift_mech}
                
                num_cols = 7 # Number of columns: substitute xl_sheet.ncols
                num_rows = 365 # Number of rows: substitute xl_sheet.nrows
                for row_idx in range(num_rows):    # Iterate through rows
                    row = []
                    for col_idx in range(num_cols):  # Iterate through columns
                        cell_obj = xl_sheet.cell_value(row_idx + 2, col_idx)  # Get cell object by row, col
                        if col_idx == 0:
                            dt_tuple = xlrd.xldate_as_tuple(cell_obj, xl_book.datemode)
                            dt_datetime = datetime(*dt_tuple)
                            cell_obj = dt_datetime.strftime("%m/%d/%Y")
                        new_row.update({col_names[col_idx + 1]:cell_obj})
                        # print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))
                row.append(new_row)

                xl_wells = pd.concat([pd.DataFrame.from_dict(row, orient='columns'), xl_wells], axis=0, sort=False).reset_index(drop=True)

            except:
                print("error reading excel")

    print(xl_wells.columns)
    print(xl_wells.head())
    print(xl_wells.tail())






