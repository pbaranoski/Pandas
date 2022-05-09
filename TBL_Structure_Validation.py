from sys import builtin_module_names
import pandas as pd
import openpyxl as excel
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import Fill, Color
import os

from pymysql import ROWID

TBL_COLS1_csv = r"C:\Users\user\OneDrive - Apprio, Inc\Documents\Snowflake\Compare\BIA_DEV_INFO.csv"
TBL_COLS2_csv = r"C:\Users\user\OneDrive - Apprio, Inc\Documents\Snowflake\Compare\BIA_TST_INFO.csv"

fDtlDiffs = r"C:\Users\user\OneDrive - Apprio, Inc\Documents\Snowflake\Compare\TBL_COL_Diffs.csv"
fDtlDiffsXLSX = r"C:\Users\user\OneDrive - Apprio, Inc\Documents\Snowflake\Compare\TBL_COL_Diffs.xlsx"

def main():
    #########################################################
    # Define variables
    #########################################################

    # truncate output file if exists
    if os.path.exists(fDtlDiffs):
        os.truncate(fDtlDiffs, 0)

    #########################################################
    # Read CME and IDR csv files into Pandas Data Frames
    #########################################################
    dfFile1 = pd.read_csv(TBL_COLS1_csv, dtype=str, na_filter=False)  
    dfFile1.sort_values(by=['TABLE_SCHEMA', 'TABLE_NAME', 'COLUMN_NAME'], inplace=True)

    #lstCols = dfFile1.columns.values.tolist()
    #print(lstCols)

    dfFile2 = pd.read_csv(TBL_COLS2_csv, dtype=str, na_filter=False)  
    dfFile2.sort_values(by=['TABLE_SCHEMA', 'TABLE_NAME', 'COLUMN_NAME'], inplace=True)

    #print(dfFile1)
    #print(dfFile2)

    # in essence a join on all columns
    comparison_df = dfFile1.merge(dfFile2, indicator=True, how='outer')

    # after merge --> row will be "marked" as 'both" if both data frame rows match
    dfDiff = comparison_df[comparison_df['_merge'] != 'both']

    # sort and write out results    
    dfDiffSorted = dfDiff.sort_values(by=['TABLE_SCHEMA', 'TABLE_NAME', 'COLUMN_NAME'])    
    #dfDiffSorted.to_csv(fDtlDiffs,sep = ",", index=False, line_terminator = "\n")
    
    # saving diffs to xlsx file
    XLSX = pd.ExcelWriter(fDtlDiffsXLSX)
    dfDiffSorted.to_excel(XLSX, index=False)    
    XLSX.save()


    ###\/######
    # Add hilighting of rows
    wrkbk = excel.load_workbook(fDtlDiffsXLSX)

    if wrkbk is None:
        print("couldn't get workbook onject")
    else:
        print("Got the workbook")    

    ###################################################
    # process Excel spreadsheet
    ###################################################
    bFlipColor = False

    sheet = wrkbk.active
    rowColor = "FFFF00"

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):

        if row[0].row == 1:
            pass
        elif (row[0].row) % 2 == 0:
            if bFlipColor == False:
                bFlipColor = True
                rowColor = 'B0FFFF'
            else:
                bFlipColor = False
                rowColor = '0DFFBE'  


        for cell in row:
            if cell.row == 1:
                cell.fill = PatternFill(start_color='66DDFF', end_color='66DDFF', fill_type = "solid")                
            else:
                cell.fill = PatternFill(start_color=rowColor, end_color=rowColor, fill_type = "solid")
                                        
    wrkbk.save(fDtlDiffsXLSX)

  
    exit(0)


if __name__ == "__main__":
    
    main()

