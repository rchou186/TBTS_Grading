###############################################################################
### Printout the result of sheet 1 in excel to printer                      ###
### Module to install: win32com, pypiwin32                                  ###
###############################################################################

import win32com.client
from pywintypes import com_error

# assign active printer here
ACTIVE_PRINTER = "Brother DCP-T510W"
#ACTIVE_PRINTER = "EPSONM1120 (M1120 Series)"
#ACTIVE_PRINTER = "HP LaserJet Pro M148fdw (E1652D)"

def print_result(result_file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        # open
        wb = excel.Workbooks.Open(result_file_path)
        excel.Visible = False
        
        # specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()

        # print to printer
        wb.ActiveSheet.PageSetup.LeftMargin = 20 #28.35
        wb.ActiveSheet.PageSetup.RightMargin = 20 #28.35
        wb.ActiveSheet.PageSetup.TopMargin = 20 #53.863
        wb.ActiveSheet.PageSetup.BottomMargin = 20 #53.863
        wb.ActiveSheet.PageSetup.HeaderMargin = 20 #22.68
        wb.ActiveSheet.PageSetup.FooterMargin = 20 #22.68
        wb.ActiveSheet.PageSetup.Orientation = 2
        wb.ActiveSheet.PageSetup.CenterHorizontally = 1
        wb.ActiveSheet.PageSetup.CenterVertically = 1
        wb.ActiveSheet.PageSetup.Zoom = False
        wb.ActiveSheet.PageSetup.FitToPagesWide = 1
        wb.ActiveSheet.PageSetup.FitToPagesTall = 1

        # printout now
        wb.ActiveSheet.PrintOut(1,1,1,False,ACTIVE_PRINTER)

    except com_error as e:
        print("Printout to printer failed.")
    else:
        print("Printout to printer succeeded.")
    finally:

        # close workbook
        wb.Close(False)     # close workbook without save by False parameter

        # end excel program
        # excel.Quit()

if __name__ == "__main__":
    # Path to original excel file
    WB_PATH = r'c:/Python_Prog/TBTSGrading/J22T00299-(20+20).xlsx'

    print_result(WB_PATH)
