# -*- coding: utf-8 -*-
"""
General XLSX to PDF scripts
Erik Winterling
15.10.2020
without tqdm -> currently buggy as .exe file

Note: COM can use VBA codes ,
https://docs.microsoft.com/de-de/office/vba/api/excel.pagesetup.fittopageswide
If the Zoom property is True, the FitToPagesWide property is ignored.
working: ws.PageSetup.
PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
"""


import os
from win32com import client
from datetime import datetime


searchpath= os.getcwd()
output_folder=searchpath+r"/export"
output_file=output_folder+r"test"

file_list=[]
if os.path.exists(output_folder) != True:
    os.mkdir(output_folder)
    
def createFileList(search_path):
    for (path, dirs, files) in os.walk(search_path):
        for entry in files:
            if (".xlsx" in entry and "~" not in entry):
    #            print(os.path.join(path,entry))
    #            print(entry.split(".")[0])
                name, ext = os.path.splitext(entry)
                file_data={
                        "xlsx_path":os.path.join(path,entry),
                        "filename":name,
                        }                
                file_list.append(file_data)



def print_file(file_to_print,output):
    excel= client.DispatchEx("Excel.Application")
    excel.Visible=0
    wb=excel.Workbooks.Open(file_to_print)
    try:
        #print(len(wb.sheets))
        for sh in wb.Sheets:             
            excel.Worksheets(sh.Name).Activate()
            ws=excel.ActiveSheet
            cvalue= ws.Cells(1,2).Value
            # print("sheetname = {}".format(sh.name))
            # print("cellvalue={}".format(cvalue))
            # print(type(cvalue))
            if cvalue=="Laborjournal":
                #print("true")
                wb.SaveAs(output+sh.name, FileFormat=57)
            else:
                pass
            #wb.SaveAs(output+sh.name, FileFormat=57)
             #  if ws.Cell(1,2).Value == "Laborjournal":
             #     try:
             #         print("generate PDF")
             #         wb.SaveAs(output+sh.name, FileFormat=57)
             #     except Exception as error:
             #         print("error in if part")
             #         print(str(error))
             # else:
             #      pass
    except Exception as e:
        print("failed to convert")
        print (str(e))
    finally:
        wb.Close()
        excel.Quit()
    
if __name__== "__main__":
    start=datetime.now()
    createFileList(searchpath)
    for file in file_list:
        try:
            print_file(file.get("xlsx_path"), output_folder+"\{}".format(file.get("filename")))
        except Exception as e:
            print("error occured with file {}!".format(file.get("filename")))
            print(str(e))
        finally:
            excel = None
            wb = None
            ws = None
            xl = None
    print("done after {}".format(datetime.now()-start))
    input("press any key to exit...")
#print_file(r"C:\Users\winte\Desktop\Test\EW347.xlsx",r"C:\Users\winte\Desktop\Test\EW347")


# =============================================================================
# Quelle https://stackoverflow.com/questions/41407824/how-can-i-iterate-over-worksheets-in-win32com
# Iterieren über worksheets in aktuellen wb:
 # for sh in wb.Sheets:
 #    excel.Worksheets(sh.Name).Activate()
 #    print(sh.name)
 #    ws=excel.ActiveSheet
 #    print(ws.Cells(1,2).Value)
#
# =============================================================================
