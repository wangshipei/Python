from openpyxl import load_workbook
import glob
import os
import win32com.client
from tqdm import tqdm

excel_dir = r'C:\Users\24910\PycharmProjects\excel2pdf\excel'
os.chdir(excel_dir)
file_list = glob.glob('*.xlsx')
for file in tqdm(file_list, position=0, leave=True, desc=f'正在处理Excel文件'):
    wb = load_workbook(file, data_only=True)
    ws1 = wb['INV.']
    ws2 = wb['PL.']
    del wb['清单明细表']
    del wb['报关要素']
    del wb['产地证数据']
    del wb['提单数据']
    del wb['SWOD']
    ws1.delete_cols(14, 4)
    ws2.delete_cols(15, 13)
    wb.save(file)

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False

pdf_dir = r'C:\Users\24910\PycharmProjects\excel2pdf\pdf'
os.chdir(excel_dir)
file_list = glob.glob('*.xlsx')
for file in tqdm(file_list, position=0, leave=True, desc=f'正在把Excel文件转换成PDF'):
    f_name, f_ext = os.path.splitext(file)
    wb_path = excel_dir + '\\' + file
    wb = o.Workbooks.Open(wb_path)
    ws_index_list = [1, 2]
    path_to_pdf = pdf_dir + '\\' + f_name + '.pdf'
    wb.WorkSheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
    wb.Close(True)

print('所有数据处理完成')
