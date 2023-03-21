# # a = 'D:\excel\\SUPERVISION2327.xls'
# # b = a[0:-4]
# # print(b+'.xls')
# import glob
# import string
#
# from openpyxl.reader.excel import load_workbook
# import pyexcel as p
# #
# # # ------RETRIEVE EXCEL FILE Paths------
# # # gets the excel file paths recursively and saves in a list
# #excel_files = glob.glob('D:/Life/*.xlsx', recursive=True)
# # print(excel_files)
# # print(len(excel_files))
# #value = []
# # # for file in excel_files:
# # #     if file[-3:] == 'xls':
# # #         filename = file[9:]
# # #         p.save_book_as(file_name='filename', dest_file_name='Updated.xlsx')
# # #     wb = load_workbook(file, data_only=True)
# # # #-----RETRIEVE DATA IN SPECIFIC CELL----
# # for excel_file in excel_files:
# #     wb = load_workbook(excel_file, data_only=True)  # load workbook and cell values as data not formula
# #     wb.active = wb[('SDR8i')] #sets active sheet to SDR10iv
# #     sheet_obj = wb.active
# #     cell_obj = sheet_obj.cell(row=16, column=5)
# #     value.append(cell_obj.value)
# # print(value)
#
# test_list =[]
# test_list = list(string.ascii_uppercase)
# print(test_list)
# b=[]
# print(len(b))
#
# sheet_now.cell(row=i - 1, column=2).value = sheet_obj.cell(row=2, column=8).value
# sheet_now.cell(row=i, column=1).value = sheet_obj.cell(row=indice, column=7).value
# sheet_now.cell(row=i, column=2).value = b
# sheet_now.cell(row=i, column=3).value = c
# sheet_now.cell(row=i - 1, column=6).value = sheet_obj.cell(row=3, column=8).value
# sheet_now.cell(row=i, column=6).value = b1
# sheet_now.cell(row=i, column=7).value = c1
# sheet_now.cell(row=i - 1, column=10).value = sheet_obj.cell(row=4, column=8).value
# sheet_now.cell(row=i, column=10).value = b2
# sheet_now.cell(row=i, column=11).value = c2
# sheet_now.cell(row=i - 1, column=14).value = sheet_obj.cell(row=5, column=8).value
# sheet_now.cell(row=i, column=14).value = b3
# sheet_now.cell(row=i, column=15).value = c3
# sheet_now.cell(row=i - 1, column=18).value = sheet_obj.cell(row=6, column=8).value
# sheet_now.cell(row=i, column=18).value = b4
# sheet_now.cell(row=i, column=19).value = c4

a=4^2
print(a)