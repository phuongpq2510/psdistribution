# import openpyxl
# def update_excel_with_data(data_dict, excel_filename, sheet_name):
#     # Mở tệp Excel hiện có
#     workbook = openpyxl.load_workbook(excel_filename)

#     # Chọn sheet cụ thể
#     sheet = workbook[sheet_name]

#     # Lấy hàng thứ 2 (dòng 2)
#     row = sheet[2]

#     # Duyệt qua từng ô trong hàng thứ 2 để tìm cột
#     a = {}
#     for cell in row:
#         if cell.value is None:
#             break
#         else:
#             value = cell.value
#             column = cell.column_letter
#             a[value] = column

#     # Bắt đầu từ hàng 3 để bỏ qua hàng tiêu đề
#     i = 3
#     for keys, values in data_dict.items():
#         cot = a['ID']
#         hang = i
#         sheet[cot + str(hang)] = keys
#         for key, value in data_dict[keys].items():
#             o = a[key] + str(i)
#             sheet[o] = value
#         i += 1

#     # Lưu lại tệp Excel
#     workbook.save(excel_filename)

# import openpyxl

# def update_excel_with_data(data_dict, excel_filename):
#     workbook = openpyxl.load_workbook(excel_filename)
#     sheet = workbook['SOURCE']
#     headers = {cell.value: cell.column_letter for cell in sheet[2]}
    
#     for row, data in data_dict.items():
#         sheet[f'{headers["ID"]}{row + 2}'] = row
#         for key, value in data.items():
#             col = headers.get(key)
#             if col:
#                 sheet[f'{col}{row + 2}'] = value
    
#     workbook.save(excel_filename)


# data_dict = {1: {'BUS_ID': 1, 'NAME': 'BUS1', 'kV': 12, 'FLAG': 1, 'CODE': None, 'vGen [pu]': 1.02, 'aGen [deg]': 0, 'Pgen': None, 'Qmax': 99999, 'Qmin': -99999, 'vGen Profile': None, 'pGen Profile': None, 'PLOT': None, 'MEMO': 'đơn vị kva, kw'}, 2: {'BUS_ID': 12, 'NAME': 'BUS12', 'kV': 12, 'FLAG': 1, 'CODE': 0, 'vGen [pu]': 1.02, 'aGen [deg]': 0, 'Pgen': 500, 'Qmax': 99999, 'Qmin': -99999, 'vGen Profile': None, 'pGen Profile': None, 'PLOT': None, 'MEMO': None}}
# excel_filename = 'default.xlsx'

# update_excel_with_data(data_dict, 'default.xlsx')

import shutil
default_file = 'default.xlsx'
new_file = 'abc.xlsx'
shutil.copy(default_file, new_file)