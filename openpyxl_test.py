# import openpyxl
# from openpyxl import load_workbook
# # 修改目前工作表
# from openpyxl.utils import get_column_letter, column_index_from_string

# ##-------------------------- openpyxl----------------------------##
# # 網址:https://hackmd.io/@amostsai/SJkC1_EcX?type=view


# fn = 'test.xlsx'
# # 預設可讀寫，若有需要可以指定write_only和read_only為True
# wb = load_workbook(fn)


# # 顯示工作表名稱-方法1 
# print("法1",wb.sheetnames)


# # 顯示工作表名稱-方法2   (以 for 迴圈逐一處理每張工作表)
# for sheet in wb:
#     print("法2",sheet.title)


# # 取得目前作用中的工作表(存檔後正在使用的工作表)
# print("目前使用工作表",wb.active.title)


# # 目前在第幾個工作表,索引值從0開始
# wb.active = 1
# print('excel活動工作表： ', wb.active)


# #將目前所在位置 改成 名為"台積電"的工作表 這個位置(並非改名 只是變換目前所在位置)
# wb.active = wb['台積電']
# print('目前工作表： ', wb.active.title)


# # 透過名稱取得工作表
# sheet = wb['工作表2']
# # 更改工作表名稱
# sheet.title = "台積電"
# # 更改工作表標籤顏色
# sheet.sheet_properties.tabColor = "1072BA"


# # 複製一個現在有的工作表(ex:台積電) 將其命名為new
# sheet = wb['台積電']
# target = wb.copy_worksheet(sheet)
# target.title = 'new'


# #刪除工作表(new1)
# sheet = wb['new1']
# wb.remove(sheet)
# wb.save(fn)


# # 讀取工作表所有內容
# wb.active = 0
# ws = wb.active
# print('excel活動工作表：', ws)
# # 欄（column）：工作表中的直欄，以A、B、C…代表
# # 行（row）：工作表中的橫列，以1、2、3…代表
# # 儲存格（cell）：工作表中的每一個都是一個儲存格
# for row in ws:
#     for cell in row:
#         print(cell.value)

# print()
# print('D1內容： ', ws['D1'].value)

# range = ws['A1': 'B20']

# for a, b in range:
#     print("{0} {1}".format(a.value, b.value))
# #------------------------------------------不懂------------------------------------
# print()
# for a, b, c in ws[ws.dimensions]:
#     print(a.value, b.value, c.value)
# #------------------------------------------不懂------------------------------------


# #寫入儲存格
# print('P2內容： ', ws['P2'].value)
# ws['P2'].value =  20
# ws.cell(column=16, row=3).value = 999

# # 新增工作表（把新增的工作表放在最後方）
# ws1 = wb.create_sheet("新增工作表1")

# # 儲存檔案(必須把excel檔關掉始可執行)
# wb.save(fn)

# for title in bsObj.find("table", {"id" : "CPHB1_bt1_g"}).find("tr"):
#     	a += title.string;
# 	#寫入儲存格
# 	ws.cell(column=1, row=1).value = title.string
# print(a)
# #