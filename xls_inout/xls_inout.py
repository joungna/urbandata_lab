import openpyxl

# 엑셀 파일 열기
wb = openpyxl.load_workbook('2017년 광고비 - 삼성전자.xlsx')
wb

# 모든 시트 이름들 얻기
wb.get_sheet_names()

# 시트 이름으로 시트 얻기
sheet = wb.get_sheet_by_name('Sheet1')
sheet

# 활성화 시트 얻기
sheet = wb.active
sheet

# 셀에 접근
sheet['A2'].value

sheet['B1'].value

# cell(row = n, column = m) 

sheet.cell(row = 1, column = 3).value

# 범위 접근
muti_cells = sheet['E2':'F14']
muti_cells

muti_cells = sheet['E2':'F13']
for row in muti_cells:
    print(row[0].value, row[1].value)

    # 모든 row 살펴보기
for row in sheet.rows:
    print([col.value for col in row])

# 모든 column 살펴보기
for col in sheet.columns:
    print([r.value for r in col])

# 엑셀로 저장하기
wb.save("sheet.xlsx")    

# 파일명 바꾸기
import os
os.rename('sheet.xlsx', '삼성전자랭킹.xlsx')