# python version 3.11.3 64-bit
# openpyxl version 3.1.2
# jinja2 version 3.1.2


import openpyxl
from jinja2 import Template
def excel_row_value_dice(file_name, title):
    wb = openpyxl.load_workbook(file_name)
    ws = wb[title]  
    cols_list = []
    for col in ws.columns:
        col_list = []
        for cell in col:
            col_list.append(cell.value)
        newlist = list(filter(lambda x : x != None, col_list))
        cols_list.append(newlist)
    # print(rows_list)
    result = {}
    for i in range(len(cols_list)):
        col_list = []
        if((len(cols_list[i])-1)>1):
            for j in range(len(cols_list[i])-1):
                col_list.append(cols_list[i][j+1])
            result[cols_list[i][0]]=col_list
        elif((len(cols_list[i])-1)==1):
            result[cols_list[i][0]]=cols_list[i][1]
    # print(reslut)
    return result
file_name, title=r'./data.xlsx','Sheet1'
template_fname=r"./template.tpl"
template = Template(open(template_fname,encoding='utf-8').read())
with open('./template/test.c','w',encoding='utf-8') as f:
    f.write(template.render(excel_row_value_dice(file_name, title)))
