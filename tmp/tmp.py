import openpyxl
import pprint

wb = openpyxl.load_workbook('./アンケート結果.xlsx')
print(type(wb))

print(wb.sheetnames)

sheet = wb['Sheet1']
sheet2 = wb['Sheet2']
cell = sheet['C1']

print(cell.value)


sheet2.cell(row=1,column=1,value='項目')
sheet2.cell(row=1,column=2,value='number')
sheet2.cell(row=1,column=3,value='班')
sheet2.cell(row=1,column=4,value='属性')
sheet2.cell(row=1,column=5,value='結果')
num = 2

print(sheet.max_row)

MaxRow = sheet.max_row + 1
exit()


for i in range(2,MaxRow):
    team = sheet.cell(row=i,column=2).value
#    print(team)
    for j in range(3,15):
        cell = sheet.cell(row=i,column=j).value
        if cell == 1 : 
            cell = '設問' + str(j - 2)
        if j == 13:
            cell = "" # mask it 
        sheet2.cell(row=num,column=1,value=num)
        sheet2.cell(row=num,column=2,value=i-1)
        sheet2.cell(row=num,column=3,value=team)
        sheet2.cell(row=num,column=4,value=j)
        sheet2.cell(row=num,column=5,value=cell)
        num = num + 1


# wb.save('./アンケート結果2.xlsx')
list = [[0 for i in range(13)] for j in range(13)]
PerList = [[0 for i in range(13)] for j in range(13)]

NumTeam = [0]*13
for i in range(2,MaxRow):
    team = sheet.cell(row=i,column=2).value
    if team == "9-1":
        team = 9
    NumTeam[team] = NumTeam[team] + 1
    for j in range(3,15):

        q = j - 2
        cell = sheet.cell(row=i,column=j).value
        if cell == 1 : 
            list[team][q] = list[team][q] + 1

pprint.pprint(NumTeam)
pprint.pprint(list)

for team in range(1,12):
    NumTeam[0] = NumTeam[0] + NumTeam[team]

for q in range(1,11):
    for team in range(1,12):
        list[0][q] = list[0][q] + list[team][q]

for team in range(0,12):
    for q in range(1,11):
        PerList[team][q] = int(list[team][q] / NumTeam[team] *100)

pprint.pprint(PerList)

sheet3 = wb.create_sheet('Sheet3')
title  =["","夏祭り","福利厚生","防災","大掃除","ゴミ","防犯灯","交通標識","金杉会館","警察","渋滞"]
for q in range(1,11):
    sheet3.cell(row=1,column=q+1,value=title[q])

for team in range(0,12):
    for q in range(1,11):
        sheet3.cell(row=team+2, column=q+1,value=PerList[team][q])
        
sheet3.cell(row=2,column=1,value="合計")
for team in range(1,12):
    sheet3.cell(row=team+2,column=1,value=team)


    "=CORREL($B$3:$K$3,$B3:$K3)"

# wb.save('./アンケート結果2.xlsx')
