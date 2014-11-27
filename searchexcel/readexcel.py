# To change this license header, choose License Headers in Project Properties.
# To change this template file, choose Tools | Templates
# and open the template in the editor.
# To change this license header, choose License Headers in Project Properties.
# To change this template file, choose Tools | Templates
# and open the template in the editor.
import xlrd
data = xlrd.open_workbook('1.xlsx')
searchType = input("searchType:")
if searchType == 'pet':
	table = data.sheets()[0]
	nrows = table.nrows
	j =[]
	searchWord = input("searchWord:")
	for i in range(nrows):
		if searchWord in table.cell(i,1).value:
			j.append(i)
	for h in j:
		print(int(table.cell(h,0).value),table.cell(h,1).value)
elif searchType =='item':
	table = data.sheets()[1]
	nrows = table.nrows
	j =[]
	searchWord = input("searchWord:")
	for i in range(nrows):
		if searchWord in table.cell(i,1).value:
			j.append(i)
	for h in j:
		print(int(table.cell(h,0).value),table.cell(h,1).value)
elif searchType =='equip':
	table = data.sheets()[2]
	nrows = table.nrows
	j =[]
	searchWord = input("searchWord:")
	for i in range(nrows):
		if searchWord in table.cell(i,1).value:
			j.append(i)
	for h in j:
		print(int(table.cell(h,0).value),table.cell(h,1).value)
input("Please input<Enter>")

