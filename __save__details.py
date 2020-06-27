import glob
import openpyxl
from xlrd import open_workbook
from warnings import filterwarnings
filterwarnings("ignore", category=DeprecationWarning)
while(1):
	try:
		start = int(input("Enter the Starting invoice number:"))
		break
	except:
		continue
while(1):
	try:
		end = int(input("Enter the Ending invoice number:"))
		break
	except:
		continue
for i in range(start ,end+1):
	if i < 10:
		a = '00' + str(i)
	elif i < 100:
		a = '0' + str(i)
	else:
		a = str(i)
	b = a
	a += '*' * 20
	for fpath in glob.glob(a):
		amt = 0
		print (fpath)
		comp = fpath.split('.')
		comp = str(comp[0])
		comp = comp.split(b)
		comp = str(comp[1])
		
		workb = open_workbook(fpath)
		sheet1 = workb.sheet_by_index(0)
		
		inv = int(sheet1.cell_value(4,4))
		date = sheet1.cell_value(4,6)
		podate = sheet1.cell_value(6,6)
		ponum = sheet1.cell_value(6,4)
		for i in range(14,25):
			try:
				amt += float(sheet1.cell_value(i,8))
			except:
				continue
		
		print(inv)
		print(date)
		print(podate)
		print(ponum)
		print(comp)
		print(amt)
		
		details = [inv,date,podate,ponum,comp,amt]

    #---------------------saving details.xlsx---------
		new = openpyxl.load_workbook('dependencies\_details.xlsx')
		first = new.get_sheet_by_name('Details')
		list1 = ['A','B','C','D','E','F']
		for i in range(6):
		    to_write = list1[i] + str(inv+1)
		    #print(to_write , details[i])
		    first[to_write] = details[i]
		new.save('dependencies\_details.xlsx')


