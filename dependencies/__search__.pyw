from tkinter import *
from xlrd import open_workbook
import sys
from tkinter import messagebox
from os import path
import runpy


blue = "#4d4dff"
lblue = '#00bfff'

global root

def submit(num = 5):
	# scrollbar = Scrollbar(search)
	# scrollbar.grid(row = 0, column = 1,stick = 'ns')
	# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
	global root
	try:
		root.destroy()
	except:
		a = 1
	root = Tk()
	root.geometry('600x400+100+100')

	text = ''
	inno = "Invoice Number:"
	indate = "Invoce Date:"
	datepo = "PO Date:"
	po = "PO Number:"
	cmpy = "Company:"
	amount = "Amount:"
	recd = "Recieved:"

	s = searchfor.get()
	if len(s) == 0:
		messagebox.showinfo("Inavlid","Please enter a valid Search term");
		return 
	opt = choice.get()

	loc = ("dependencies\\_details.xlsx")
	workb = open_workbook(loc)
	sheet1 = workb.sheet_by_index(0)
	rcd = ''
	if opt == 'invoice number':
		try:
			s = int(s)
		except:
			messagebox.showinfo("Inavlid","Please enter a valid Invoice Number");
			return

		for i in range(10000):
			try:
				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				comp = sheet1.cell_value(i,4)
				amt = sheet1.cell_value(i,5) + sheet1.cell_value(i,5) * 0.18
				try:
					rcd = sheet1.cell_value(i,8)
				except IndexError:
					rcd ="no"
				if inv == int(s):
					amt = 'Rs.'+str(round(amt,2))
					if rcd == '':
						rcd = 'no'
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv))
					indate += '\n' + str(date)
					datepo += '\n' + str(podate)
					po += '\n' +str(ponum)
					cmpy += '\n' + str(comp)
					amount += '\n'+(amt)
					recd += '\n' + str(rcd)
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					break
			except Exception as e:
				print(e,i)
				continue
	
	if opt == 'company name':
		s = s.lower()
		for i in range(10000):
			try:
				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				comp = sheet1.cell_value(i,4)
				amt = sheet1.cell_value(i,5) + sheet1.cell_value(i,5) * 0.18
				try:
					rcd = sheet1.cell_value(i,8)
				except IndexError:
					rcd ="no"
				#if found add the details
				if s in comp.lower():
					amt = 'Rs.'+str(round(amt,2))
					if rcd == '':
						rcd = 'no'
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv))
					indate += '\n' + str(date)
					datepo += '\n' + str(podate)
					po += '\n' +str(ponum)
					cmpy += '\n' + str(comp)
					amount += '\n'+str(amt)
					recd += '\n' + str(rcd)
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					
			except Exception as e:
				print(e,i)
				continue
	if opt == 'PO number':
		for i in range(10000):
			try:
				inv = sheet1.cell_value(i,0)
				date = sheet1.cell_value(i,1)
				podate = sheet1.cell_value(i,2)
				ponum = sheet1.cell_value(i,3)
				comp = sheet1.cell_value(i,4)
				amt = sheet1.cell_value(i,5) + sheet1.cell_value(i,5) * 0.18
				try:
					rcd = sheet1.cell_value(i,8)
				except IndexError:
					rcd ="no"
				if ponum == s:
					amt = 'Rs.'+str(round(amt,2))
					if rcd == '':
						rcd = 'no'
					# text = 'invoice number'+ '\t' +'invoice date'+ '\t' +'PO Date'+ '\t' +'PO Number'+ '\t' +'Company name'+ '\t' +'Amount'+ '\t' +'Received'
					inno += '\n' + str(int(inv))
					indate += '\n' + str(date)
					datepo += '\n' + str(podate)
					po += '\n' +str(ponum)
					cmpy += '\n' + str(comp)
					amount += '\n'+ amt
					recd += '\n' + str(rcd)
					print(str(int(inv)) + '\t' + str(date) + '\t' + str(podate) + '\t' + str(ponum) + '\t' + str(comp) + '\t' +str(amt) + '\t' +str(rcd))
					break
			except Exception as e:
				print(e,i)
				continue
	canvas = Canvas(root,bg = lblue)
	scroll_y = Scrollbar(root, orient="vertical", command=canvas.yview)
	frame1 = Frame(canvas,relief = RAISED, bg = lblue)
	label = Label(frame1,text = inno,bg = lblue)
	label.grid(row = 5,column =1)
	label = Label(frame1,text = indate,bg = lblue)
	label.grid(row = 5,column =2)
	label = Label(frame1,text = datepo,bg = lblue)
	label.grid(row = 5,column =3)
	label = Label(frame1,text = po,bg = lblue)
	label.grid(row = 5,column =4)
	label = Label(frame1,text = cmpy,bg = lblue)
	label.grid(row = 5,column =5)
	label = Label(frame1,text = amount,bg = lblue)
	label.grid(row = 5,column =6)
	label = Label(frame1,text = recd,bg = lblue)
	label.grid(row = 5,column =7)
	canvas.create_window(0, 0, anchor='nw', window=frame1)
# make sure everything is displayed before configuring the scrollregion
	canvas.update_idletasks()
	canvas.configure(scrollregion=canvas.bbox('all'), yscrollcommand=scroll_y.set)
	canvas.pack(fill='both', expand=True, side='left')
	scroll_y.pack(fill='y', side='right')


def close():
	search.destroy()
	basepath = path.dirname(__file__)
	filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
	file_globals = runpy.run_path(filepath)


search = Tk()
search.title("Search")
searchfor = StringVar()
choice = StringVar()
search.iconbitmap(r'dependencies\icon.ico')
search.geometry('350x250+100+100')
search.bind('<Return>',submit)
search.configure(bg = blue)
Label(text = "KKR Engineering",bg = blue,fg = 'white',font = ("Times New Roman CE",16)).grid(row = 1,column = 8)



options = ["invoice number","company name","PO number"]
choice.set(options[0])
frame = Frame(search,bg = lblue,relief = RAISED,bd = 10)
frame.grid(row = 3 ,padx = 25,column = 8)

entry = Entry(frame,textvariable = searchfor).grid(row = 1, column = 2,padx = 10)

label = Label(frame,text = "Search:").grid(row = 1,column = 1,padx = 25)

label = Label(frame,text = "Search By:").grid(row = 2,column = 1,padx = 25,pady = 10)

menu = OptionMenu(frame,choice,*options).grid(row = 2, column = 2, padx = 25, pady = 10)

submitbutton = Button(frame,text = "Search", bg = blue,activebackground = lblue, command = submit).grid(row = 3, column = 1,padx = 10)

cancel = Button(frame,text = "Cancel", bg = blue,activebackground = lblue, command = close).grid(row = 3, column = 2,padx = 10)


search.mainloop()