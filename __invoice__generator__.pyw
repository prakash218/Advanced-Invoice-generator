import time
t1 = time.time()
from tkinter import *
from tkinter import messagebox
from tkcalendar import Calendar,DateEntry
import shutil
from time import localtime,strftime
from xlrd import open_workbook
from num2words import num2words
import openpyxl
from warnings import filterwarnings
import runpy


blue = "#4d4dff"
lblue = '#00bfff'

named_tuple = localtime() # get struct_time
time_string = strftime("%d/%m/%Y", named_tuple)
global amount,listitem,win,inp

def Search():
    screen.destroy()
    file_globals = runpy.run_path('dependencies\\__search__.pyw') #to get items for invoice
    

def getitem(num):
    screen.destroy()
    file_globals = runpy.run_path('dependencies\\__list__.pyw') #to get items for invoice
    

def delete():
    screen.destroy()
    file_globals = runpy.run_path('dependencies\__delete__.pyw') # to delete an invoice

    
def total():
    global amount
    try:
        amount.destroy()
    except:
        a = 1
    amount = Tk()
    loc = ("dependencies\_details.xlsx") 
    workb = open_workbook(loc) 
    sheet2 = workb.sheet_by_index(0)
    received = 0
    pending = 0
    for i in range(1,1000):
        try:
            val = sheet2.cell_value(i,8)
        except:
            val = "no"
        if val == 'yes':
            try:
                received += (sheet2.cell_value(i,5) + (sheet2.cell_value(i,5))*0.18)
            except:
                received += 0
        else:
            try:
                pending += (sheet2.cell_value(i,5) + (sheet2.cell_value(i,5))*0.18)
            except:
                pending += 0
    received1 = 'Rs. ' + str(int(received))
    pending1 = 'Rs. ' + str(int(pending))
    
    amount.title('Amount')
    amount.geometry('180x200+700+100')
    amount.resizable(False,False)
    amount.configure(bg = blue )
    amount.iconbitmap(r'dependencies\icon.ico')
    label = Label(amount,text = "KKR Engineering",font = ('Calibri',15),bg = blue,fg = 'white')
    label.place(x = 20,y = 10)
    frame2 = Frame(amount,bg = '#00bfff',relief = RAISED , bd = 10)
    frame2.place(x = 10,y= 40)
    label = Label(frame2,text = "Received").grid(row = 1,column = 1,padx = 10,pady = 10)
    label = Label (frame2,text = received1).grid(row = 1,column = 4,padx = 10,pady = 10)
    label = Label (frame2,text = "Pending").grid(row = 2,column = 1,padx = 10,pady = 10)
    label = Label  (frame2,text = pending1).grid(row = 2,column = 4,padx = 10,pady = 10)
    label = Label(frame2)
    button = Button(amount,text = "Close",bg = '#00bfff',activebackground = 'blue',command = amount.destroy)
    button.place(x = 70,y = 160)

def edit():
    screen.destroy()
    file_globals = runpy.run_path('dependencies\__edit__address__.pyw')

def yesorno():
    screen.destroy()
    file_globals = runpy.run_path('dependencies\__yes__or__no__.pyw')
    

def new():
    screen.destroy()
    file_globals = runpy.run_path('dependencies\__new__entry.pyw')

def check(inv):
    loc = ("dependencies\_details.xlsx") 
    workb = open_workbook(loc) 
    sheet1 = workb.sheet_by_index(0)
    try:
        val = sheet1.cell_value(inv,0)
        if( val == inv):
            return 0
        return 1
    except IndexError:
        return 1

def submit1(n = 5):
    global inp
    filterwarnings("ignore", category=DeprecationWarning)
    try:
        in_no = int(invoice_number.get())
    except:
        messagebox.showinfo("Invoice number","Please enter a valid Invoice number");
        return
    if not check(in_no):
        messagebox.showinfo("Invoice number","Invoice already found");
        return
    if in_no <= 0:
        messagebox.showinfo("Invoice number","Please enter a valid Invoice number");
        return
    try:
        cmp = company.get()
        if cmp == "None":
            messagebox.showinfo("Company name","Please choose a valid company");
            return
    except:
        messagebox.showinfo("Company name","Please enter a valid Company Name");
        return
    
    invName = cmp
    po_num = po_number.get()
    
    if len(po_num) == 0:
        messagebox.showinfo("PO number","Please enter a valid PO Number");
        return
    podate = po_date.get()

    try:
        invAmt = int(invoice_amount.get())
    except:
        messagebox.showinfo("Invoice Amount","Please enter a valid Invoice Amount");
        return

    
    
    
    #------------------path creation-----------
    var = str(in_no)
    if(len(var) == 1):
        var = '00' + str(in_no)
    elif (len(var) == 2):
        var = '0'+ str(in_no)
    var = var+cmp+'.xlsx'

    #---------------number to word-----------------------
    tax = invAmt * 0.18
    tax = int(tax)
    tax_word = num2words(tax,lang = 'en_IN')
    total = invAmt + tax
    total = int(total)
    total_word = num2words(total,lang = 'en_IN')
    

    #-----------------creating new invoice-----------
    original = r'dependencies\_test_invoice.xlsx'
    target = var
    shutil.copyfile(original, target)

    #--------------------filling details in invoice----------

    loc = ("dependencies\__companies.xlsx")
    workb = open_workbook(loc)
    sheet1 = workb.sheet_by_index(0)
    write = []
    
    for i in range(10000):
        try:
            name = sheet1.cell_value(i,0)
            if invName == name:
                invName = name
                write.append(sheet1.cell_value(i,1))
                write.append(sheet1.cell_value(i,2))
                write.append(sheet1.cell_value(i,3))
                write.append(sheet1.cell_value(i,4))
                write.append(sheet1.cell_value(i,5))
                break
        except:
            return
            break


    xfile = openpyxl.load_workbook(var)
    sheet = xfile.get_sheet_by_name('original')
    sheet['E5'] = in_no
    sheet['E7'] = po_num
    sheet['G5'] = time_string
    sheet['G7'] = podate
    sheet['A32']= 'INR '+total_word+' ONLY'
    sheet['A38']= 'INR '+tax_word+' ONLY'
    sheet['A9'] = write[0]
    sheet['A10']= write[1]
    sheet['A11']= write[2]
    sheet['A12']= write[3]			
    sheet['A13']= write[4]

    xfile.save(var)

    details = [in_no,time_string,podate,po_num,invName,invAmt]

    #---------------------saving details.xlsx---------
    new = openpyxl.load_workbook('dependencies\_details.xlsx')
    first = new.get_sheet_by_name('Details')
    list = ['A','B','C','D','E','F']
    for i in range(6):
        to_write = list[i] + str(in_no+1)
        #print(to_write , details[i])
        first[to_write] = details[i]
    new.save('dependencies\_details.xlsx')
    
    getitem(5)
    #----------------------------close button------------------
    
def cancel():
    global amount
    MsgBox = messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
        try:
            amount.destroy()
            screen.destroy()
        except:
            screen.destroy()


#--------------screen------------------
try:
    screen.destroy()
except:
    a = 1
screen = Tk()
screen.title("Invoice generator ")
screen.configure(bg = blue)
screen.iconbitmap(r'dependencies\icon.ico')
screen.geometry('600x400+100+100')
screen.resizable(False,False)
screen.bind('<Return>',submit1)
screen.focus_set()
screen.protocol("WM_DELETE_WINDOW", cancel)
Label(text = "KKR Engineering",bg = blue,fg = 'white',font = ("Times New Roman CE",16)).grid(row = 1)
hi = Label(text = "Invoice generator",bg = blue, fg = 'black',font = ("calibri",15)).grid(row = 2)
#label(text="").pack()
#--------------------------------------


#---------------variables-------------
loc = ("dependencies\__companies.xlsx") 
workb = open_workbook(loc) 
sheet1 = workb.sheet_by_index(0)

companies = []

for i in range(1000):
    try:
        val = sheet1.cell_value(i,0)
        if len(val) > 0:
            companies.append(val)
        else:
            break
    except IndexError:
        break

if len(companies) == 0:
    companies.append("No companies found..")


loc = ("dependencies\_details.xlsx") 
workb = open_workbook(loc) 
sheet1 = workb.sheet_by_index(0)
for i in range(1,1000):
    try:
        val = sheet1.cell_value(i,0)
        if val == i:
            continue
        else:
            break
    except IndexError:
        break



invoice_number = StringVar()
invoice_number.set(i)
invoice_date = StringVar()
po_number = StringVar()
po_date = StringVar()
invoice_amount = StringVar()
company = StringVar()
company.set(companies[0])

#--------------------details--------------------
frame = Frame(relief = RAISED,bd = 10, bg = lblue)
frame.grid(row = 5,padx = 10,pady = 10)


invoice = Label(frame,text = " Invoice Number: ").grid(row = 1,column = 1,pady = 5,padx = 5,sticky = W)
entry1 = Entry(frame,textvariable = invoice_number)
entry1.grid(row = 1,column = 4,pady = 5,sticky = W)


invoice = Label(frame,text = "Company Name:").grid(row = 2,column = 1,pady = 5,padx = 5,sticky = W)
entry = OptionMenu(frame,company,*companies)
entry.grid(row = 2,column = 4,pady = 5,padx = 5,sticky = W)


invoice = Label(frame,text = "      PO Number:  ").grid(row = 3,column = 1,pady = 5,padx = 5,sticky = W)
entry1 = Entry(frame,textvariable = po_number)
entry1.grid(row = 3,column = 4,pady = 5,sticky = W)

entry1.focus_set()

invoice = Label(frame,text = "          PO Date:    ").grid(row = 4,column = 1,pady = 5,padx = 5,sticky = W)
cal = DateEntry(frame,locale = 'en_IN',date_pattern = "dd/mm/yyyy",width=30,bg="#4d4dff",fg="black",textvariable = po_date)
cal.grid(row = 4,column = 4,pady = 5,sticky = W)


invoice = Label(frame,text = " Invoice Amount:").grid(row = 5,column = 1,pady = 5,padx = 5,sticky = W)
entry1 = Entry(frame,textvariable = invoice_amount).grid(row = 5,column = 4,pady = 5,sticky = W)

#-------------buttons-----------

newButton = Button(frame, text="Add New Company",background = blue,activebackground = 'blue',fg = 'white',command = new,cursor = 'hand2')
newButton.grid(row = 10,column = 1,padx = 10,pady = 10,sticky = S)

newButton = Button(frame, text="Add Received Payment",background = blue,activebackground = 'blue',fg = 'white',command = yesorno,cursor = 'hand2')
newButton.grid(row = 10,column = 4,padx = 10,pady = 10,sticky = S)

cancelButton = Button(frame, text="Cancel",background = blue,activebackground = 'blue',fg = 'white',command = cancel,cursor = 'hand2')
cancelButton.grid(row = 11,column = 6,padx = 10,pady = 10,sticky = S)

okButton = Button(frame, text="Edit Address",command = edit,background = blue,activebackground = 'blue',fg = 'white',cursor = 'hand2')
okButton.grid(row = 10 ,column = 5,padx = 10,pady = 10,sticky = S)

okButton = Button(frame, text="Submit",command = submit1,background = blue,activebackground = 'blue',fg = 'white',cursor = 'hand2')
okButton.grid(row = 11 ,column = 5,pady = 10,padx = 10,sticky = S)

totalButton = Button(frame,text = "Total and Pending", command = total,background = blue,activebackground = 'blue',fg = 'white',cursor = 'hand2')
totalButton.grid(row = 10,column = 6,pady = 10, padx = 10,sticky = S)

delete = Button(frame,text = "Delete an Invoice", command = delete,background = blue,activebackground = 'blue',fg = 'white',cursor = 'hand2')
delete.grid(row = 11,column = 1,pady = 10, padx = 10,sticky = S)

search = Button(frame,text = "Search details" ,command = Search,background = blue,activebackground = 'blue',fg = 'white',cursor = 'hand2')
search.grid(row = 11,column = 4,pady = 10, padx = 10,sticky = S)

#Label(frame,text = "",bg = "#00bfff").grid(row = 11,column = 1,ipady = 10)
Label(text = "",bg = "#4d4dff").grid(row = 11,column = 1,ipady = 10)


#-----------------------------
t2 = time.time()
print('time:%.2f'%(t2-t1))


screen.mainloop()
