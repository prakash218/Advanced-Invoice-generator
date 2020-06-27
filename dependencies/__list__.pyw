from tkinter import *
from tkinter import messagebox
import openpyxl
from xlrd import open_workbook
import runpy
from os import path

global items,item,no,inputval,flag
flag = 0

def setval(n = 5):
    global flag
    flag = 1
    global items,no,item
    try:
        item = int(items.get())
    except:
        messagebox.showinfo("Warning","Please Choose a valid Number")
        return
    print(item)
    no.destroy()
    

def noof():
    global no,items
    no = Tk()
    items = StringVar()
    no.geometry('300x100+100+100')
    no.configure(bg = '#00bfff')
    no.bind('<Return>',setval)
    no.focus_set()
    num = [1,2,3,4,5]
    items.set('none')
    label = Label(no,text = 'Enter the number of items:')
    label.place(x = 10,y = 20)
    entry = OptionMenu(no,items,*num)
    entry.focus_set()
    entry.place(x = 170,y = 20)
    button = Button(no,text = 'submit',bg = '#4d4dff',command = setval)
    button.place(x = 50,y = 50)
    no.mainloop()
    
    
    

def getinp(n = 5):
    global inp1,inp2,inp3,inp4,inp5,item
    to_write=[]
    loc = ("dependencies\_details.xlsx") 
    workb = open_workbook(loc)
    if item >= 1:
        item1 = inp1.get()
        item1 = list(item1.split(','))
        if len(item1) < 4:
            messagebox.showinfo("Warning","Please enter the details correctly");
            return
        to_write.append(item1)
    if item >= 2:
        item2 = inp2.get()
        item2 = list(item2.split(','))
        if len(item2) < 4:
            messagebox.showinfo("Warning","Please enter the details correctly");
            return
        to_write.append(item2)
    if item >= 3:
        item3 = inp3.get()
        item3 = list(item3.split(','))
        if len(item3) < 4:
            messagebox.showinfo("Warning","Please enter the details correctly");
            return
        to_write.append(item3)
    if item >= 4:
        item4 = inp4.get()
        item4 = list(item4.split(','))
        if len(item4) < 4:
            messagebox.showinfo("Warning","Please enter the details correctly");
            return
        to_write.append(item4)
    if item >= 5:
        item5 = inp5.get()
        item5 = list(item5.split(','))
        if len(item5) < 4:
            messagebox.showinfo("Warning","Please enter the details correctly");
            return
        to_write.append(item5)
    invno = 0  
    sheet1 = workb.sheet_by_index(0)
    for i in range(1,10000):
        try:
            val = sheet1.cell_value(i,0)
            try:
                val1 = sheet1.cell_value(i+1,0)
            except:
                invno = i
                break
        except:
            messagebox.showinfo("Warning","No invoice Found...");
            win.destroy()
            basepath = path.dirname(__file__)
            filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
            file_globals = runpy.run_path(filepath)
            
    print(invno,"   ",to_write)
    cmp = str(sheet1.cell_value(invno,4)) 
    var = str(invno)
    if(len(var) == 1):
        var = '00' + str(invno)
    elif (len(var) == 2):
        var = '0'+ str(invno)
    var = var+cmp+'.xlsx'
    xfile = openpyxl.load_workbook(var)
    sheet = xfile.get_sheet_by_name('original')
    if item >= 1:
        sheet['A15'] = '1'
        sheet['B15'] = to_write[0][0]
        sheet['D15'] = to_write[0][1]
        sheet['E15'] = '18%'
        sheet['G15'] = int(to_write[0][2])
        sheet['F15'] = int(to_write[0][3])
        sheet['H15'] = 'no.'
        sheet['I15'] = (int(to_write[0][2]) * int(to_write[0][3]))
    if item >= 2:
        sheet['A17'] = '2'
        sheet['B17'] = to_write[1][0]
        sheet['D17'] = to_write[1][1]
        sheet['E17'] = '18%'
        sheet['G17'] = int(to_write[1][2])
        sheet['F17'] = int(to_write[1][3])
        sheet['H17'] = 'no.'
        sheet['I17'] = (int(to_write[1][2]) * int(to_write[1][3]))
    if item >= 3:
        sheet['A19'] = '3'
        sheet['B19'] = to_write[2][0]
        sheet['D19'] = to_write[2][1]
        sheet['E19'] = '18%'
        sheet['G19'] = int(to_write[2][2])
        sheet['F19'] = int(to_write[2][3])
        sheet['H19'] = 'no.'
        sheet['I19'] = (int(to_write[2][2]) * int(to_write[2][3]))

    if item >= 4:
        sheet['A21'] = '4'
        sheet['B21'] = to_write[3][0]
        sheet['D21'] = to_write[3][1]
        sheet['E21'] = '18%'
        sheet['G21'] = int(to_write[3][2])
        sheet['F21'] = int(to_write[3][3])
        sheet['H21'] = 'no.'
        sheet['I21'] = (int(to_write[3][2]) * int(to_write[3][3]))

    if item >= 5:
        sheet['A23'] = '5'
        sheet['B23'] = to_write[4][0]
        sheet['D23'] = to_write[4][1]
        sheet['E23'] = '18%'
        sheet['G23'] = int(to_write[4][2])
        sheet['F23'] = int(to_write[4][3])
        sheet['H23'] = 'no.'
        sheet['I23'] = (int(to_write[4][2]) * int(to_write[4][3]))

    xfile.save(var)
    win.destroy()
    basepath = path.dirname(__file__)
    filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
    file_globals = runpy.run_path(filepath)  
    
    
    
noof()
if flag:
    win = Tk()
    win.iconbitmap(r'dependencies\icon.ico')
    win.focus_set()
    win.title('Add items')
    win.bind('<Return>',getinp)
    inp1 = StringVar()
    inp2 = StringVar()
    inp3 = StringVar()
    inp4 = StringVar()
    inp5 = StringVar()
    win.geometry('500x300+100+100')
    win.configure(bg = '#00bfff')
    label = Label(win,text = 'Enter the item name,hsn,qty,price',fg = 'red', bg = '#00bfff',font = ('Calibri',14)).grid(row = 1, column = 4)

    if item >= 1:
        label = Label(win,text = 'List Item:').grid(row = 2,column = 1,padx = 5,pady = 10)
        entry = Entry(win,textvariable = inp1)
        entry.grid(row = 2,column = 4,pady = 10,padx = 5)
        entry.focus_set()
    if item >= 2:
        label = Label(win,text = 'List Item:').grid(row = 3,column = 1,pady = 10,padx = 5)
        entry = Entry(win,textvariable = inp2).grid(row = 3,column = 4,pady = 10,padx = 5)
    if item >= 3:
        label = Label(win,text = 'List Item:').grid(row = 4,column = 1,pady = 10,padx = 5)
        entry = Entry(win,textvariable = inp3).grid(row = 4,column = 4,pady = 10,padx = 5)
    if item >= 4:
        label = Label(win,text = 'List Item:').grid(row = 5,column = 1,pady = 10,padx = 5)
        entry = Entry(win,textvariable = inp4).grid(row = 5,column = 4,pady = 10,padx = 5)
    if item >= 5:
        label = Label(win,text = 'List Item:').grid(row = 6,column = 1,pady = 10,padx = 5)
        entry = Entry(win,textvariable = inp5).grid(row = 6,column = 4,pady = 10,padx = 5)
        



    button = Button(win,text = 'Submit',bg = '#4d4dff',command = getinp)
    button.grid(row = 10,column = 5)

    win.mainloop()
