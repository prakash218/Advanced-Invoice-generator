from tkinter import *
from tkinter import messagebox
import sys
from xlrd import open_workbook
import openpyxl
import warnings
from os import path
from os import remove
import runpy

global passw,attempt,flag,password1,inv_no,root

flag = 1
attempt = 0

def main():
    global root
    global inv_no
    root = Tk()
    root.geometry('366x450+100+100')
    root.focus_set()
    root.iconbitmap(r'dependencies\icon.ico')
    root.title('Delete an Invoice')
    root.configure(bg = '#00bfff')
    root.bind('<Return>',check)

    inv_no = StringVar()
    label = Label(root,text = "KKR Engineering",bg = '#00bfff',fg = 'white',font= ("calibri",16)).place(x = 102,y = 20)

    label = Label(root,text = "Invoice Number:",fg = 'black').place(x = 10,y = 100)
    entry = Entry(root,textvariable = inv_no)
    entry.place(x = 160,y = 100)
    entry.focus_set()


    submit = Button(root,text = "Search",command = check,background = '#4d4dff',activebackground = 'blue').place(x = 230,y = 400)
    cancel = Button(root,text = "Cancel",command = close,background = '#4d4dff',activebackground = 'blue').place(x = 290,y = 400)


    root.mainloop()
def enter(n=5):
    global password1
    global attempt
    check = passw.get()
    if (check == 'kkr2008'):
        password1.destroy()
        main() 
    else:
        attempt += 1
        psw.delete(0,END)
    if attempt >= 3:
        password1.destroy()
        basepath = path.dirname(__file__)
        filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
        file_globals = runpy.run_path(filepath)

    

def close2():
    password1.destroy()
    basepath = path.dirname(__file__)
    filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
    file_globals = runpy.run_path(filepath)

def save(write,cmp):
    global root
    in_no = int(write)-1
    Label(text = 'YES                                     ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)
    new = openpyxl.load_workbook('dependencies\_details.xlsx')
    first = new.get_sheet_by_name('Details')
    
    var = str(in_no)
    if(len(var) == 1):
        var = '00' + str(in_no)
    elif (len(var) == 2):
        var = '0'+ str(in_no)
    var = var+cmp+'.xlsx'
    
    list = ['A','B','C','D','E','F','I']
    
    for i in range(7):
        to_write = list[i] + write
        #print(to_write , details[i])
        first[to_write] = ''
        
    new.save('dependencies\_details.xlsx')
    basepath = path.dirname(__file__)
    
    filepath = path.abspath(path.join(basepath, "..", var))
    try:
        remove(filepath)
    except:
        return

def no(write):
    Label(text = 'YES                                     ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)
    return

def check(n=5):
    global root
    global inv_no
    msg = StringVar()
    val = 0
    TEXT = ""

  
    
    try:
        inv = int(inv_no.get())
    except:
        messagebox.showinfo("Invoice number","Please enter a valid Invoice number");
        return
    entry2 = Label(root,text = TEXT).place(x = 152,y = 206,height = 115,width = 200)
    loc = ("dependencies\_details.xlsx") 
    workb = open_workbook(loc)
    label = Label(root,text = "Invoice details" ,fg = 'black').place(x = 10,y = 200)
    
    
    sheet1 = workb.sheet_by_index(0)
    
    try:
        val = sheet1.cell_value(inv,0)
        
    except IndexError:
        TEXT = "Invoice not found"
        
    if( val == inv):
        TEXT = ''
        TEXT = str(sheet1.cell_value(0,0)) + str(" - ") + str(int(sheet1.cell_value(inv,0)))+ '\n'
        TEXT = TEXT +str(sheet1.cell_value(0,1))+str("   - ")+str(sheet1.cell_value(inv,1))+'\n'
        TEXT = TEXT+str(sheet1.cell_value(0,2))+str("        - ")+str(sheet1.cell_value(inv,2))+'\n'
        TEXT = TEXT+str(sheet1.cell_value(0,3))+str("      - ")+str(sheet1.cell_value(inv,3)) +'\n'
        TEXT = TEXT+str(sheet1.cell_value(0,4))+str("   - ")+str(sheet1.cell_value(inv,4))+'\n'
        TEXT = TEXT+str(sheet1.cell_value(0,5))+str("   - ")+str(sheet1.cell_value(inv,5))+'\n'
        
        
        to_write = str(inv+1)
        cmp = str(sheet1.cell_value(inv,4))

        
    else:
        TEXT = "No Invoice Found ....."+"\n"+"Please Add a Invoice..."
            
        
    if len(TEXT) == 0:
        TEXT = "Invoice not found"
    entry2 = Label(root,text = TEXT).place(x = 152,y = 206)
    button = Button(root,text = "Delete",command = lambda:save(to_write,cmp),background = '#4d4dff',activebackground = 'blue')
    entry3 = Button(root,text = "Back",command = lambda:no(to_write),background = '#4d4dff',activebackground = 'blue')
    
    if len(TEXT) > 100:
        button.place(x = 160,y = 330)
        entry3.place(x = 210,y = 330)
    else:
        Label(text = 'YES               ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)



def close():
    global root
    root.destroy()
    basepath = path.dirname(__file__)
    filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
    file_globals = runpy.run_path(filepath)

password1 = Tk()
passw = StringVar()
password1.geometry('300x100+100+100')
password1.configure(bg = '#00bfff')
password1.bind('<Return>',enter)
password1.focus_set()


    

hi = Label(password1,text='',bg='#00bfff')
hi.grid(row = 1)

label = Label(password1,text = 'Password:')
label.grid(row= 2,column = 1,padx = 10)

psw = Entry(password1,textvariable = passw,show='*')
psw.grid(row=2,column = 2)
psw.focus_set()

hi1 = Label(password1,text='',bg='#00bfff')
hi1.grid(row = 3)

but1 = Button(password1,text = "Submit",command = enter,background = '#4d4dff',activebackground = 'blue')
but1.grid(row = 5,column = 2,padx = 10)

cancel = Button(password1,text = "Cancel",command = close2,background = '#4d4dff',activebackground = 'blue')
cancel.grid(row = 5,column = 3)

password1.mainloop()
    
