from tkinter import *
from tkinter import messagebox
import sys
from xlrd import open_workbook
import openpyxl
import warnings
from os import path
import runpy


def save(write):
    Label(text = 'YES                                     ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)
    new = openpyxl.load_workbook('dependencies\_details.xlsx')
    first = new.get_sheet_by_name('Details')
    first[write] = "yes"
    new.save('dependencies\_details.xlsx')

def no(write):
    Label(text = 'YES                                     ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)
    new = openpyxl.load_workbook('dependencies\_details.xlsx')
    first = new.get_sheet_by_name('Details')
    first[write] = "no"
    new.save('dependencies\_details.xlsx')

def check(n = 5):
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
        if(sheet1.cell_value(1,0) == 1):
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
                
                
                to_write = 'I' + str(inv+1)

            
        else:
            TEXT = "No Invoice Found ....."+"\n"+"Please Add a Invoice..."
            
    except IndexError:
        TEXT = "No Invoice Found .....Please Add a Invoice..."
    if len(TEXT) == 0:
        TEXT = "Invoice not found"
    entry2 = Label(root,text = TEXT).place(x = 152,y = 206)
    button = Button(root,text = "YES",command = lambda:save(to_write),background = '#4d4dff',activebackground = 'blue')
    entry3 = Button(root,text = "NO",command = lambda:no(to_write),background = '#4d4dff',activebackground = 'blue')
    
    if len(TEXT) > 100:
        button.place(x = 160,y = 330)
        entry3.place(x = 200,y = 330)
    else:
        Label(text = 'YES               ',bg = '#00bfff',fg = '#00bfff').place(x = 160,y = 330,height = 40)



def close():
    root.destroy()
    basepath = path.dirname(__file__)
    filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))
    file_globals = runpy.run_path(filepath)
    
root = Tk()
root.geometry('366x450+100+100')
root.iconbitmap(r'dependencies\icon.ico')
root.title('Add received payment')
root.bind('<Return>',check)
root.configure(bg = '#00bfff')
root.focus_set()

inv_no = StringVar()
label = Label(root,text = "KKR Engineering",bg = '#00bfff',fg = 'white',font= ("calibri",16)).place(x = 102,y = 20)

label = Label(root,text = "Invoice Number:",fg = 'black').place(x = 10,y = 100)
entry = Entry(root,textvariable = inv_no)
entry.place(x = 160,y = 100)
entry.focus_set()


submit = Button(root,text = "Search",command = check,background = '#4d4dff',activebackground = 'blue').place(x = 230,y = 400)
cancel = Button(root,text = "Cancel",command = close,background = '#4d4dff',activebackground = 'blue').place(x = 290,y = 400)





    


mainloop()
