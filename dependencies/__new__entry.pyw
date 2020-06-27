from tkinter import *
from tkinter import messagebox
from xlrd import open_workbook
import openpyxl
from os import path
import runpy

def close():
    window.destroy()
    basepath = path.dirname(__file__)
    filepath = path.abspath(path.join(basepath, "..", "__invoice__generator__.pyw"))

    file_globals = runpy.run_path(filepath)
    
def new_entry(n = 5):
    short = short_name.get()
    comp = company_name.get()
    add1 = add_1.get()
    add2 = add_2.get()
    add3 = add_3.get()
    gst = gst_num.get()

    if len(short) == 0 or len(comp) == 0 or len(add1) == 0 or len(add1) == 0 or len(add2) == 0 or len(add3) == 0 or len(gst) == 0: 
        messagebox.showinfo("Empty Field","Fill all the fields");
        return
    

    loc = ("dependencies\__companies.xlsx")
    workb = open_workbook(loc)
    sheet1 = workb.sheet_by_index(0)
    
    
    for i in range(10000):
        try:
            val = sheet1.cell_value(i,0)
            if len(val) > 0:
                continue
            else:
                break
        except IndexError:
            break
    to_write = []
    xfile = openpyxl.load_workbook("dependencies\__companies.xlsx")
    sheet = xfile.get_sheet_by_name('companies')
    alpha = 'A'

    gst_no = "GST NO : " + str(gst)
    
    for j in range(6):
        val = alpha + str(i+1)
        nxt = ord(alpha)
        nxt += 1
        alpha = chr(nxt)
        to_write.append(val)

    loc = ("dependencies\__companies.xlsx")
    workb = open_workbook(loc)
    sheet1 = workb.sheet_by_index(0)
    print(i)
    for k in range(i):
        value = sheet1.cell_value(k,0)
        print(value)
        if value == short:
            messagebox.showinfo("Already Found","Company already found");
            return
    sheet[to_write[0]] = short   
    sheet[to_write[1]] = comp
    sheet[to_write[2]] = add1
    sheet[to_write[3]] = add2
    sheet[to_write[4]] = add3
    sheet[to_write[5]] = gst_no

    xfile.save("dependencies\__companies.xlsx")
    Label(frame1,text = "Saved Successfully.",bg = '#00bfff').grid(row = 10,column = 5,padx = 5, pady = 10,ipady = 7)
    Button(frame1,text = "Go back",command = close,bg = '#4d4dff',activebackground = 'blue').grid(row = 10,column = 6,padx = 15, pady = 10)
    
window = Tk()
window.title("Add a Company..")
window.iconbitmap(r'dependencies\icon.ico')
window.geometry('500x400+100+100')
window.configure(bg = '#4d4dff')


Label(window,text = "KKR Engineering", font = ('Calibri',16),bg = '#4d4dff',fg = 'white').grid()
frame1 = Frame(window,relief = RAISED,bd = 10, bg = '#00bfff')
frame1.grid(row = 5,padx = 10,pady = 10)
window.bind('<Return>',new_entry)
window.focus_set()


short_name = StringVar()
company_name = StringVar()
add_1 = StringVar()
add_2 = StringVar()
add_3 = StringVar()
gst_num = StringVar()

Label(frame1,text = "Short name:").grid(row = 1,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = short_name)
entry1.grid(row = 1,column = 4,padx = 5,pady = 10)
entry1.focus_set()


Label(frame1,text = "Company Name:").grid(row = 2,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = company_name).grid(row = 2,column = 4,padx = 5,pady = 10)

Label(frame1,text = "Address line 1:").grid(row = 3,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = add_1).grid(row = 3,column = 4,padx = 5,pady = 10)

Label(frame1,text = "Address line 2:").grid(row = 4,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = add_2).grid(row = 4,column = 4,padx = 5,pady = 10)

Label(frame1,text = "Address line 3:").grid(row = 5,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = add_3).grid(row = 5,column = 4,padx = 5,pady = 10)

Label(frame1,text = "GST Number:").grid(row = 6,column = 1,padx = 5,pady = 10)
entry1 = Entry(frame1,textvariable = gst_num).grid(row = 6,column = 4,padx = 5,pady = 10)


button1 = Button(frame1,text = "Submit",command = new_entry,bg = '#4d4dff',activebackground = 'blue')
button1.grid(row = 10,column = 5,padx = 5,pady = 10)

button2 = Button(frame1,text = "Cancel",command = close,bg = '#4d4dff',activebackground = 'blue')
button2.grid(row = 10 , column = 6,padx = 5,pady = 10)
window.mainloop()
