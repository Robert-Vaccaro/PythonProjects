import tkinter 
from tkinter import *
import xlrd
import xlsxwriter
import tkinter.messagebox
from tkinter import simpledialog as simpledialog
import tkinter.simpledialog as simpledialog
from tkinter.messagebox import showinfo
from tkinter import messagebox, PhotoImage

class AutoScrollbar(Scrollbar):
    # a scrollbar that hides itself if it's not needed.  only
    # works if you use the grid geometry manager.
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            # grid_remove is currently missing from Tkinter!
            self.tk.call("grid", "remove", self)
        else:
            self.grid()
        Scrollbar.set(self, lo, hi)
       
global rowz
rowz= 0

def aClick(self):
    global rowz
    if len(age_ent.get().upper()) == 0:
        pass
    else:
        rowz += 1
        return rowz

def enter1(self):
    DATE = entry001.get()
    D_O_W =entry002.get()
    patFir_ent2=patFir_ent.get().upper()
    phone_ent2=phone_ent.get().upper()
    proc_ent2=proc_ent.get().upper()
    age_ent2=age_ent.get().upper()
    comments_ent2=comments_ent.get().upper()


#########################################################
    site_ent2 = site_ent.get().upper()
    if len(site_ent.get().upper()) == 0:
        site_var2 = site_var.get()
    else:
        site_var2 = site_ent2
#########################################################
    room_ent2 =room_ent.get().upper()

    if len(room_ent.get().upper()) == 0:
        room_var2 = room_var.get()
    else:
        room_var2 = room_ent2
#########################################################
    surg_ent2=surg_ent.get().upper()

    if len(surg_ent.get().upper()) == 0:
        surg_var2=surg_var.get()
    else:
        surg_var2=surg_ent2
#########################################################
    gen_ent2=gen_ent.get().upper()

    if len(gen_ent.get().upper()) == 0:
        gen_var2=gen_var.get()
    else:
        gen_var2=gen_ent2
#########################################################

    anesT_ent2=anesT_ent.get().upper()

    if len(anesT_ent.get().upper()) == 0:
        anesT_var2=anesT_var.get()
    else:
        anesT_var2=anesT_ent2
#########################################################
    
    rnum_ent2 = rnum_ent.get().upper()

    if len(rnum_ent.get().upper()) == 0:
        rnum_var2 =rnum_var.get()
    else:
        rnum_var2 =rnum_ent2
#########################################################

    anesN_ent2=anesN_ent.get().upper()

    if len(anesN_ent.get().upper()) == 0:
        anesN_var2=anesN_var.get()
    else:
        anesN_var2=anesN_ent2
#########################################################


    start_ent2=start_ent.get().upper()

    if len(start_ent.get().upper()) == 0:
        startT_var2=startT_var.get()
    else:
        startT_var2=start_ent2
#########################################################

    dura_ent2=dura_ent.get().upper()

    if len(dura_ent.get().upper()) == 0:
        dura_var2=dura_var.get()

    else:
        dura_var2=dura_ent2
#########################################################

    patLast_ent2=patLast_ent.get().upper()
    if len(patLast_ent.get().upper()) == 0:
        proc_var2=proc_var.get()

    else:
        proc_var2=patLast_ent2
#########################################################
    if len(age_ent.get().upper()) == 0:
        tkinter.messagebox.showinfo("Error", "Please Enter in all patient information before clicking Enter.")
    else:
        rowz = aClick(self)
        patnum=rowz
        rnum_ent.delete(0, 'end')
        site_ent.delete(0,'end')
        room_ent.delete(0,'end')
        anesN_ent.delete(0,END)
        gen_ent.delete(0,END)
        start_ent.delete(0,END)
        dura_ent.delete(0,END)
        patLast_ent.delete(0,END)
        patFir_ent.delete(0,END)
        phone_ent.delete(0,END)
        surg_ent.delete(0,END)
        anesT_ent.delete(0,END)
        proc_ent.delete(0,END)
        proc_ent.delete(0,END)
        age_ent.delete(0,END)
        comments_ent.delete(0,END)
        site_ent.focus_set()
        site_ent.focus()
        worksheet.write(rowz, 0, site_var2)
        worksheet.write(rowz, 1, room_var2)
        worksheet.write(rowz, 2, rnum_var2)
        worksheet.write(rowz, 3, DATE)
        worksheet.write(rowz, 4, D_O_W)
        worksheet.write(rowz, 5, startT_var2)
        worksheet.write(rowz, 6, dura_var2)
        worksheet.write(rowz, 7, patFir_ent2)
        worksheet.write(rowz, 8, phone_ent2)
        worksheet.write(rowz, 9, proc_ent2)
        worksheet.write(rowz,10, surg_var2)
        worksheet.write(rowz,11, anesT_var2)
        worksheet.write(rowz,12, proc_var2)
        worksheet.write(rowz,13, age_ent2)
        worksheet.write(rowz,14, gen_var2)
        worksheet.write(rowz,15, anesN_var2)
        worksheet.write(rowz,16, comments_ent2)
        site_ent.focus_set()
        site_ent.focus()
        infile = open(r'C:\Users\vaccaroR\Documents\Python\Last_Patient_Entered.txt','w')
        
        infile.write(site_var2 + '\n')
        infile.write(room_var2 + '\n')
        infile.write(surg_var2 + '\n')
        infile.write(gen_var2+ '\n')
        infile.write((anesT_var2) + '\n')
        infile.write((rnum_var2) + '\n')
        infile.write(anesN_var2 + '\n')
        infile.write(startT_var2+ '\n')
        infile.write(dura_var2 + '\n')
        infile.write(proc_var2 + '\n')
        infile.write(patLast_ent2+ '\n')
        infile.write(patFir_ent2 + '\n')
        infile.write(phone_ent2 + '\n')
        infile.write(proc_ent2 + '\n')
        infile.write(age_ent2 + '\n')
        infile.write(comments_ent2 + '\n')
        site_ent.focus_set()
        site_ent.focus()
        infile.close()
        site_ent.focus_set()
        site_ent.focus()

        return rowz,patFir_ent2,phone_ent2

        
    root.bind('<Return>', lambda rowz:[enter1(self),aClick(self)])

    site_ent.focus_set()
    site_ent.focus()
    root2.mainloop()


def Error101():
    tkinter.messagebox.showinfo("Error", "Needs information to continue!")


def PatNumClicked():
    tkinter.messagebox.showinfo("Error", "Please Press Enter after patient information has been written.")

def exitLast():
    root2.withdraw()


def question():
    Exit_Or_No = tkinter.messagebox.askquestion("Anesthesia Data Input", "Are you sure you would like to exit? ")
    if Exit_Or_No == 'yes':
        workbook.close()
        root4.destroy()
        root.destroy()
        root2.destroy()

def Escape(self):
    answer10 = messagebox.askquestion("Exiting Program", "Would you like to exit the program?\nYour Progress will saved up to this point.\n\nClick Yes to Exit.")
    if answer10 == 'yes':    
        workbook.close()
        root4.destroy()
        root.destroy()
        root2.destroy()

def Escape1():
    answer11 = messagebox.askquestion("Exiting Program", "Would you like to exit the program?\nYour Progress will saved up to this point.\n\nClick Yes to Exit.")
    if answer11 == 'yes':    
        workbook.close()
        root4.destroy()
        root.destroy()
        root2.destroy()


def on_closing3():
    answer99 = messagebox.askquestion("Exiting Program", "Would you like to exit the program?\nIf you proceed, the program will end.\n\nClick Yes to Exit.")
    if answer99 == 'yes':    
        workbook.close()
        root4.destroy()
        root.destroy()
        root2.destroy()

def on_closing():
    answer9 = messagebox.askquestion("Exiting Program", "Would you like to exit the program?\nYour Progress will saved up to this point.\n\nClick Yes to Exit.")
    if answer9 == 'yes':    
        workbook.close()
        root4.destroy()
        root.destroy()
        root2.destroy()



def DateDOW(self):
    DATE = entry001.get()
    D_O_W =entry002.get()
    root4.withdraw()
    root.deiconify()



workbook = xlsxwriter.Workbook('Daily_Anes_Schedule.xlsx')
worksheet = workbook.add_worksheet()
worksheet.add_table(0, 0, 199 , 16)

worksheet.set_column('A:A', 15, None)
worksheet.set_column('B:B', 15, None)
worksheet.set_column('C:C', 15, None)
worksheet.set_column('D:D', 15, None)
worksheet.set_column('E:E', 15, None)
worksheet.set_column('F:F', 15, None)
worksheet.set_column('G:G', 15, None)
worksheet.set_column('H:H', 30, None)
worksheet.set_column('I:I', 15, None)
worksheet.set_column('J:J', 30, None)
worksheet.set_column('K:K', 30, None)
worksheet.set_column('L:L', 15, None)
worksheet.set_column('M:M', 15, None)
worksheet.set_column('N:N', 15, None)
worksheet.set_column('O:O', 15, None)
worksheet.set_column('P:P', 15, None)
worksheet.set_column('Q:Q', 40, None)
worksheet.add_table("A1:Q15", {'columns': [{'header': 'Site'},
                                        {'header': 'Room Type (OP/AM/IP)'},
                                        {'header': 'Room Number'},
                                        {'header': 'Date'},
                                        {'header': 'Day Of Week'},
                                        {'header': 'Start Time'},
                                        {'header':'Duration'},
                                        {'header':'PT Last Name'},
                                        {'header':'PT First Name'},
                                        {'header':'PT Phone Number'},
                                        {'header':'Surgeon'},
                                        {'header':'Anesthesia Type'},
                                        {'header':'Procedure'},
                                        {'header':'Age'},
                                        {'header':'Gender'},
                                        {'header':'Assigned Anesthesiologist'},
                                        {'header':'Comments'}
                                        ]})

root =Tk() 
root4=Tk()
root4.title("Date and Day of the Week")
label00 = Label(root4, text = "Please Enter the Date  ")
label001= Label(root4, text = "Please Enter the Day of the Week:  ")
label221=Label(root4, text = "\n\nWhen finished entering \nthe Date and Day of the Week:   ")
label2221=Label(root4, text = "Press Shift + Enter",font=('Helvetica', 12, 'bold'))
label33 = Label(root4, text = "To go to next Entry Box:   ")
label332= Label(root4,text = "Press Tab",font=('Helvetica', 12, 'bold'))
entry001= Entry(root4) #like input() but for windows, root is where you want it, shows single blank line
entry002= Entry(root4)
label00.grid(row=0, sticky = E) #where on the grid do you want to put it, column is by default 0, use N, W, E, and S for postion to stick to a sort of Wall of the grid
label001.grid(row=1, sticky = E)
label221.grid(row=2, column= 2+3)
label2221.grid(row=3, column=2+3)
label33.grid(row=0, column=2+3)
label332.grid(row=1,column=2+3)
entry001.grid(row=0, column=1) #want it to the right of the label in column 2
entry002.grid(row=1, column=1)
entry001.focus()
root4.bind('<Shift-Return>', DateDOW)
root4.protocol("WM_DELETE_WINDOW", on_closing3)
DATE = entry001.get()
D_O_W =entry002.get()

 #creating a blank window, that is what root is going to be equal to
root2 = Tk()

menu1 = Menu(root)
root.config(menu=menu1)

subMenu = Menu(menu1)

menu1.add_cascade(label = "File", menu=subMenu)

subMenu.add_command(label="Exit", command= Escape1)


site_var = StringVar(root)
site_var.set("GSH")

room_var = StringVar(root)
room_var.set("OP")

surg_var = StringVar(root)
surg_var.set("")

gen_var = StringVar(root)
gen_var.set("F")

anesT_var = StringVar(root)
anesT_var.set("GEN")

rnum_var = StringVar(root)
rnum_var.set('1')

anesN_var = StringVar(root)
anesN_var.set('')

startT_var = StringVar(root)
startT_var.set('7:30AM')

dura_var= StringVar(root)
dura_var.set('')

proc_var= StringVar(root)
proc_var.set('')



site_menu = OptionMenu(root, site_var,"", "GSH", "LGSC", "BSC", "SVSC", "EMISON", "FSC", "ECSV", "SSSV", "GSH OB", "ASC", "KNOWLES")
site_menu.grid(row=0, column=1,sticky='W'+'E')

room_menu = OptionMenu(root, room_var,"", "OP", "AM", "IP","OTHER")
room_menu.grid(row=1, column=1,sticky='W'+'E')

surg_menu = OptionMenu(root, surg_var, "")
surg_menu.grid(row=2, column=1,sticky='W'+'E')

gen_menu = OptionMenu(root, gen_var,'', "F","M","UNKNOWN")
gen_menu.grid(row=3, column=1,sticky='W'+'E')

anesT_menu = OptionMenu(root, anesT_var,"", "GEN", "MAC", "CHOICE","GEN + REGIONAL", "REGIONAL")
anesT_menu.grid(row=4, column=1,sticky='W'+'E')

anesN_menu = OptionMenu(root, anesN_var,"")
anesN_menu.grid(row=5, column=1,sticky='W'+'E')

rnum_menu = OptionMenu(root, rnum_var, '1','2','3','4','5','6','7','8','9','10','11','12','ENDO 1','ENDO 2','PEDI 1','PEDI 2','MRAD','RADIO','CVS SS RM','CVS RM 1','CVS RM 2','CVS RM 3','CVS RM 4','')
rnum_menu.grid(row = 6, column=1,sticky='W'+'E')

startT_menu = OptionMenu(root, startT_var, "",
                         '6:00AM',
                         '6:30AM',
                         '7:00AM',
                         '7:30AM',
                         '8:00AM',
                         '8:30AM',
                         '9:00AM',
                         '9:30AM',
                         '10:00AM',
                         '10:30AM',
                         '11:00AM',
                         '11:30AM',
                         '12:00',
                         '12:30',
                         '13:00',
                         '13:30',
                         '14:00',
                         '14:30',
                         '15:00',
                         '15:30',
                         '16:00',
                         '16:30',
                         '17:00',
                         '17:30',
                         '18:00',
                         '18:30',
                         '19:00',
                         '19:30',
                         '20:00',
                         '20:30',
                         '21:00',
                         '21:30',
                         '22:00',
                         '22:30',
                         '23:00',
                         '23:30',
                         '00:00AM')
startT_menu.grid(row=7, column=1,sticky='W'+'E')

dura_menu = OptionMenu(root, dura_var, "",'30MINS','60MINS','90MINS','120MINS','150MINS','180MINS','210MINS','240MINS','270MINS','300MINS','330MINS','360MINS','390MINS','420MINS','450MINS','480MINS')
dura_menu.grid(row=8, column=1,sticky='W'+'E')

proc_menu = OptionMenu(root, proc_var, "")
proc_menu.grid(row=9, column=1,sticky='W'+'E')






root.title('Anesthesia Scheduling Program')
hos_lab = Label(root, text = "The hospital site:")
room_lab = Label(root, text = "Type of room:")
surg_lab = Label(root, text = "Surgeon Last Name:")
gen_lab= Label(root, text = "Gender: ")
anesT_lab= Label(root, text = "Anesthesia Type: ")
anesN_lab= Label(root, text = "Anesthesiologist: ")
rnum_lab = Label(root, text = "Room Number:")
start_lab = Label(root, text = "Start Time:")
dura_lab = Label(root, text = "Duration:")
patLast_lab = Label(root, text = "Patient Last Name:")
patFir_lab = Label(root, text = "Patient First Name:")
phone_lab = Label(root, text = "Phone Number:")
proc_lab = Label(root, text = "Procedure:")
age_lab = Label(root, text = "Age:")
comments_lab = Label(root, text = "Comments: ")



hos_lab.grid(row=0, sticky = 'E')
room_lab.grid(row=1, sticky = 'E')
surg_lab.grid(row=2, sticky = 'E')
gen_lab.grid(row=3, sticky='E')
anesT_lab.grid(row=4, sticky='E')
anesN_lab.grid(row=5, sticky='E')
rnum_lab.grid(row=6, sticky = 'E')
start_lab.grid(row=7, sticky = 'E')
dura_lab.grid(row=8, sticky = 'E')
proc_lab.grid(row=9, sticky = 'E')
patLast_lab.grid(row=10, sticky = 'E')
patFir_lab.grid(row=11, sticky = 'E')
phone_lab.grid(row=12, sticky = 'E')
age_lab.grid(row=13, sticky = 'E')
comments_lab.grid(row=14, sticky ='E')


site_ent= Entry(root)
room_ent= Entry(root)
surg_ent= Entry(root)
gen_ent= Entry(root)
anesT_ent= Entry(root)
anesN_ent= Entry(root)
rnum_ent= Entry(root)
start_ent= Entry(root)
dura_ent= Entry(root)
patLast_ent= Entry(root)
patFir_ent= Entry(root)
phone_ent= Entry(root)
proc_ent= Entry(root)
age_ent= Entry(root)
comments_ent = Entry(root)


site_ent.focus()


site_ent.grid(row=0, column=2, padx = 15, pady=3, sticky=W+E)
room_ent.grid(row=1, column=2, padx = 15, pady=3, sticky=W+E)
surg_ent.grid(row=2, column=2, padx = 15, pady=3, sticky=W+E)
gen_ent.grid(row=3, column=2, padx = 15, pady=3, sticky=W+E)
anesT_ent.grid(row=4, column=2, padx = 15, pady=3, sticky=W+E)
anesN_ent.grid(row=5, column=2, padx = 15, pady=3, sticky=W+E)
rnum_ent.grid(row=6, column=2, padx = 15, pady=3, sticky=W+E) 
start_ent.grid(row=7, column=2, padx = 15, pady=3, sticky=W+E) 
dura_ent.grid(row=8, column=2, padx = 15, pady=3, sticky=W+E) 
patLast_ent.grid(row=9, column=2, padx = 15, pady=3, sticky=W+E) 
patFir_ent.grid(row=10, column=1, padx = 15, pady=3, sticky=W+E) 
phone_ent.grid(row=11, column=1, padx = 15, pady=3, sticky=W+E) 
proc_ent.grid(row=12, column=1, padx = 15, pady=3, sticky=W+E) 
age_ent.grid(row=13, column=1, padx = 15, pady=3, sticky=W+E) 
comments_ent.grid(row=14, column=1, padx = 15, pady=3, sticky=W+E) 

site_ent.focus()
pat_label2 = Label(root, text = '                                      ')
pat_label2.grid(row=17,column=1)
button5 = Button( text = "Exit", fg = "black", command=question)
button5.grid(row=17, column=2, pady=25)

root.protocol("WM_DELETE_WINDOW", on_closing)



root2.withdraw()

root2.protocol("WM_DELETE_WINDOW", exitLast)
root.bind('<Return>', enter1)
root.bind('<Escape>', Escape)
infile = open(r'Last_Patient_Entered.txt','w')
site_ent.focus()
root.iconify()

root4.mainloop()
root.mainloop()
