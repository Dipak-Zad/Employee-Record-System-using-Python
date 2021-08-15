#Employee Record System 
from tkinter import*
from tkinter import messagebox
from openpyxl import load_workbook
import xlrd
import pandas as pd


def emp_dict(*args):                   #To add a new entry and check if entry already exist in excel sheet
    #print("done")
    workbook_name="sample.xlsx"
    workbook=xlrd.open_workbook(workbook_name)
    worksheet=workbook.sheet_by_index(0)

    wb=load_workbook(workbook_name)
    page=wb["Employee"]
    
    p=0
    for i in range(worksheet.nrows):
        for j in range(worksheet.ncols):
            cellvalue=worksheet.cell_value(i,j)
            print(cellvalue)   
            sheet_data.append([])
            sheet_data[p]=cellvalue
            p+=1
    print(sheet_data)
    fl=firstname.get()
    fsl=fl.lower()
    ll=lastname.get()
    lsl=ll.lower()
    if (fsl and lsl) in sheet_data:
        print("found")
        messagebox.showerror("Error","This Employee already exist")
    else:
        print("not found")
        for info in args:
            page.append(info)
        messagebox.showinfo("Done","Successfully added the employee record")

    wb.save(filename=workbook_name)
    
def add_entries():                       #to append all data and add entries on click the button
    a=" "
    e=empid.get()
    ei=e.lower()
    f=firstname.get()
    f1=f.lower()
    l=lastname.get()
    l1=l.lower()
    d=dept.get()
    d1=d.lower()
    de=designation.get()
    de1=de.lower()
    ad=empaddress.get()
    ea=ad.lower()
    pno=emppno.get()
    epn=pno.lower()
    blg=empbg.get()
    ebg=blg.lower()
    ml=empmail.get()
    em=ml.lower()
    list1=list(a)
    list1.append(ei)
    list1.append(f1)
    list1.append(l1)
    list1.append(d1)
    list1.append(de1)
    list1.append(ea)
    list1.append(epn)
    list1.append(ebg)
    list1.append(ml)
    emp_dict(list1)
    print(list1)

def add_info():                                           #for taking user input to add the enteries
    frame2.pack_forget()
    frame3.pack_forget()
    emp_id=Label(frame1,text="Enter employee ID: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_id.grid(row=1,column=1,padx=10)
    e0=Entry(frame1,textvariable=empid)
    e0.grid(row=1,column=2,padx=10)
    emp_first_name=Label(frame1,text="Enter first name of the employee: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_first_name.grid(row=2,column=1,padx=10)
    e1=Entry(frame1,textvariable=firstname)
    e1.grid(row=2,column=2,padx=10)
    e1.focus()
    emp_last_name=Label(frame1,text="Enter last name of the employee: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_last_name.grid(row=3,column=1,padx=10)
    e2=Entry(frame1,textvariable=lastname)
    e2.grid(row=3,column=2,padx=10)
    emp_dept=Label(frame1,text="Select department of employee: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_dept.grid(row=4,column=1,padx=10)
    dept.set("Select Option")
    e4=OptionMenu(frame1,dept,"Select Option","IT","Operations","Sales")
    e4.configure(font=('Rubik',10,'bold'), fg='white', bg='black', border='0')
    e4.grid(row=4,column=2,padx=10)
    emp_desig=Label(frame1,text="Select designation of Employee: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_desig.grid(row=5,column=1,padx=10)
    designation.set("Select Option")
    e5=OptionMenu(frame1,designation,"Select Option","Manager","Asst Manager","Project Manager","Team Lead","Senior Tester", 
                  "Junior Tester","Senior Developer","Junior Developer","Intern")
    e5.configure(font=('Rubik',10,'bold'), fg='white', bg='black', border='0')
    e5.grid(row=5,column=2,padx=10)
    emp_address=Label(frame1,text="Enter employee address: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_address.grid(row=6,column=1,padx=10)
    e6=Entry(frame1,textvariable=empaddress)
    e6.grid(row=6,column=2,padx=10)
    emp_phone_number=Label(frame1,text="Enter employee phone no: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_phone_number.grid(row=7,column=1,padx=10)
    e7=Entry(frame1,textvariable=emppno)
    e7.grid(row=7,column=2,padx=10)
    emp_blood_group=Label(frame1,text="Enter employee blood grp: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_blood_group.grid(row=8,column=1,padx=10)
    e8=Entry(frame1,textvariable=empbg)
    e8.grid(row=8,column=2,padx=10)
    emp_mail=Label(frame1,text="Enter employee mail id: ",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_mail.grid(row=9,column=1,padx=10)
    e9=Entry(frame1,textvariable=empmail)
    e9.grid(row=9,column=2,padx=10)
    button4=Button(frame1,text="Add Employee",command=add_entries,font=('Rubik',10,'bold'), fg='white', bg='black')
    button4.grid(row=10,column=2,pady=10)
    
    frame1.configure(background="Red")
    frame1.pack(pady=10)


def clear_all():  
    # f.pack_forget()           #for clearing the entry widgets
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()
    loginF.pack_forget()

    
def remove_emp():                #for taking user input to remove enteries
    clear_all()
    emp_first_name=Label(frame2,text="Enter first name of the employee:",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_first_name.grid(row=1,column=1,padx=10)
    e10=Entry(frame2,textvariable=remove_firstname)
    e10.grid(row=1,column=2,padx=10)
    e10.focus()
    emp_last_name=Label(frame2,text="Enter last name of the employee:",bg="red",fg="white",font=('Lato',11,'bold'))
    emp_last_name.grid(row=2,column=1,padx=10)
    e11=Entry(frame2,textvariable=remove_lastname)
    e11.grid(row=2,column=2,padx=10)
    remove_button=Button(frame2,text="Remove Employee",command=remove_entry,font=('Rubik',10,'bold'), fg='white', bg='black')
    remove_button.grid(row=3,column=2,pady=10)
    frame2.configure(background="Red")
    frame2.pack(pady=10)

def remove_entry():  #to remove entry from excel sheet
    rsf=remove_firstname.get()
    rsf1=rsf.lower()
    print(rsf1)
    rsl=remove_lastname.get()
    rsl1=rsl.lower()
    print(rsl1)
    # workbook_name="sample.xlsx"
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    
    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        print(row_value)
        if (row_value[2]==rsf1 and row_value[3]==rsl1):
            print(row_value)
            print("found")
            file="sample.xlsx"
            x=pd.ExcelFile(file)
            writer=pd.ExcelWriter('sample.xlsx', engine='xlsxwriter')
            df1=x.parse(x.sheet_names[0])
            df2=x.parse(x.sheet_names[1])

            df1=df1[df1['First Name']!=rsf]
            dfs = {'Employee':df1,'Login':df2}
            for sheet_name in dfs.keys():
                dfs[sheet_name].to_excel(writer,sheet_name=sheet_name,index=False)
            writer.save()
            messagebox.showinfo("Done","Successfully removed the Employee record")
            break
        # else:
        #     print("Error occured")
        #     messagebox.showerror("Error","Employee does not exist")
    clear_all()


def search_emp():     #can implement search by 1st name,last name,emp id, designation
    clear_all()
    emp_first_name=Label(frame3,text="Enter first name of the employee:",bg="red",fg="white",font=('Lato',12,'bold'))   #to take user input to seach
    emp_first_name.grid(row=1,column=1,padx=10)#place(x=10, y=20)
    e12=Entry(frame3,textvariable=searchfirstname)
    e12.grid(row=1,column=2,padx=10)
    e12.focus()
    emp_last_name=Label(frame3,text="Enter last name of the employee:",bg="red",fg="white",font=('Lato',12,'bold'))
    emp_last_name.grid(row=2,column=1,padx=10)
    e13=Entry(frame3,textvariable=searchlastname)
    e13.grid(row=2,column=2,padx=10)
    search_button=Button(frame3,text="Search Employee",command=search_entry,font=('Rubik',10,'bold'), fg='white', bg='black')
    search_button.grid(row=3,column=2,pady=10)
    nameval=""


    frame3.configure(background="Red")
    frame3.pack(pady=10)

def displaySearchItem(fnameval,lnameval,deptval,designnationval,addval,phoneval,bgroupval,mailval):
    
    searchresultF=Tk()
    searchresultF.geometry("400x400")
    searchresultF.title("Employee Details")
    searchresultF.config(bg='red')


    sresultTitle = Label(searchresultF,text="Search Result") 
    sresultTitle.config(font=("Helvatica",20,"bold"))   
    sresultTitle.grid(row=1,column=0,columnspan=6,pady=10)

    slabel = Label(searchresultF,text="Employee First Name : ")
    slabel.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel.grid(row=2,column=0,pady=10)
    sfname = Label(searchresultF,text=fnameval)
    sfname.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sfname.grid(row=2,column=3,pady=10)
    
    slabel2 = Label(searchresultF,text="Employee Last Name: ")
    slabel2.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel2.grid(row=3,column=0,pady=10)
    slname = Label(searchresultF,text=lnameval)
    slname.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    slname.grid(row=3,column=3,pady=10)
    
    slabel3 = Label(searchresultF,text="Employee Department : ")
    slabel3.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel3.grid(row=4,column=0,pady=10)
    sdept = Label(searchresultF,text=deptval)
    sdept.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sdept.grid(row=4,column=3,pady=10)
    
    slabel4 = Label(searchresultF,text="Employee Designation : ")
    slabel4.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel4.grid(row=5,column=0,pady=10)
    sdesignation = Label(searchresultF,text=designnationval)
    sdesignation.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sdesignation.grid(row=5,column=3,pady=10)
    
    slabel5 = Label(searchresultF,text="Employee Address : ")
    slabel5.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel5.grid(row=6,column=0,pady=10)
    sadd = Label(searchresultF,text=addval)
    sadd.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sadd.grid(row=6,column=3,pady=10)
    
    slabel6 = Label(searchresultF,text="Employee Phone Number : ")
    slabel6.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel6.grid(row=7,column=0,pady=10)
    sphone = Label(searchresultF,text=phoneval)
    sphone.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sphone.grid(row=7,column=3,pady=10)
    
    slabel7 = Label(searchresultF,text="Employee Blood Group : ")
    slabel7.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel7.grid(row=8,column=0,pady=10)
    sbgroup = Label(searchresultF,text=bgroupval)
    sbgroup.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    sbgroup.grid(row=8,column=3,pady=10)
    
    slabel8 = Label(searchresultF,text="Employee Gmail : ")
    slabel8.config(font=("Helvatica",10,"bold"), fg='black', bg='red')
    slabel8.grid(row=9,column=0,pady=10)
    smail = Label(searchresultF,text=mailval)
    smail.config(font=("Helvatica",10,"bold"), fg='white', bg='black')
    smail.grid(row=9,column=3,pady=10)

    searchresultF.mainloop()
    
def search_entry():
    sf=searchfirstname.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=searchlastname.get()
    ssl1=sl.lower()
    print(ssl1)
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)
    val=0
    log=1
    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[2]==ssf1 and row_value[3]==ssl1):
            fnameval=row_value[2]
            lnameval=row_value[3]
            deptval=row_value[4]
            designnationval=row_value[5]
            addval=row_value[6]
            phoneval=row_value[7]
            bgroupval=row_value[8]
            mailval=row_value[9]
            print(row_value)
            print("found")
            displaySearchItem(fnameval,lnameval,deptval,designnationval,addval,phoneval,bgroupval,mailval)
            clear_all()
            val=1
            home()
            break
        else:
            if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
                print("Not found")
                log=0
                clear_all()
    if(log==0):
        if(val==0):
           messagebox.showerror("Sorry","Employee Record does not Exist") 
           clear_all()
           frame3.pack()


def home():
    clear_all()
    label2=Label(f,text="Select an action: ", background="Black", fg="White", font=('Lato',11,'bold'))
    label2.pack(side=LEFT,pady=10)
    button1=Button(f,text="Add", background="Red", fg="White", command=add_info, width=8, font=('Italic',11,'bold'))
    button1.pack(side=LEFT,ipadx=20,pady=10)
    button2=Button(f,text="Remove", background="Red", fg="white", command=remove_emp, width=8, font=('Italic',11,'bold'))
    button2.pack(side=LEFT,ipadx=20,pady=10)
    button3=Button(f,text="Search", background="Red", fg="White", command=search_emp, width=8, font=('Italic',11,'bold'))
    button3.pack(side=LEFT,ipadx=20,pady=10)
    button6=Button(f,text="Close", background="Red", fg="White", width=8, command=root.destroy, font=('Italic',11,'bold'))
    button6.pack(side=LEFT,ipadx=20,pady=10)
    f.configure(background="Black")
    f.pack()

def validatelogin():
    sf=userval.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=passval.get()
    ssl1=sl.lower()
    print(ssl1)
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(1)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==ssf1 and row_value[2]==ssl1):
            clear_all()
            val=1
            home()
            break
        else:
            if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
                print("Not found")
                log=0
                clear_all()
                loginF.pack()
                
    if(log==0):
        if(val==0):
            userval.delete(0,END)
            passval.delete(0,END)
            messagebox.showerror("Sorry","Username/Password Invalid")

def login():
    clear_all()
    logtitle=Label(loginF,text="Login To Continue")
    logtitle.config(font=('Lato',16,'bold'))
    logtitle.configure(background="red")
    #root.wm_attributes('-transparentcolor',root['bg'])
    # logtitle.grid(row=0,column=0,columnspan=3)
    logtitle.pack()

    userlabel=Label(loginF,text="Enter Username")
    userlabel.config(font=('Lato',11,'bold'))
    userlabel.configure(background="darkgrey")
    userlabel.pack()
    # userlabel.grid(row=2,column=0,columnspan=2)

    global userval
    userval=Entry(loginF)
    userval.pack()
    # userval.grid(row=2,column=0)

    passlabel=Label(loginF,text="Enter Password")
    passlabel.config(font=('Lato',11,'bold'))
    passlabel.configure(background="darkgrey")
    passlabel.pack()
    # passlabel.grid(row=3,column=0,columnspan=2)

    global passval
    passval=Entry(loginF, show='*')
    passval.pack()
    # passval.grid(row=2,column=0)

    loginB=Button(loginF,text="Login", command=validatelogin, font=('Rubik',11,'bold'), fg='white')
    # loginB.grid(row=4,column=2)
    loginB.configure(background="black")
    loginB.pack()
    loginF.configure(background="darkgrey")
    loginF.pack()





root=Tk()  

loginF=Frame(root)                      #Main window 
f=Frame(root)
frame1=Frame(root)
frame2=Frame(root)
frame3=Frame(root)


root.title("Employee Record Management System")
root.geometry("830x395")
#root.configure(background="ERSbg1.png")
bg = PhotoImage(file="E:\Downloads\mainbg.png")
my_label = Label (root, image=bg)
my_label.place(x=0, y=0, relwidth=1, relheight=1)
my_label.lower()
# scrollbar=Scrollbar(root)
# scrollbar.pack(side=RIGHT, fill=Y)

empid=StringVar()                       #Declaration of all variables
firstname=StringVar()                    
lastname=StringVar()
id=StringVar()
dept=StringVar()
designation=StringVar()
empaddress=StringVar()
emppno=StringVar()
empbg=StringVar()
empmail=StringVar()
remove_firstname=StringVar()
remove_lastname=StringVar()
searchfirstname=StringVar()
searchlastname=StringVar()
sheet_data=[]
row_data=[]
val=0
log=1

#Main window buttons and labels
        
label1=Label(root,text="EMPLOYEE RECORD MANAGEMENT SYSTEM")
label1.config(font=('Italic',16,'bold'), justify=CENTER, background="Black",fg="Red", anchor="center")
label1.pack(fill=X)

login()


root.mainloop()
