from tkinter import*
from PIL import ImageTk #pip install pillow
from tkinter import messagebox,filedialog
import os
import pandas as pd #pip install pandas
import email_function

class bulk_email:
    def __init__(self,root):
        self.root=root
        self.root.title("Bulk Email Application")
        self.root.geometry("1000x550+165+100")
        self.root.resizable(False,False)
        self.root.config(bg="white")

        #=========== Icons / Iconos ========================

        root.iconbitmap("images/icon.ico")
        self.pass_icon=ImageTk.PhotoImage(file="images/pass_block.png")
        self.pass2_icon=ImageTk.PhotoImage(file="images/pass_normal.png")
        self.email_icon=ImageTk.PhotoImage(file="images/email.png")
        self.setting_icon=ImageTk.PhotoImage(file="images/setting.png")

        #=========== Title / Título ========================
        title=Label(self.root,text="Bulk Email Send Pannel",image=self.email_icon,padx=10,compound=LEFT,font=("Tahoma",40,"bold"),bg="#1289A7",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        desc=Label(self.root,text="Use Excel File to Send the Bulk Email at once, with just click. Ensure the Email Column Name must be Email",font=("calibri (body)",14),bg="#f6e58d",fg="black").place(x=0,y=78,relwidth=1)

        btn_setting=Button(self.root,image=self.setting_icon,bd=0,activebackground="#1289A7",bg="#1289A7",cursor="hand2",command=self.setting_win).place(x=900,y=5)

        #===================================================
        self.var_choice=StringVar()
        single=Radiobutton(self.root,text="Single",value="single",variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=50,y=150)
        bulk=Radiobutton(self.root,text="Bulk",value="bulk",variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=200,y=150)
        self.var_choice.set("single")

        #===================================================

        to=Label(self.root,text="To (Email Adress)",font=("times new roman",18),bg="white").place(x=50,y=250)
        subj=Label(self.root,text="SUBJECT",font=("times new roman",18),bg="white").place(x=50,y=300)
        msg=Label(self.root,text="MESSAGE",font=("times new roman",18),bg="white").place(x=50,y=350)

        self.txt_to=Entry(self.root,font=("times new roman",14),bg="#c8d6e5")
        self.txt_to.place(x=300,y=250,width=350,height=30)

        self.btn_browse=Button(self.root,command=self.browse_file,text="BROWSE",font=("times new roman",13,"bold"),bg="#12CBC4",fg="black",activebackground="#12CBC4",activeforeground="black",cursor="hand2",state=DISABLED)
        self.btn_browse.place(x=670,y=250,width=120,height=30)

        self.txt_subj=Entry(self.root,font=("times new roman",14),bg="#c8d6e5")
        self.txt_subj.place(x=300,y=300,width=450,height=30)

        self.txt_msg=Text(self.root,font=("times new roman",12),bg="#c8d6e5")
        self.txt_msg.place(x=300,y=350,width=650,height=100)

        #================= Status  =========================
        self.lbl_total=Label(self.root,font=("times new roman",18),bg="white")
        self.lbl_total.place(x=50,y=485)

        self.lbl_sent=Label(self.root,font=("times new roman",18),bg="white",fg="green")
        self.lbl_sent.place(x=296,y=485)

        self.lbl_left=Label(self.root,font=("times new roman",18),bg="white",fg="orange")
        self.lbl_left.place(x=412,y=485)

        self.lbl_failed=Label(self.root,font=("times new roman",18),bg="white",fg="red")
        self.lbl_failed.place(x=542,y=485)

        #===================================================

        btn_clear=Button(self.root,text="CLEAR",command=self.clear1,font=("Tahoma",18,"bold"),bg="#30336b",fg="white",activebackground="#30336b",activeforeground="white",cursor="hand2").place(x=690,y=480,width=120,height=35)
        btn_send=Button(self.root,text="SEND",command=self.send_email,font=("Tahoma",18,"bold"),bg="#32ff7e",fg="white",activebackground="#32ff7e",activeforeground="white",cursor="hand2").place(x=830,y=480,width=120,height=35)
        self.check_file_exist()

    def browse_file(self):
        op=filedialog.askopenfile(initialdir="/",title="Select Excel File for Emails",filetypes=(("All Files","*.*"),("Excel Files",".xlsx")))
        if op != None:
            data = pd.read_excel(op.name)
            if "Email" in data.columns:
                self.emails=list(data["Email"])
                c=[]
                for i in self.emails:
                    if pd.isnull(i) == False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state="readonly")
                    self.lbl_total.config(text="Total: "+str(len(self.emails)))
                    self.lbl_sent.config(text="SENT: ")
                    self.lbl_left.config(text="LEFT: ")
                    self.lbl_failed.config(text="FAILED: ")
                else:
                    messagebox.showerror("Error","This file has no email address",parent=self.root)

            else:
                messagebox.showerror("Error","Please select file which have Email Columns",parent=self.root)


    def send_email(self):
        x= len(self.txt_msg.get("1.0",END))
        if self.txt_to.get() == "" or self.txt_subj.get() == "" or x == 1:
            messagebox.showerror("Error","All fields are required",parent=self.root)
        else:
            if self.var_choice.get()=="single":
                status=email_function.email_send(self.txt_to.get(),self.txt_subj.get(),self.txt_msg.get("1.0",END),self.from_,self.pass_)
                if status == "s":
                    messagebox.showinfo("Success","Email has been sent!",parent=self.root)
                if status == "f":
                    messagebox.showerror("Failed","Email has not been sent, Try Again",parent=self.root)

            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send(x,self.txt_subj.get(),self.txt_msg.get("1.0",END),self.from_,self.pass_)
                    if status == "s":
                        self.s_count+=1
                    if status == "s":
                        self.f_count+=1
                    self.status_bar()
                messagebox.showinfo("Success","The emails were sent, Please Check Status!",parent=self.root)

    def status_bar(self):
        self.lbl_total.config(text="STATUS: "+str(len(self.emails))+"=>>")
        self.lbl_sent.config(text="SENT: "+str(self.s_count))
        self.lbl_left.config(text="LEFT: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="FAILED: "+str(self.f_count))

    def check_single_or_bulk(self):
        if self.var_choice.get() == "single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.clear1()
        if self.var_choice.get() == "bulk":
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state="readonly")

    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_subj.delete(0,END)
        self.txt_msg.delete("1.0",END)
        self.var_choice.set("single")
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")

    def setting_win(self):
        self.check_file_exist()
        self.root2=Toplevel()
        self.root2.title("Settings")
        self.root2.geometry("700x350+330+140")
        self.root2.resizable(False,False)
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="white")
        self.root2.iconbitmap("images/config.ico")

        title2=Label(self.root2,text="Credentials Setting",image=self.setting_icon,padx=10,compound=LEFT,font=("Tahoma",38,"bold"),bg="#1289A7",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        desc2=Label(self.root2,text="Enter the Email address and password from which to send all emails.",font=("calibri (body)",14),bg="#f6e58d",fg="black").place(x=0,y=70,relwidth=1)

        from_=Label(self.root2,text="Email Address",font=("times new roman",18),bg="white").place(x=50,y=150)
        pass_=Label(self.root2,text="PASSWORD",font=("times new roman",18),bg="white").place(x=50,y=200)

        self.txt_from=Entry(self.root2,font=("times new roman",14),bg="#c8d6e5")
        self.txt_from.place(x=250,y=150,width=330,height=30)

        self.txt_pass=Entry(self.root2,font=("times new roman",14),bg="#c8d6e5",show="*")
        self.txt_pass.place(x=250,y=200,width=330,height=30)

        self.btn_visibility=Button(self.root2,image=self.pass_icon,command=self.pass_visibility,bd=0,bg="white",fg="white",activebackground="white",activeforeground="white",cursor="hand2")
        self.btn_visibility.place(x=585,y=190,width=50,height=45)

        btn_clear=Button(self.root2,command=self.clear2,text="CLEAR",font=("Tahoma",18,"bold"),bg="#30336b",fg="white",activebackground="#30336b",activeforeground="white",cursor="hand2").place(x=285,y=260,width=120,height=35)
        btn_save=Button(self.root2,text="SAVE",command=self.save_setting,font=("Tahoma",18,"bold"),bg="#32ff7e",fg="white",activebackground="#32ff7e",activeforeground="white",cursor="hand2").place(x=425,y=260,width=120,height=35)
        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)

    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)

    def pass_visibility(self):
        if self.txt_pass["show"] == "*":
            self.txt_pass["show"] = ""
            self.btn_visibility.config(image=self.pass2_icon)
        else:
            self.txt_pass["show"] = "*"
            self.btn_visibility.config(image=self.pass_icon)

    def check_file_exist(self):
        if os.path.exists("important.txt") == False:
            f=open("important.txt","w")
            f.write(",")
            f.close()
        f2=open("important.txt","r")
        self.credentials=[]
        for i in f2:
            self.credentials.append([i.split(",")[0],i.split(",")[1]])
        self.from_=self.credentials[0][0]
        self.pass_=self.credentials[0][1]

    def save_setting(self):
        if self.txt_from.get() == "" or self.txt_pass.get() == "":
            messagebox.showerror("Error","All fields are required",parent=self.root2)
        else:
            f=open("important.txt","w")
            f.write(self.txt_from.get()+","+self.txt_pass.get())
            f.close()
            messagebox.showinfo("Success","The e-mail address and password have been saved!")

root=Tk()
obj=bulk_email(root)
root.mainloop()