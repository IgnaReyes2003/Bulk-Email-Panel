from tkinter import*

class bulk_email:
    def __init__(self,root):
        self.root=root
        self.root.title("Bulk Email")
        self.root.geometry("1000x450+165+100")

root=Tk()
obj=bulk_email(root)
root.mainloop()