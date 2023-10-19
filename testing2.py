from tkinter import *

root = Tk()
root.title("Test")
root.geometry("800x150")
root.resizable(False, False)


def show():
    global top
    global entry1
    top = Toplevel(root)
    top.geometry("200x200")
    button2 = Button(top, text="send mail", command=hide)
    button2.grid(row=0, column=1)

    entry1 = Entry(top)
    entry1.grid(row=0, column=0)


def hide():
    print(entry1.get())
    top.withdraw()


button1 = Button(root, text="testing clicking", command=show)
button1.grid(row=0, column=0)
root.mainloop()
