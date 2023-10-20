from tkinter import *
import gspread
from tkinter import filedialog
from fpdf import FPDF
from fpdf.fonts import FontFace
from tkinter import ttk
from dotenv import dotenv_values
import json
import smtplib, ssl

config = dotenv_values(".env")

sa = gspread.service_account_from_dict(config)
sh = sa.open("Testing").sheet1


class Window(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()
        self.TABLE_DATA = []
        self.CELL_WIDTH = 20
        self.file_path = ""

    def init_window(self):
        global entry1
        global entry2
        global route
        global success1
        global success3
        global error_str
        global ProgressBar

        error_str = StringVar()
        success1 = StringVar()
        success3 = StringVar()
        route = StringVar()
        success1.set("")
        success3.set("")
        error_str.set("")

        f = open("file_path_chosen.txt", "r")
        self.file_path = f.readline().strip("\n")
        route.set(self.file_path)
        f.close()

        Label1 = Label(root, textvariable=success1)
        Label1.grid(row=0, column=1)
        Label3 = Label(root, textvariable=route)
        Label3.grid(row=4, column=0)
        Label4 = Label(root, textvariable=error_str)
        Label4.grid(row=0, column=1)
        Label5 = Label(root, textvariable=success3)
        Label5.grid(row=1, column=3)

        entry1 = Entry(root)
        entry1.grid(row=0, column=0)
        entry2 = Entry(root)
        entry2.grid(row=1, column=0)

        btn = Button(root, text="Wybierz miejsce zapisu pliku", command=self.save)
        btn.grid(row=3, column=0, padx=10)
        button1 = Button(
            root, text="Kliknij aby znalezc dane", command=self.find_cell_data
        )
        button1.grid(row=3, column=1, padx=10)

        button2 = Button(
            root, text="Kliknij aby wygenerowac fakture", command=self.create_pdf
        )
        button2.grid(row=3, column=2, padx=10)

        button3 = Button(root, text="Ile rzedow do faktury", command=self.get_row_end)
        button3.grid(row=1, column=1, padx=10)

        button4 = Button(root, text="Wyslij email", command=self.show_email_window)
        button4.grid(row=5, column=0)

        ProgressBar = ttk.Progressbar(
            root, orient=HORIZONTAL, length=100, mode="determinate"
        )
        ProgressBar.grid(row=0, column=3)

    def get_row_end(self):
        self.row_end = int(entry2.get())
        success3.set(f"Wybrano zakres: {self.row_end}")

    def find_cell_data(self):
        try:
            success1.set("Szukanie Danych")
            cell1 = sh.find(str(entry1.get()))
            row1 = cell1.row

            TABLE_DATA = []
            for i in range(0, self.row_end + 1):
                TABLE_DATA.append(sh.row_values(row1 + i))
                ProgressBar["value"] += self.row_end + 3
                root.update_idletasks()

            error_str.set("")
            success1.set("Znaleziono dane!")

            self.TABLE_DATA = TABLE_DATA

        except AttributeError:
            error_str.set("Nie znaleziono wpisanej danej,\nlub nie wpisano zakresu")
        except UnboundLocalError:
            error_str.set("Nie znaleziono wpisanej danej")

    def save(self):
        global file_path
        file_path = filedialog.asksaveasfilename()
        with open("file_path_chosen.txt", "w") as f:
            f.write(file_path)
        route.set(file_path)
        f.close()

    def get_num_of_lines_in_multicell(self, pdf, message):
        # divide the string in words

        words = message.split(" ")
        line = ""
        n = 1
        for word in words:
            line += word + " "
            line_width = pdf.get_string_width(line)
            if line_width > (self.CELL_WIDTH - 1) * 2:
                # the multi_cell() insert a line break
                n += 2
                # reset of the string
                line = word + " "
            elif line_width > self.CELL_WIDTH - 1:
                # the multi_cell() insert a line break
                n += 1
                # reset of the string
                line = word + " "
        return n

    def get_num_lines_max(self, pdf, row):
        num_lines_max = 0
        for item in row:
            num_lines = self.get_num_of_lines_in_multicell(pdf, item)
            if num_lines > num_lines_max:
                num_lines_max = num_lines
        return num_lines_max

    def create_pdf(self, spacing=2):
        try:
            pdf = FPDF(orientation="L")

            pdf.add_font("Times New Roman", "", r"C:\Windows\Fonts\timesbd.ttf")
            pdf.set_font("Times New Roman")

            pdf.add_page()
            counter = 0

            col_width = pdf.w / 9.9
            self.CELL_WIDTH = col_width
            row_height = pdf.font_size + 1

            for row in self.TABLE_DATA:
                temp = True
                counter += 1
                new_x = 10
                saveY = 0
                num_lines_max = self.get_num_lines_max(pdf, row)
                multicell_height = row_height * num_lines_max

                if counter == 1:
                    for item in row:
                        pdf.set_xy(new_x, 10)
                        pdf.set_fill_color(128, 128, 128)
                        num_lines = self.get_num_of_lines_in_multicell(pdf, item)
                        if item == "":
                            pdf.multi_cell(
                                col_width,
                                row_height,
                                txt=item,
                                border=0,
                                align="C",
                                fill=False,
                            )
                        elif num_lines == 1:
                            pdf.multi_cell(
                                col_width,
                                row_height * 3,
                                txt=item,
                                border="T",
                                align="C",
                                fill=True,
                            )
                        elif num_lines == 2:
                            pdf.multi_cell(
                                col_width,
                                (row_height * 3) / 2,
                                txt=item,
                                border="T",
                                align="C",
                                fill=True,
                            )
                        else:
                            pdf.multi_cell(
                                col_width,
                                row_height,
                                txt=item,
                                border="T",
                                align="C",
                                fill=True,
                            )
                        if pdf.get_y() > saveY:
                            saveY = pdf.get_y()
                        new_x += col_width
                        pdf.set_xy(new_x, 10)
                    pdf.set_xy(10, saveY)

                else:
                    num_lines_max = self.get_num_lines_max(pdf, row)
                    if num_lines_max == 1:
                        for item in row:
                            if item == "":
                                pdf.cell(
                                    col_width,
                                    row_height * spacing,
                                    txt=item,
                                    border="T",
                                    align="C",
                                    fill=False,
                                )
                            else:
                                pdf.cell(
                                    col_width,
                                    row_height * spacing,
                                    txt=item,
                                    border="T",
                                    align="C",
                                )
                        pdf.ln(row_height * spacing)
                    else:
                        multicell_height = row_height * num_lines_max
                        y_multicell_row = pdf.get_y()
                        for item in row:
                            if item == "":
                                pdf.multi_cell(
                                    col_width,
                                    row_height,
                                    txt=item,
                                    border=0,
                                    align="C",
                                    fill=False,
                                )
                            else:
                                num_lines = self.get_num_of_lines_in_multicell(
                                    pdf, item
                                )
                                pdf.multi_cell(
                                    col_width,
                                    multicell_height / num_lines,
                                    txt=item,
                                    border="T",
                                    align="C",
                                    fill=False,
                                )
                            if pdf.get_y() > saveY:
                                saveY = pdf.get_y()
                            new_x += col_width
                            pdf.set_xy(new_x, y_multicell_row)
                        pdf.set_xy(10, saveY)

            pdf.output(f"{route.get()}")
            ProgressBar["value"] = 0
            error_str.set("")
            success1.set("Wygenerowano Fakture!")
        except NameError:
            error_str.set("Nie wybrano ściezki")

    def show_email_window(self):
        global top
        global email_entry1
        global email_entry2
        global password_entry

        top = Toplevel(root)
        top.geometry("200x200")

        label1 = Label(top, text="Nadawca")
        label1.grid(row=0, column=0)
        label2 = Label(top, text="Odbiorca")
        label2.grid(row=1, column=0)
        label3 = Label(top, text="Hasło")
        label3.grid(row=2, column=0)

        button1 = Button(top, text="Wyslij Email", command=self.hide_email_window)
        button1.grid(row=3, column=0)

        email_entry1 = Entry(top)
        email_entry1.grid(row=0, column=1)
        email_entry2 = Entry(top)
        email_entry2.grid(row=1, column=1)
        password_entry = Entry(top)
        password_entry.grid(row=2, column=1)

    def hide_email_window(self):
        port = 465  # For SSL
        smtp_server = "smtp.gmail.com"
        sender_email = email_entry1.get()
        receiver_email = email_entry2.get()
        password = password_entry.get()

        message = """\
        Subject: Hi there

        This message is sent from Python."""

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message)

        top.withdraw()


root = Tk()
root.title("Generator Faktur")
root.geometry("800x150")
root.resizable(False, False)
app = Window(root)
root.mainloop()
