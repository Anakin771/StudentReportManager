from openpyxl import Workbook, load_workbook, cell
from openpyxl.styles import *

import tkinter as tk
import tkinter.font
from tkinter import ttk, filedialog, messagebox
from tkinter import *

import _editor

WIDTH = 1000
HEIGHT = 550

VERSION = "0.0.1 (Alpha)"


class MainMenu:

    font_family = "TH Sarabun New"
    easy_read = (font_family, '18')
    input_font = (font_family, '15')
    h1_font = (font_family, '32', 'bold')
    h2_font = (font_family, '22', 'bold')
    h3_font = (font_family, '20', 'bold')
    button_font = (font_family, '18', 'bold')
    hyperlink_font = (font_family, '14', 'bold underline')

    def __init__(self, master, uniname, version, canvas_width, canvas_height, new_obj):
        self.UNINAME = uniname
        self.version = version
        self.canvas_width = canvas_width
        self.canvas_height = canvas_height
        self.master = master
        self.new_obj = new_obj
        self.editor = None

        #  ====================== ====================== ========= ====================== ======================
        #  ====================== ====================== START GUI ====================== ======================
        #  ====================== ====================== ========= ====================== ======================

        master.title(self.UNINAME + " | v" + self.version.lower())

        self.master.protocol("WM_DELETE_WINDOW", self.quit)

        self.canvas = tk.Canvas(master, width=WIDTH, height=HEIGHT, highlightthickness=0)
        self.canvas.pack()

        #  ====================== ====================== ====== ====================== ======================
        #  ====================== ====================== HEADER ====================== ======================
        #  ====================== ====================== ====== ====================== ======================

        header = tk.Frame(self.canvas)
        header.place(relwidth=0.6, relx=0.2, rely=0.05)

        tk.Label(header, text=self.UNINAME, font=self.h1_font).pack()
        tk.Label(header, text="Version " + VERSION, font=self.h2_font).pack()

        ttk.Separator(self.canvas, orient=HORIZONTAL).place(rely=0.25, relx=0.1, relwidth=0.8)

        #  ====================== ====================== ======== ====================== ======================
        #  ====================== ====================== NEW FILE ====================== ======================
        #  ====================== ====================== ======== ====================== ======================

        new_frame = tk.Frame(self.canvas, borderwidth=5)
        tk.Label(new_frame, text="- - - - -    สร้างไฟล์ใหม่    - - - - -", font=self.h3_font).pack(anchor='n')
        new_btn = ttk.Button(new_frame, text="สร้าง", command=self.newfile)

        new_btn.pack(expand=True, fill='both', padx=25)
        new_frame.place(relwidth=0.7, relheight=0.2, relx=0.15, rely=0.3)

        ttk.Separator(self.canvas, orient=HORIZONTAL).place(rely=0.55, relx=0.1, relwidth=0.8)

        #  ====================== ====================== ========= ====================== ======================
        #  ====================== ====================== OPEN FILE ====================== ======================
        #  ====================== ====================== ========= ====================== ======================

        open_frame = tk.Frame(self.canvas, borderwidth=5)
        open_frame.place(relwidth=0.7, relx=0.15, relheight=0.2, rely=0.6)
        tk.Label(open_frame, text="- - - - -    เปิดไฟล์งาน    - - - - -", font=self.h3_font).pack()
        ttk.Button(open_frame, text="เปิดไฟล์งาน", command=self.openfile).pack(expand=True, fill='both', padx=25)

        ttk.Separator(self.canvas, orient=HORIZONTAL).place(rely=0.85, relx=0.1, relwidth=0.8)

        #  ====================== ====================== ====== ====================== ======================
        #  ====================== ====================== FOOTER ====================== ======================
        #  ====================== ====================== ====== ====================== ======================

        about = tk.Label(self.canvas, text="เกี่ยวกับโปรแกรม",
                         font=self.hyperlink_font, foreground="blue", cursor="hand2")
        about.place(relx=0.1, rely=0.9)
        about.bind("<Button-1>", self.about_callback)

        copy_lbl = tk.Label(self.canvas, text="© 2019 - Yanakorn Chaeyprasert")
        copy_lbl.place(relx=0.725, rely=0.9)

        #  ====================== ====================== ======= ====================== ======================
        #  ====================== ====================== END GUI ====================== ======================
        #  ====================== ====================== ======= ====================== ======================

    def newfile(self):
        self.hide()
        self.new_obj.show()

    def openfile(self):
        file = tk.filedialog.askopenfilename(initialdir="Document", title="Select file",
                                             filetypes=[("xlsx files", "*.xlsx")])
        if file != "":
            try:
                load_workbook(file)
            except FileNotFoundError:
                tk.messagebox.showerror("เกิดข้อผิดพลาด", "ไม่สามารถเปิดไฟล์ของท่านได้\nโปรดลองอีกครั้งภายหลัง, "
                                                          "ตรวจสอบไฟล์ด้วยตนเอง, หรือติดต่อผู้ที่เกี่ยวข้อง")
                print("A FilNotFoundError has occurred when opening file in _mainMenu.py")
            else:
                self.close()
                self.editor = _editor.Editor(master=self.master, file=file, main_obj=self)

    def about_callback(self, event):
        print(str(self))
        print(str(event))

    def quit(self):
        if messagebox.askokcancel("ออก", "ออกจากโปรแกรมหรือไม่?"):
            self.master.destroy()

    def close(self):
        """ Close the entire app window. """
        self.master.withdraw()

    def open(self):
        """ Reopen the app window. """
        self.master.update()
        self.master.deiconify()
        self.show()
        self.new_obj.hide()

    def hide(self):
        """ Hide the entire canvas on call. """
        self.canvas.pack_forget()

    def show(self):
        """ Show this canvas on call. """
        self.canvas.pack()
