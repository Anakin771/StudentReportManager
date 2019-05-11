from openpyxl.styles import *

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import *

import tkcalendar

import datetime

import sheetUtil
import datePicker
import _editor

black_side = Side(border_style="thin", color="000000")
normal_edge = Border(left=black_side,
                     right=black_side,
                     bottom=black_side,
                     top=black_side)
dateframe_ = []


class NewFile:
    font_family = "TH Sarabun New"
    easy_read = (font_family, '18')
    input_font = (font_family, '15')
    button_font = (font_family, '12', 'bold')
    h1_font = (font_family, '32', 'bold')
    h2_font = (font_family, '22', 'bold')
    h3_font = (font_family, '20', 'bold')
    hyperlink_font = (font_family, '14', 'bold underline')

    def __init__(self, master, version, canvas_width, canvas_height, main_menu=None):

        """ -------- Constructor -------- """

        self.version = version
        self.canvas_width = canvas_width
        self.canvas_height = canvas_height
        self.master = master
        self.main_menu = main_menu
        self.direc = ""

        self.editor = None
        self.date_picker = None

        def clear_entries():
            """ Clear all entry in the form. """
            self.form_file_name.delete(0, END)
            self.direc = ""

            dir_txt.config(text="None", justify=CENTER)
            self.grade.delete(0, END)
            self.room.delete(0, END)
            self.affiliate.current(0)
            self.semester.delete(0, END)
            self.total_hour.delete(0, END)
            self.subject.delete(0, END)
            self.teacher_name.delete(0, END)

        def cancel_entry():
            """ Cancel "New File" action and return to main menu. """
            clear_entries()
            self.hide()
            self.main_menu.show()

        def choose_dir():
            """ Event: Select a directory and returns it"""
            self.direc = tk.filedialog.askdirectory()
            dir_txt.config(text=self.direc, justify=LEFT)

        #  ====================== ====================== ========= ====================== ======================
        #  ====================== ====================== START GUI ====================== ======================
        #  ====================== ====================== ========= ====================== ======================

        self.canvas = tk.Canvas(master, highlightthickness=0, width=self.canvas_width, height=self.canvas_height)
        self.canvas.pack()

        #  ====================== ====================== ====== ====================== ======================
        #  ====================== ====================== HEADER ====================== ======================
        #  ====================== ====================== ====== ====================== ======================

        header = tk.Frame(self.canvas)
        header.place(relwidth=0.8, relx=0.1, rely=0.05)

        tk.Label(header, text="สร้างไฟล์ใหม่", font=self.h1_font).pack()
        tk.Label(header, text="โปรดกรอกข้อมูลต่อไปนี้ให้ครบถ้วนและถูกต้อง",
                 font=self.h2_font).pack()

        #  ====================== ====================== ==== ====================== ======================
        #  ====================== ====================== FORM ====================== ======================
        #  ====================== ====================== ==== ====================== ======================

        form = tk.LabelFrame(self.canvas, text="รายละเอียดของไฟล์")
        form.place(relwidth=0.8, relx=0.1, relheight=0.7, rely=0.25)

        #  ================= ================= FILE NAME ================= =================

        form_name = tk.Frame(form)
        tk.Label(form_name, text="ชื่อไฟล์: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.form_file_name = tk.Entry(form_name, font=self.input_font)
        self.form_file_name.pack(side=LEFT, padx=6, expand=True, fill='x')
        tk.Label(form_name, text=".xlsx", font=self.easy_read).pack(side=LEFT)

        #  ================= ================= DIRECTORY ================= =================

        form_dir = tk.Frame(form)
        tk.Label(form_dir, text="ที่อยู่ของไฟล์: ", font=self.easy_read).pack(side=LEFT, padx=10)
        dir_txt = tk.Label(form_dir, text="ไม่มี", font=self.input_font)
        directory = ttk.Button(form_dir, text="เลือกที่อยู่", command=choose_dir, width=15)

        dir_txt.config(relief=SUNKEN)

        dir_txt.pack(side=LEFT, expand=True, fill='x', padx=10)
        directory.pack(side=LEFT)

        #  ================= ================= SEMESTER ================= =================
        form_semester = tk.Frame(form)
        tk.Label(form_semester, text="ภาคเรียนที่: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.term = ttk.Spinbox(form_semester, from_=1, to=100, font=self.input_font, width=8)
        self.term.pack(side=LEFT, padx=15)
        tk.Label(form_semester, text="ปีการศึกษา (พ.ศ.): ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.semester = ttk.Spinbox(form_semester, from_=2000, to=9999, font=self.input_font, width=8)
        self.semester.pack(side=LEFT, padx=15)

        #  ================= ================= CLASS ================= =================

        form_class = tk.Frame(form)
        tk.Label(form_class, text="ชั้น ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.grade = tk.Entry(form_class, width=4, font=self.input_font, justify=CENTER)
        self.grade.pack(side=LEFT, expand=True)
        tk.Label(form_class, text="/", font=self.easy_read).pack(side=LEFT, expand=True)
        self.room = ttk.Entry(form_class, width=4, font=self.input_font, justify=CENTER)
        self.room.pack(side=LEFT, expand=True)
        self.affiliate = ttk.Combobox(form_class, value=("ระดับชั้น", "ประถมศึกษา", "มัฐยมศึกษา"),
                                      font=self.input_font)
        self.affiliate.current(0)
        self.affiliate.pack(side=LEFT, padx=50)

        #  ================= ================= SUBJECT & SU ================= =================
        form_subject = tk.Frame(form)
        # Subject
        tk.Label(form_subject, text="วิชา: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.subject = ttk.Entry(form_subject, font=self.input_font, width=20)
        self.subject.pack(side=LEFT, padx=5)
        # Score Unit (SU)
        tk.Label(form_subject, text="จำนวน: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.s_unit = ttk.Spinbox(form_subject, font=self.input_font, width=10)
        self.s_unit.pack(side=LEFT, padx=5)
        tk.Label(form_subject, text="หน่วยกิต", font=self.easy_read).pack(side=LEFT, padx=10)

        #  ================= ================= TOTAL HOURS ================= =================
        form_hour = tk.Frame(form)
        tk.Label(form_hour, text="จำนวนชั่วโมง: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.total_hour = ttk.Spinbox(form_hour, from_=1, to=999, font=self.input_font, width=10)
        self.total_hour.pack(side=LEFT, padx=60, expand=True, fill='x')

        #  ================= ================= TEACHER'S NAME ================= =================
        form_teacher = tk.Frame(form)
        tk.Label(form_teacher, text="ชื่อครูผู้สอน: ", font=self.easy_read).pack(side=LEFT, padx=10)
        self.teacher_name = tk.Entry(form_teacher, font=self.input_font)
        self.teacher_name.pack(side=LEFT, expand=True, fill='x', padx=25)

        #  ================= ================= BUTTONS ================= =================
        form_btns = tk.Frame(form)
        ttk.Button(form_btns, text="ยืนยัน", command=self.confirm_entry) \
            .pack(side=LEFT, padx=20, expand=True, fill='x')
        ttk.Button(form_btns, text="ล้างการกรอก", command=clear_entries) \
            .pack(side=LEFT, padx=20, expand=True, fill='x')
        ttk.Button(form_btns, text="ยกเลิก", command=cancel_entry) \
            .pack(side=LEFT, padx=20, expand=True, fill='x')

        #  ================= ================= FORM : FINALIZE ================= =================

        form_name.pack(padx=50, pady=5, fill='x')
        form_dir.pack(padx=50, pady=5, fill='x')
        form_semester.pack(padx=50, pady=5, fill='x')
        form_class.pack(padx=50, pady=5, fill='x')
        form_subject.pack(padx=50, pady=5, fill='x')
        form_hour.pack(padx=50, pady=5, fill='x')
        form_teacher.pack(padx=50, pady=5, fill='x')
        form_btns.pack(padx=20, pady=10, fill='x')

        #  ====================== ====================== ======= ====================== ======================
        #  ====================== ====================== END GUI ====================== ======================
        #  ====================== ====================== ======= ====================== ======================

    def confirm_entry(self):
        """ Turn in form and new file from the template. """
        try:
            workbook_name = str(self.form_file_name.get())
            selected_dir = str(self.direc)
            title_grade = int(self.grade.get())
            title_room = int(self.room.get())
            title_aff = self.affiliate.get()
            title_semester = self.semester.get()
            desc_total_hour = int(self.total_hour.get())
            desc_subject = str(self.subject.get())
            sign_teacher_name = str(self.teacher_name.get())
        except ValueError:
            tk.messagebox.showerror("เกิดข้อผิดพลาด", "ข้อมูลที่ท่านระบุไม่สามารถนำไปใช้ได้\n"
                                                      "กรุณาระบุข้อมูลให้ถูกต้องด้วยค่ะ")
            print("A Value Error has occurred on submitting new file form.")
        else:
            if workbook_name == "" or selected_dir == "" or title_grade == "" or title_room == "" \
                    or title_aff == "ระดับขั้น" or title_semester == "" \
                    or desc_total_hour == "" or desc_subject == "" or sign_teacher_name == "":
                tk.messagebox.showerror("เกิดข้อผิดพลาด", "ท่านระบุข้อมูลไม่ครบ\nกรุณาระบุข้อมูลให้ครบถ้วนด้วยค่ะ")
                print("An incomplete form entry has occurred on submitting new file form.")
            else:
                self.date_picker = datePicker.PickDate(self.master, title_semester, title_aff)
                self.master.wait_window(window=self.date_picker.main)
                if self.date_picker.full_dateframe:
                    sheetUtil.generate_sheet(workbook_name, selected_dir,
                                             self.date_picker.full_dateframe, title_grade,
                                             title_room, title_aff, title_semester,
                                             desc_total_hour, desc_subject,
                                             sign_teacher_name)
                    self.close()
                    self.editor = _editor.Editor(master=self.master, file=selected_dir + "/" + workbook_name + ".xlsx",
                                                 main_obj=self.main_menu)

    def close(self):
        """ Close the entire app window. """
        self.master.withdraw()

    def open(self):
        """ Reopen the app window. """
        self.master.update()
        self.master.deiconify()

    def hide(self):
        """ Hide the entire canvas on call """
        self.canvas.pack_forget()

    def show(self):
        """ Show this canvas on call """
        self.canvas.pack()

    def link_to_main(self, sub_main_obj):
        """ Link to the main menu object outside of Constructor """
        self.main_menu = sub_main_obj
        self.hide()

    def link_to_editor(self, editor):
        """ Link to the main menu object outside of Constructor """
        self.editor = editor


if __name__ == '__main__':
    root = Tk()
    main = tk.Frame(root)
    new_file = NewFile(root, "TEST", 800, 600)
    root.mainloop()
