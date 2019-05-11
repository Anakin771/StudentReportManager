from openpyxl.styles import *

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import *

import tkcalendar

import datetime

import sheetUtil

dateframe_ = None


class PickDate:
    font_family = "TH Sarabun New"
    easy_read = (font_family, '18')
    input_font = (font_family, '14')
    button_font = (font_family, '12', 'bold')
    h1_font = (font_family, '32', 'bold')
    h2_font = (font_family, '22', 'bold')
    h3_font = (font_family, '20', 'bold')
    hyperlink_font = (font_family, '14', 'bold underline')
    small_para = (font_family, '14')

    def __init__(self, root_window, semester, aff, preset_sheet=None, callback_func=None, startup_date=None):
        """ Constructor """

        if startup_date is None:
            startup_date = []

        if aff in "ประถมศึกษา":
            self.is_mid = False
        else:
            self.is_mid = True
        self.root = root_window
        self.selected_date = []
        self.excel_date = []
        self.dateframe_list = []
        self.enable_classnum = BooleanVar()
        self.callback_func = callback_func
        self.startup_date = startup_date
        self.semester = semester
        self.full_dateframe = None
        self.preset_sheet = preset_sheet

        # ====================================================================
        # =============================== GUI ================================
        # ====================================================================

        # ==================== Main Box ====================
        self.main = tk.Toplevel(self.root)
        self.main.title("Date Picker")

        # ==================== Calendar ====================

        top_wrapper = tk.Frame(self.main)
        top_wrapper.pack()

        self.calendar = tkcalendar.Calendar(top_wrapper, firstweekday='sunday')
        self.calendar.pack(side=LEFT, padx=10, ipadx=50, pady=20, expand=True, fill='both')

        # ==================== Side Box ====================
        side_box = tk.Frame(top_wrapper)
        side_box.pack(side=LEFT, padx=10)

        # --------------- Top-Side ---------------
        top_text = tk.Frame(side_box)
        top_text.pack(pady=10)
        tk.Label(top_text, text="จำนวนวันที่เลือกแล้ว:", font=self.easy_read, justify=CENTER). \
            pack(side=LEFT, padx=5)
        self.num_of_days = tk.Label(top_text, text=str(0) + " / 22", relief=SUNKEN, font=self.easy_read, justify=CENTER)
        self.num_of_days.pack(side=LEFT, ipadx=10, padx=5)

        # --------------- Selection Box ---------------
        sel_box = tk.Frame(side_box)
        sel_box.pack(padx=25)

        # - Selection List & Scroll Bar -
        self.sel_list = tk.Listbox(sel_box, height=7, width=40, highlightthickness=0, selectmode=BROWSE,
                                   font=self.input_font)
        scrb = tk.Scrollbar(sel_box, orient=VERTICAL, command=self.sel_list.yview)
        self.sel_list.configure(yscrollcommand=scrb.set)
        self.sel_list.pack(side=LEFT)
        scrb.pack(anchor="e", side=LEFT, expand=True, fill="y")

        self.sel_list.bind("<<ListboxSelect>>", lambda e: self.enable_remove())

        # - Buttons -
        sel_btn_frame = tk.Frame(side_box)
        sel_btn_frame.pack(pady=10, expand=True, fill='x')
        self.confirm_btn = ttk.Button(sel_btn_frame, text="Confirm", command=self.confirm_date, state=DISABLED)
        self.confirm_btn.pack(side=LEFT, expand=True)
        self.remove_btn = ttk.Button(sel_btn_frame, text="Remove", state=DISABLED)
        self.remove_btn.pack(side=LEFT, expand=True, padx=5)
        ttk.Button(sel_btn_frame, text="Clear", command=self.remove_all_date) \
            .pack(side=LEFT, expand=True, padx=5)
        ttk.Button(sel_btn_frame, text="Cancel").pack(side=LEFT, expand=True)

        # ==================== Class Number by date ====================

        # --------------- Toggle Checkbox ---------------

        bottom_box = tk.LabelFrame(self.main, text="กำหนดชั่วโมงเรียนของแต่ละวัน", font=self.easy_read)
        bottom_box.pack(padx=50, ipadx=10, pady=10, ipady=15)

        checkbox_left = tk.Frame(bottom_box)
        checkbox_left.pack(side=LEFT, expand=True, fill='y', padx=20)
        self.check_enable_classnum = tk.Checkbutton(checkbox_left, text="เปิดใช้การกำหนดชั่วโมงเรียนต่อวัน",
                                                    variable=self.enable_classnum,
                                                    command=self.toggle_days_assignment, onvalue=True, offvalue=False)
        self.check_enable_classnum.pack()
        tk.Label(checkbox_left, text="ด้วยการเปิดการใช้งานคณสมบัตินี้ ในแต่ละวัน\n"
                                     "จะมีเลขคายกำกับไว้ในแผ่นงานด้วย\n "
                                     "(ใช้เฉพาะกับแผ่นงานระดับมัฐยมเท่านั้น)",
                 font=self.small_para, relief=RIDGE, justify=LEFT).pack()

        # --------------- Days Assignment ---------------

        checkbox_right = tk.Frame(bottom_box)
        checkbox_right.pack(side=LEFT, expand=True, fill=BOTH, ipadx=10, padx=5)

        self.day_list = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"]
        self.day_assignment = []

        for day in self.day_list:
            new_box = tk.Frame(checkbox_right)
            new_box.pack(side=LEFT, padx=5, ipadx=5)
            tk.Label(new_box, text=day).pack(side=LEFT, padx=1)
            new_day = ttk.Spinbox(new_box, from_=1, to=12, width=3, state=DISABLED)
            new_day.pack(side=LEFT)
            self.day_assignment.append(new_day)

        # ====================================================================
        # =========================== END GUI ================================
        # ====================================================================

        self.calendar.bind("<<CalendarSelected>>", lambda event: self.calendar_select_callback())

    def calendar_select_callback(self):
        """ Event: When selected a date on the calendar, add and update the date list. """
        selected = self.calendar.selection_get()
        invalid = False

        # Check wether a duplicate, weekend date, or out-of-semester scope date has been selected.
        if selected.strftime("%a") in "Sat Sun":
            invalid = True
            warning_msg = "วันที่นี้ตรงกับวันหยุดเสาร์-อาทิตย์"
        elif int(selected.strftime("%Y")) > int(self.semester) - 543 and int(selected.strftime("%m")) > 3:
            invalid = True
            warning_msg = "ปีของวันที่นี้ไม่ได้อยู่ในขอบเขตปีการศึกษาที่ท่านระบุ"
        else:
            for date in self.selected_date:
                if date == selected:
                    invalid = True
                    break
            warning_msg = "ท่านได้เลือกวันที่นี้แล้ว"

        if not invalid:
            # Add the date into a list (with crucial data extracted)
            excel_date_info = (selected.strftime("%d"), selected.strftime("%B"), selected.strftime("%m"),
                               selected.strftime("%Y"), selected.strftime("%a"))
            self.selected_date.append(selected)
            self.excel_date.append(excel_date_info)
            self.selected_date.sort()
            self.excel_date.sort(key=lambda e: datetime.datetime(int(e[-2]), int(e[-3]), int(e[0])))
            self.listbox_update()
        else:
            messagebox.showwarning("ไม่สามารถเพี่มวันที่ได้", "ไม่สามารถเพี่มวันที่นี้ได้\n" + warning_msg)

    def listbox_update(self):
        """ Update the Listbox interface (Called within methods only!) """
        self.num_of_days.configure(text=str(len(self.selected_date)) + " / 22")
        self.sel_list.delete(0, END)
        for date in self.selected_date:
            date_str = sheetUtil.weekday_to_thai[date.strftime("%a")] \
                       + " / " + \
                       str(int(date.strftime("%d"))) + " " + sheetUtil.month_to_thai[date.strftime("%B")] \
                       + " / " + \
                       str(int(date.strftime("%Y")) + 543)
            self.sel_list.insert(END, date_str)
        if len(self.selected_date) >= 22:
            self.calendar.configure(state='disabled')
        elif len(self.selected_date) >= 10:
            self.confirm_btn.config(state='normal')
        else:
            self.confirm_btn.config(state='disabled')
            self.calendar.configure(state='normal')

    def toggle_days_assignment(self, state=None):
        """
        Event: Triggers when toggle 'Enable Class Hour Assignment' Checkbox;
        Toggle the state of each of the Spin-boxes.
        :param state: Optional; set the state of those widget to that of this parameter.
        :return: Nothing, but result in change of availability of the class number selection.
        """
        if state is None:
            if self.enable_classnum.get():
                for day_spinbox in self.day_assignment:
                    day_spinbox.configure(state=NORMAL)
            else:
                for day_spinbox in self.day_assignment:
                    day_spinbox.configure(state=DISABLED)
        else:
            for day_spinbox in self.day_assignment:
                day_spinbox.configure(state=state)

    @staticmethod
    def translate_date_dict(date_dict):
        """
        Translate dateframe dictionary data into a simpler one.

        :param date_dict: the dictionary of the date frame.
        :return: A list of date & month list as follows:
                [ [M1, D1], [M1, D2], ... [M1, Dn], ... [Mn, Dn] ]
        """
        result = []
        for month, dates in date_dict.items():
            for date in dates:
                result.append([month, date])
        return result

    def confirm_date(self):
        """
        Event: Triggers when click on 'Confirm' button;
        Translate selected data into usable list of dictionaries.
        """
        current_month = ""
        current_date_list = []
        current_classnum_list = []
        current_dateframe = {}
        date_assignment = {
            "Mon": self.day_assignment[0].get(),
            "Tue": self.day_assignment[1].get(),
            "Wed": self.day_assignment[2].get(),
            "Thu": self.day_assignment[3].get(),
            "Fri": self.day_assignment[4].get()
        }

        for ix, date in enumerate(self.excel_date):
            if self.excel_date.index(date) == 0 or self.excel_date[int(self.excel_date.index(date) - 1)][1] == date[1]:
                current_month = date[1]
                current_date_list.append(date[0])
                if self.is_mid:
                    current_classnum_list.append(date_assignment.get(date[-1]))
            else:
                current_dateframe["month"] = current_month
                current_dateframe["dates"] = current_date_list.copy()
                if self.is_mid and self.enable_classnum.get():
                    current_dateframe["class_num"] = current_classnum_list.copy()
                completed_dateframe = current_dateframe.copy()
                self.dateframe_list.append(completed_dateframe)

                current_date_list.clear()
                current_classnum_list.clear()

                current_month = date[1]
                current_date_list.append(date[0])
                if self.is_mid and self.enable_classnum.get():
                    current_classnum_list.append(date_assignment.get(date[-1]))

        current_dateframe["month"] = current_month
        current_dateframe["dates"] = current_date_list
        if self.is_mid and self.enable_classnum.get():
            current_dateframe["class_num"] = current_classnum_list
        self.dateframe_list.append(current_dateframe)

        global dateframe_
        dateframe_ = self.dateframe_list.copy()
        self.full_dateframe = dateframe_.copy()
        self.main.destroy()

    def enable_remove(self):
        """ Event : When select a date on a Listbox, enable Remove button and bind it to removal event."""
        selection = self.sel_list.curselection()
        self.remove_btn.configure(state=ACTIVE)
        self.remove_btn.bind("<Button-1>", lambda ev: self.remove_date(selection))

    def remove_date(self, selection):
        """ Event : When Remove button is pressed, remove te selected date"""
        for idx in selection:
            self.selected_date.pop(idx)
            self.excel_date.pop(idx)
        self.listbox_update()

    def remove_all_date(self):
        """ Event : When Clear button is pressed, remove all date in the date list."""
        self.selected_date.clear()
        self.excel_date.clear()
        self.listbox_update()
