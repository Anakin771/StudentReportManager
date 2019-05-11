"""
Here the TreeView widget is configured as a multi-column listbox
with adjustable column width and column-header-click sorting.
"""
import tkinter as tk
import tkinter.font as tkFont
import tkinter.ttk as ttk
from tkinter.ttk import Style


class TreeviewTable(object):
    """use a ttk.TreeView as a multicolumn ListBox"""

    def __init__(self, master, header=None, content=None, max_height=10, enable_vsb=False, enable_hsb=False,
                 odd_row_color="white"):
        if header is None:
            header = ["Column"]
        if content is None:
            content = []
        self.tree = None
        self.master = master
        self.header = header.copy()
        self.content = content.copy()
        self.max_height = max_height
        self.enable_vsb = enable_vsb
        self.enable_hsb = enable_hsb
        self.odd_row_color = odd_row_color

        self.current_selection = None

        self._setup_widgets(self.master, self.enable_vsb, self.enable_hsb)
        self._build_tree(self.header, self.content, odd_row_color)

    def set_content(self, content):
        """Set the treeview's column content"""
        self.content = content.copy()
        self.tree.delete(*self.tree.get_children())

        for index, item in enumerate(content, 1):
            if index % 2 == 1:
                self.tree.insert('', 'end', values=item, tag="odd_row")
            else:
                self.tree.insert('', 'end', values=item)

            # adjust column's width if necessary to fit each value
            for ix, val in enumerate(item):
                col_w = tkFont.Font().measure(val)
                if self.tree.column(self.header[ix], width=None) < col_w:
                    self.tree.column(self.header[ix], width=col_w)

        # Set the color of the odd row item to the specified color
        self.tree.tag_configure("odd_row", background=self.odd_row_color)

    def _setup_widgets(self, master, enable_vscroll, enable_hscroll):
        """Called by __init__ ; sets up interface of th table"""
        self.container = ttk.Frame(master)
        # Create a Treeview scrollbars
        self.tree = ttk.Treeview(master=self.container, columns=self.header, show="headings", height=self.max_height,
                                 selectmode="browse")
        if enable_vscroll:
            vsb = ttk.Scrollbar(self.container, orient="vertical", command=self.tree.yview)
            vsb.grid(column=1, row=0, sticky='ns')
            self.tree.configure(yscrollcommand=vsb.set)
        if enable_hscroll:
            hsb = ttk.Scrollbar(self.container, orient="horizontal", command=self.tree.xview)
            hsb.grid(column=0, row=1, sticky='ew')
            self.tree.configure(xscrollcommand=hsb.set)
        self.tree.grid(column=0, row=0, sticky='nsew', in_=self.container)
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.selection_callback())
        self.tree.bind("<Button-1>", self.disable_resize_and_cursor)
        self.tree.bind("<Motion>", self.disable_resize_and_cursor)
        self.set_widget_state("active")

    def _build_tree(self, header, content, odd_row_color):
        """Called by __init__ ; assign value into header/column"""
        for col in header:
            self.tree.heading(col, text=col.title(), command=lambda c=col: self.sortby(c, 0))
            # Adjust the column's width to the header string
            self.tree.column(col, width=tkFont.Font().measure(col.title()))

        for index, item in enumerate(content, 1):
            if index % 2 == 1:
                self.tree.insert('', 'end', values=item, tag="odd_row")
            else:
                self.tree.insert('', 'end', values=item)

            # Adjust column's width if necessary to fit each value
            for ix, val in enumerate(item):
                col_w = tkFont.Font().measure(val)
                if self.tree.column(self.header[ix], width=None) < col_w:
                    self.tree.column(self.header[ix], width=col_w)

        # Set the color of the odd row item to the specified color
        self.tree.tag_configure("odd_row", background=odd_row_color)

    def sortby(self, col, descending):
        """sort tree contents when a column header is clicked on"""
        # grab values to sort
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        # if the data to be sorted is numeric change to float
        # data =  change_numeric(data)
        # now sort the data in place
        data.sort(reverse=descending)
        for ix, item in enumerate(data):
            self.tree.move(item[1], '', ix)
        # switch the heading so it will sort in the opposite direction
        self.tree.heading(col, command=lambda c=col: self.sortby(c, int(not descending)))

    def selection_callback(self):
        self.current_selection = self.tree.item(self.tree.selection())
        self.tree.focus(self.tree.selection())

    def disable_resize_and_cursor(self, event):
        """ Event; stops user from interfere with the well being of the header (except from sorting) """
        if self.tree.identify_region(event.x, event.y) == "separator":
            return "break"

    def set_col_width(self, width_list):
        for ix, width in enumerate(width_list):
            self.tree.column(self.header[ix], width=int(width*10))

    def set_widget_state(self, state):

        if state in "disabled":
            self.tree.state(("disabled",))
            self.tree.bind('<Button-1>', lambda e: 'break')
            self.tree.bind('<Motion>', lambda e: 'break')
        elif state in "normal" or state in "active":
            self.tree.state(("!disabled",))
            self.tree.bind("<Button-1>", self.disable_resize_and_cursor)
            self.tree.bind("<Motion>", self.disable_resize_and_cursor)

    def bind_tree(self, event="<Button-1>", command=lambda event: print(event)):
        self.tree.bind(event, command)


# Tester

if __name__ == '__main__':

    def this_tree_callback():
        treeview.selection_callback()
        res_text.config(text="Data: " + str(treeview.current_selection['values']))

    car_header = ['Car', 'Repair', 'Working Team']
    car_list = [
        ('Hyundai', 'brakes'),
        ('Honda', 'light'),
        ('Lexus', 'battery'),
        ('Benz', 'wiper'),
        ('Ford', 'tire'),
        ('Chrysler', 'piston'),
        ('Toyota', 'brake pedal'),
        ('BMW', 'seat'),
        ('Honda', 'light'),
        ('Lexus', 'battery'),
        ('Benz', 'wiper'),
        ('Ford', 'tire'),
        ('Chrysler', 'piston'),
        ('Toyota', 'brake pedal'),
        ('BMW', 'seat'),
        ('Honda', 'light'),
        ('Lexus', 'battery'),
        ('Benz', 'wiper'),
        ('Ford', 'tire'),
        ('Chrysler', 'piston'),
        ('Toyota', 'brake pedal'),
        ('BMW', 'seat'),
        ('Honda', 'light'),
        ('Lexus', 'battery'),
        ('Benz', 'wiper'),
        ('Ford', 'tire'),
        ('Chrysler', 'piston'),
    ]

    root = tk.Tk()
    root.title("Multicolumn Treeview/Listbox")
    treeview = TreeviewTable(root, header=car_header, enable_vsb=True)
    treeview.set_content(car_list)
    treeview.container.pack(expand=True, fill='x', padx=150)
    treeview.set_col_width([10, 25, 1])
    res_text = tk.Label(text="Data: ")
    res_text.pack()
    treeview.bind_tree("<<TreeviewSelect>>", lambda e: this_tree_callback())
    root.mainloop()
