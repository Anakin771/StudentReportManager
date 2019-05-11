
from tkinter import *

import _mainMenu
import _newSheet

"""
****************************************
*********  CONSTANT VARIABLES  *********
****************************************
"""
# ************** General **************
UNINAME = "Student Attendant Manager"
VERSION = "0.0.1 (Alpha)"

# ************* Main Menu *************
MAIN_WIDTH = 1000
MAIN_HEIGHT = 600

# ************* New File *************
NEW_WIDTH = 800
NEW_HEIGHT = 600

root = Tk()
new_gui = _newSheet.NewFile(root, VERSION, NEW_WIDTH, NEW_HEIGHT)
main_gui = _mainMenu.MainMenu(root, UNINAME, VERSION, MAIN_WIDTH, MAIN_HEIGHT, new_gui)
new_gui.link_to_main(main_gui)
root.mainloop()
