#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from attendance_gui import AttendanceGUI

def main():
    root = tk.Tk()
    app = AttendanceGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
    