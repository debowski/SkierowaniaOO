# import sys
import tkinter as tk
# import ttkbootstrap as ttk


from pathlib import Path
from tkinter import PhotoImage
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

# bootstyle="inverse-success", foreground="black"

# from main_window import Application


class Application(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill=BOTH, expand=YES)
        self.create_widgets()

    def create_widgets(self):

        self.col1 = ttk.Frame(self, padding=10, bootstyle="success",
                              borderwidth=2, relief='groove')
        self.col1.grid(row=0, column=0, sticky="nsew")

        self.wyb_plik = ttk.Button(
            self.col1, text="Plik...", command=self.say_hi)
        self.wyb_plik.grid(row=0, column=0, sticky="w")

        self.col1.columnconfigure(0, weight=1)

    def say_hi(self):
        print("hi there, everyone!")


if __name__ == "__main__":
    root = ttk.Window(title="v0.22")
    app = Application(root)
    app.mainloop()
