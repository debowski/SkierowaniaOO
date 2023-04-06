# import sys
# import tkinter as tk
# import ttkbootstrap as ttk


from pathlib import Path
from tkinter import PhotoImage
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox


# from main_window import Application


class Application(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        # self.parent.geometry("300x600")
        # self.parent.resizable(False, False)

        self.pack()
        self.create_widgets()

    def create_widgets(self):
        # dodajemy wagę do kolumny 0
        self.columnconfigure(0, weight=1, minsize=300)

        # kolumna 1
        col1 = ttk.Frame(self, padding=10)
        col1.grid(row=0, column=0, sticky=NSEW)

        self.wyb_plik = ttk.Button(
            col1, text="Plik...", command=self.say_hi)
        # pierwszy argument to siatka, pierwszy to rozmiar
        self.wyb_plik.pack(side=TOP, fill=BOTH, expand=YES)

    def say_hi(self):
        print("hi there, everyone!")


if __name__ == "__main__":
    root = ttk.Window(title="v0.22", themename="flatly", iconphoto="icon.png")
    app = Application(root)
    app.mainloop()
