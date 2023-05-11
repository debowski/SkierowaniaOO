# aplikacja która pobiera dane z pliku xlsx i na ich podstawie tworzy dokument docx z wykorzystaniem obsługi szblonów
# aplikacja napisana w pythonie i tkinter oraz biblioteką docxtpl
# powinna zawierać przycisk do wybierania pliku xlsx z danymi oraz przycisk do wygenerowania dokumentu docx dodatkowo powinna zawierać pola do wyboru daty wystawienia, data rozpoczęcia i daty zakonczenia
# oraz godziny rozpoczecia zajęc
# dane do aplikacji z listą zawodów, nazwiskami, imionami, datami urodzenia i numerami PESEL powinny być pobierane z pliku xlsx
# dane powinny być pobierane z pliku xlsx i zapisywane w słowniku

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import docxtpl
from datetime import datetime
from datetime import date
from datetime import timedelta
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

class App(tk.Tk):
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TButton', font=('Helvetica', 12))
    style.configure('TLabel', font=('Helvetica', 12))
    style.configure('TCombobox', font=('Helvetica', 12))
    style.configure('TEntry', font=('Helvetica', 12))
    style.configure('TFrame', font=('Helvetica', 12))
    style.configure('TRadiobutton', font=('Helvetica', 12))
    style.configure('TCheckbutton', font=('Helvetica', 12))
    style.configure('TNotebook', font=('Helvetica', 12))
    style.configure('TNotebook.Tab', font=('Helvetica', 12))
    style.configure('TScrollbar', font=('Helvetica', 12))
    style.configure('TMenubutton', font=('Helvetica', 12))
    style.configure('TCombobox', font=('Helvetica', 12))
    style.configure('TCheckbutton', font=('Helvetica', 12))
    style.configure('TLabel', font=('Helvetica', 12))
    style.configure('TNotebook.Tab', font=('Helvetica', 12))
    style.configure('TFrame', font=('Helvetica', 12))
    style.configure('TLabel', font=('Helvetica', 12))
    style.configure('TButton', font=('Helvetica', 12))

    def __init__(self):
        super().__init__()
        self.title("Aplikacja do generowania dokumentów")
        self.geometry("600x400")
        self.resizable(False, False)
        self.create_widgets()
        self.mainloop()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text='Generowanie dokumentów')
        self.notebook.add(self.tab2, text='Ustawienia')
        self.create_widgets_tab1()
        self.create_widgets_tab2()

    def create_widgets_tab1(self):

        self.label1 = ttk.Label(self.tab1, text="Wybierz plik z danymi")
        self.label1.grid(row=0, column=0, padx=10, pady=10)
    
        self.entry1 = ttk.Entry(self.tab1, width=50)
        self.entry1.grid(row=0, column=1, padx=10, pady=10)

        self.button1 = ttk.Button(self.tab1, text="Wybierz plik", command=self.open_file)
        self.button1.grid(row=0, column=2, padx=10, pady=10)

        self.label2 = ttk.Label(self.tab1, text="Wybierz szablon")
        self.label2.grid(row=1, column=0, padx=10, pady=10)

        self.entry2 = ttk.Entry(self.tab1, width=50)
        self.entry2.grid(row=1, column=1, padx=10, pady=10)

        self.button2 = ttk.Button(self.tab1, text="Wybierz plik", command=self.open_file)
        self.button2.grid(row=1, column=2, padx=10, pady=10)

        self.label3 = ttk.Label(self.tab1, text="Wybierz folder docelowy")
        self.label3.grid(row=2, column=0, padx=10, pady=10)

        self.entry3 = ttk.Entry(self.tab1, width=50)
        self.entry3.grid(row=2, column=1, padx=10, pady=10)

        self.button3 = ttk.Button(self.tab1, text="Wybierz folder", command=self.open_file)
        self.button3.grid(row=2, column=2, padx=10, pady=10)


    def open_file(self):
        self.file = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        self.entry1.insert(0, self.file) 

    def create_widgets_tab2(self):
        self.label4 = ttk.Label(self.tab2, text="Wybierz folder z szablonami")
        self.label4.grid(row=0, column=0, padx=10, pady=10)

        self.entry4 = ttk.Entry(self.tab2, width=50)
        self.entry4.grid(row=0, column=1, padx=10, pady=10)

        self.button4 = ttk.Button(self.tab2, text="Wybierz folder", command=self.open_file)
        self.button4.grid(row=0, column=2, padx=10, pady=10)

        self.label5 = ttk.Label(self.tab2, text="Wybierz folder z danymi")
        self.label5.grid(row=1, column=0, padx=10, pady=10)

        self.entry5 = ttk.Entry(self.tab2, width=50)
        self.entry5.grid(row=1, column=1, padx=10, pady=10)

        self.button5 = ttk.Button(self.tab2, text="Wybierz folder", command=self.open_file)
        self.button5.grid(row=1, column=2, padx=10, pady=10)

        self.label6 = ttk.Label(self.tab2, text="Wybierz folder docelowy")
        self.label6.grid(row=2, column=0, padx=10, pady=10)

        self.entry6 = ttk.Entry(self.tab2, width=50)
        self.entry6.grid(row=2, column=1, padx=10, pady=10)

        self.button6 = ttk.Button(self.tab2, text="Wybierz folder", command=self.open_file)
        self.button6.grid(row=2, column=2, padx=10, pady=10)

if __name__ == "__main__":
    app = App()
