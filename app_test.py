import tkinter as tk
import tkinter.ttk as ttk
import ttkbootstrap as ttkb
from tkinter import StringVar
from tkinter import filedialog
import pandas as pd


class App:
    def __init__(self):

        self.napis1 = "to jest test"
        self.root = ttkb.Window(themename="flatly")
        self.root.title("Skierowania 0.22")

        self.root.grid()
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.dodaj_widzety()

    def dodaj_widzety(self):
        self.frame = ttkb.Frame(self.root)
        self.frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=1)
        self.frame.columnconfigure(2, weight=1)

        self.btn_wyb_plik = ttkb.Button(
            self.frame, text="Wybierz plik", bootstyle="info", command=self.otwarcie_pliku)
        self.btn_wyb_plik.grid(row=0, column=0, sticky="nsew",
                               columnspan=3, padx=5, pady=5)

        # Wybieranie klasy

        self.var = StringVar()

        self.radio1 = ttkb.Radiobutton(
            self.frame, text="Klasa 1", variable=self.var, value="1", bootstyle="success-outline-toolbutton", command=self.set_lista_zawodow_1)
        self.radio2 = ttkb.Radiobutton(
            self.frame, text="Klasa 2", variable=self.var, value="2", bootstyle="warning-outline-toolbutton", command=self.set_lista_zawodow_2)
        self.radio3 = ttkb.Radiobutton(
            self.frame, text="Klasa 3", variable=self.var, value="3", bootstyle="danger-outline-toolbutton", command=self.set_lista_zawodow_3)
        self.radio1.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.radio2.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        self.radio3.grid(row=1, column=2, sticky="nsew", padx=5, pady=5)

        self.radiobuttons = []
        self.lista_zawodow = StringVar()

        self.zawody = ('Sprzedawca',
                       'cukiernik',
                       'Konserwator')

        self.frame2 = ttkb.Frame(self.root, bootstyle="success")
        self.frame2.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        # wstaw combobox z wyborem zawodow

        self.combo_current_var = tk.StringVar()
        self.combobox = ttkb.Combobox(
            self.frame, values=self.zawody, textvariable=self.combo_current_var, bootstyle="success")
        self.combobox.grid(row=2, column=0, sticky="nsew",
                           padx=15, pady=15, columnspan=3)

        self.data_wystawienia = ttkb.DateEntry(
            self.frame, firstweekday=0)
        self.data_wystawienia.grid(
            row=3, column=0, sticky="nsew", padx=5, pady=5)
        self.data_rozpoczecia = ttkb.DateEntry(
            self.frame)
        self.data_rozpoczecia.grid(
            row=4, column=0, sticky="nsew", padx=5, pady=5)
        self.data_zakonczenia = ttkb.DateEntry(
            self.frame)
        self.data_zakonczenia.grid(
            row=5, column=0, sticky="nsew", padx=5, pady=5)

    def set_lista_zawodow_1(self):
        self.combobox['values'] = list(self.zawody1)
        # ustaw styl bootstyle na success
        # self.combobox.configure(bootstyle="success")

    def set_lista_zawodow_2(self):
        self.combobox['values'] = list(self.zawody2)

    def set_lista_zawodow_3(self):
        self.combobox['values'] = list(self.zawody3)

    def utworz_lz(self):
        df = pd.read_excel(self.plik)
        # print(df.head)

        # Ekstrakcja liczby z kolumny 'Dane oddziału'
        df['Oddział'] = df['Dane oddziału'].str.extract(
            '(\d+)', expand=False).astype(int)

        # Utworzenie trzech zbiorów na podstawie wartości kolumny 'Oddział'
        self.zawody1 = set(df[df['Oddział'] == 1]
                           ['Specjalność/Zawód'].tolist())
        self.zawody2 = set(df[df['Oddział'] == 2]
                           ['Specjalność/Zawód'].tolist())
        self.zawody3 = set(df[df['Oddział'] == 3]
                           ['Specjalność/Zawód'].tolist())

        print(self.zawody1)

    def otwarcie_pliku(self):
        filetypes = (
            ('Arkusze', '*.xlsx'),
            ('All files', '*.*')
        )

        self.plik = filedialog.askopenfilename(title='Wybierz plik',
                                               initialdir='..\\Data',
                                               filetypes=filetypes)

        self.btn_wyb_plik.configure(text=self.plik)
        self.utworz_lz()
        self.radio1.invoke()


if __name__ == "__main__":
    app = App()
    app.root.mainloop()
