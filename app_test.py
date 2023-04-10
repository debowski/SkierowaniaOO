import tkinter as tk
import tkinter.ttk as ttk
import ttkbootstrap as ttkb
from tkinter import StringVar
from tkinter import filedialog
import pandas as pd


class App:
    def __init__(self):

        self.napis1 = "to jest test"

        # self.root = tk.Tk()
        self.root = ttkb.Window(themename="flatly")
        self.root.title("Skierowania 0.22")
        # self.root.geometry("300x300")
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
        # self.radio1.invoke()

        self.radiobuttons = []
        self.lista_zawodow = StringVar()

        self.zawody1 = ('Sprzedawca',
                        'cukiernik',
                        'Konserwator')

        self.zawody2 = ('Sprzedawca',
                        'cukiernik',
                        'Konserwator')

        self.zawody3 = ('Sprzedawca',
                        'cukiernik',
                        'Konserwator')

        self.frame2 = ttkb.Frame(self.root, bootstyle="success")
        self.frame2.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        self.label1 = ttkb.Label(
            self.frame2, text="lajkshfkjahskjfhaksjhfkjahs fkjahs kfjha skjf")
        self.label1.grid(row=0, column=0)

    def ble(self):
        self.button.configure(text=self.napis1)

    def set_lista_zawodow_1(self):
        for a in self.radiobuttons:
            a.grid_forget()
        for zawod in self.zawody1:
            self.r = ttkb.Radiobutton(
                self.frame, text=zawod, value=zawod, variable=self.lista_zawodow, bootstyle="success-outline-toolbutton")
            self.r.grid(padx=5, pady=5, sticky="nsew", columnspan=3)
            self.radiobuttons.append(self.r)

    def set_lista_zawodow_2(self):
        for b in self.radiobuttons:
            b.grid_forget()

        for zawod in self.zawody2:
            self.r = ttkb.Radiobutton(
                self.frame, text=zawod, value=zawod, variable=self.lista_zawodow, bootstyle="warning-outline-toolbutton")
            self.r.grid(padx=5, pady=5, sticky="nsew", columnspan=3)
            self.radiobuttons.append(self.r)

    def set_lista_zawodow_3(self):
        for c in self.radiobuttons:
            c.grid_forget()
        for zawod in self.zawody3:
            self.r = ttkb.Radiobutton(
                self.frame, text=zawod, value=zawod, variable=self.lista_zawodow, bootstyle="danger-outline-toolbutton")
            self.r.grid(padx=5, pady=5, sticky="nsew", columnspan=3)
            self.radiobuttons.append(self.r)

    def utworz_lz(self):
        df = pd.read_excel(self.plik)
        print(df.head)

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
        print(self.zawody2)
        print(self.zawody3)

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


if __name__ == "__main__":
    app = App()
    app.root.mainloop()
