import tkinter as tk
from tkinter import ttk
from tkinter import StringVar


class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Przykład list radiobuttonów")
        self.geometry("300x200")
        # zmienna globalna przechowująca aktualnie wyświetlaną listę radiobuttonów
        self.current_rb_list = None

        self.button1 = ttk.Button(
            self, text="Lista 1", command=self.show_rb_list_1)
        self.button1.pack(pady=10)

        self.button2 = ttk.Button(
            self, text="Lista 2", command=self.show_rb_list_2)
        self.button2.pack(pady=10)

    def show_rb_list_1(self):
        if self.current_rb_list is not None:  # usuwanie aktualnej listy radiobuttonów
            self.current_rb_list.destroy()
        self.current_rb_list = tk.Frame(self)
        self.current_rb_list.pack(pady=10)

        lista_zawodow = tk.StringVar()
        zawody = (("Sprzedawca", "Sprzedawca"),
                  ("Cukiernik", "Cukiernik"),
                  ("Piekarz", "Piekarz"),
                  ("Fryzjer", "Fryzjer"))

        for zawod in zawody:
            r = ttk.Radiobutton(
                self.current_rb_list, text=zawod[0], value=zawod[1], variable=lista_zawodow)
            r.pack(pady=5, anchor="w")

    def show_rb_list_2(self):
        if self.current_rb_list is not None:  # usuwanie aktualnej listy radiobuttonów
            self.current_rb_list.destroy()
        self.current_rb_list = tk.Frame(self)
        self.current_rb_list.pack(pady=10)

        lista_wielkosci = tk.StringVar()
        wielkosci = (("Mały", "Mały"),
                     ("Średni", "Średni"),
                     ("Duży", "Duży"))

        for rozmiar in wielkosci:
            r = ttk.Radiobutton(
                self.current_rb_list, text=rozmiar[0], value=rozmiar[1], variable=lista_wielkosci)
            r.pack(pady=5, anchor="w")


if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
