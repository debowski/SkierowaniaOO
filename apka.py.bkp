import tkinter as tk
import tkinter.ttk as ttk

class App:

    def wstaw_kontrolki(self):
        # Funkcja wstawiająca kontrolki do okna.
        # Należy w niej utworzyć wszystkie kontrolki i wstawić je do okna.
        # Zwraca: nic
        self.label = tk.Label(self.root, text=self.napis1)
        self.label.grid(column=0, row=0, sticky="nsew")

    def __init__(self):

        self.napis1: str = "Ala ma kota"
        
        self.root = tk.Tk()
        self.root.title("Moja wspaniała aplikacja")
        self.root.geometry("500x400")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.wstaw_kontrolki()

    def wstaw_kontrolki(self): 
        self.ramka: tk.Frame = tk.Frame(self.root, bg="red")
        self.ramka.grid(row=0, column=0, sticky="nsew")
        self.ramka.columnconfigure(0, weight=1)
        self.label1: ttk.Label = ttk.Label(self.ramka, text=self.napis1)
        self.label1.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.btn1: ttk.Button = ttk.Button(self.ramka, text="Kliknij mnie", command=self.zmien_napis)
        self.btn1.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    def zmien_napis(self) -> None:
        try:
            self.napis2: str = "Kot ma Alę"
            self.label1.config(text=self.napis2)
        except Exception as e:
            print(e)


if __name__ == "__main__":

    app = App()
    app.root.mainloop()





