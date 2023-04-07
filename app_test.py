import tkinter as tk
import tkinter.ttk as ttk
import ttkbootstrap as ttkb

class App:
    def __init__(self):

        self.napis1 = "to jest test"


        self.root = tk.Tk()
        self.root.title("Skierowania 0.22")
        self.root.geometry("300x300")
        self.root.grid()
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.dodaj_widzety()




    def dodaj_widzety(self):
        self.frame = ttkb.Frame(self.root, bootstyle="info")
        self.frame.grid(row=0, column=0, sticky="nsew")

        self.frame.columnconfigure(0, weight=1)

        self.button = ttkb.Button(self.frame, text="Wybierz plik", command=self.ble)
        self.button.grid(row=0, column=0, sticky="nsew")
    
    def ble(self):
        print(self.napis1)


if __name__ == "__main__":
    app = App()
    app.root.mainloop()


