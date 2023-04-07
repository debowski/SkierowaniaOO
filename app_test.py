import tkinter as tk
import tkinter.ttk as ttk

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Widzety")
        self.root.geometry("300x300")
        self.root.grid()
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.dodaj_widzety()




    def dodaj_widzety(self):
        self.frame = tk.Frame(self.root, background="red")
        self.frame.grid(row=0, column=0, sticky="nsew")

        self.frame.columnconfigure(0, weight=1)

        self.button = ttk.Button(self.frame, text="Dodaj widzety", command=self.ble)
        self.button.grid(row=0, column=0, sticky="nsew")
    
    def ble(self):
        print("test")


if __name__ == "__main__":
    app = App()
    app.root.mainloop()


