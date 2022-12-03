import tkinter as tk
from data import MainFrame


def main():
    root = tk.Tk()
    root.title('Raport z wypłat')
    root.geometry('650x350')
    root.resizable(True, True)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    main_frame = MainFrame(root)
    root.mainloop()


if __name__ == '__main__':
    main()
