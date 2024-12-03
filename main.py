import tkinter as tk
from ui import PatientForm


def main():
    root = tk.Tk()
    root.option_add('*Font', 'Arial 14')
    PatientForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()
