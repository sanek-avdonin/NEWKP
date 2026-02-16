import tkinter as tk

from kp_generator.gui import AppGUI  # абсолютный импорт


def main():
    root = tk.Tk()
    AppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
