import tkinter as tk
from ui import EmailSenderApp


def main():
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()