from tkinter import *

def main_window():
    global window

    window = Tk()
    window.title("Espm import 1.0")
    window.geometry("600x300")

    menubar = Menu(window)
    menubar.add_command(label="Login")
    
    Label(text="Espm import 1.0", width="600",
        height="2", font=("calibri", 15)).pack()
    
    window.config(menu=menubar)
    window.mainloop()

main_window()