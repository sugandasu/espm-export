from tkinter import *
import requests

def main_window():
    global window

    window = Tk()
    window.title("Espm import 1.0")
    window.geometry("600x300")

    menubar = Menu(window)
    menubar.add_command(label="Login", command=login_window)

    Label(text="Espm import 1.0", width="600",
          height="2", font=("calibri", 15)).pack()

    window.config(menu=menubar)
    window.mainloop()


def login_window():
    global wlogin
    global username_strvar
    global password_strvar
    global username_entry
    global password_entry

    wlogin = Toplevel(window)
    wlogin.title("Login")
    wlogin.geometry("300x250")

    username_strvar = StringVar()
    password_strvar = StringVar()

    Label(wlogin, text="Login", width="600",
          height="2", font=("calibri", 15)).pack()
    Label(wlogin, text="Username *").pack()
    username_entry = Entry(wlogin, textvariable=username_strvar)
    username_entry.pack()
    Label(wlogin, text="Password *").pack()
    password_entry = Entry(wlogin, textvariable=password_strvar)
    password_entry.pack()
    Label(wlogin, text="").pack()

    Button(wlogin, text="Login", command=login_request).pack()


def login_request():
    global access_token
    username = username_strvar.get()
    password = password_strvar.get()

    req = requests
    payload = {"username": username, "password": password}

    req = requests.post("http://espm.test/api/auth/login", json=payload, headers={
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json"
    })

    fg = 'red'
    text = 'login gagal'

    if req.status_code == requests.codes.ok:
        username_entry.delete(0, END)
        password_entry.delete(0, END)

        req_json = req.json()
        user = req_json["user"]
        access_token = req_json["access_token"]
        
        # Label(wlogin, text=user).pack()
        # Label(wlogin, text=access_token).pack()
        
        fg = 'green'
        text = 'login berhasil'

    Label(wlogin, text=text, fg=fg).pack()

main_window()
