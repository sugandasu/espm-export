from tkinter import *
from tkinter import messagebox
import requests
import pyodbc


def reset_form():
    username_lbl.destroy()
    username_ntr.destroy()
    password_lbl.destroy()
    password_ntr.destroy()
    login_btn.destroy()


def import_form():
    global user_lbl
    global spm_number_lbl
    global last_import_lbl

    user_lbl = Label(window, text="Pengguna: " + user["username"])
    user_lbl.grid(row=1, column=0, sticky=W, padx=4)

    try:
        req = requests.get("http://espm.test/api/import-app/meta", headers={
            "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest",
            "Authorization": "Bearer " + access_token
        })

        req_json = req.json()
        spm_count = req_json["spm_count"]
        last_import = req_json["last_import"]

        text = "Jumlah Spm: " + str(spm_count)
        spm_number_lbl = Label(window, text=text)
        spm_number_lbl.grid(row=2, column=0, sticky=W, padx=4)
        text = "Tanggal Terakhir Import: " + str(last_import)
        last_import_lbl = Label(window, text=text)
        last_import_lbl.grid(row=3, column=0, sticky=W, padx=4)

        import_btn = Button(window, text="import", command=import_request)
        import_btn.grid(row=4, column=1, sticky=E, pady=4)
    except:
        pass


def import_request():
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=.\SQLEXPRESS;'
                          'Database=Simda_2019;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()
    cursor.execute('SELECT * FROM [Simda_2019].[dbo].[Ta_SPM]')

    url = "http://espm.test/api/import-app/spm"
    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    spms = []

    for row in cursor:
        spm = {
            'tahun': row[0],
            'no_spm': row[1],
            'kd_urusan': row[2],
            'kd_bidang': row[3],
            'kd_unit': row[4],
            'kd_sub': row[5],
            'no_spp': row[6],
            'jn_spm': row[7],
            'tgl_spm': str(row[8]),
            'uraian': row[9],
            'nm_penerima': row[10],
            'bank_penerima': row[11],
            'rek_penerima': row[12],
            'npwp': row[13],
            'bank_pembayar': row[14],
            'nm_penandatangan': row[16],
            'nip_penandatangan': row[17],
            'jbt_penandatangan': row[18],
            'kd_edit': row[19]
        }

        spms.append(spm)
    
    payloads = {
        "spms": spms
    }

    try:
        req = requests.post(url, json=payloads, headers=headers)
    except:
        messagebox.showerror("Import", "import gagal")

    if req.status_code == requests.codes.ok:
        messagebox.showinfo("Import", "import berhasil")
    
    print(req.text)

def login_form():
    global username_ntr
    global username_lbl
    global password_ntr
    global password_lbl
    global login_btn

    username_lbl = Label(window, text="Username")
    username_lbl.grid(row=1, column=0, sticky=W, padx=4)
    username_ntr = Entry(window, width=20)
    username_ntr.grid(row=1, column=1, sticky=E, pady=4)

    password_lbl = Label(window, text="Password")
    password_lbl.grid(row=2, column=0, sticky=W, padx=4)
    password_ntr = Entry(window, width=20)
    password_ntr.grid(row=2, column=1, sticky=E, pady=4)

    login_btn = Button(window, text="submit", command=login_request)
    login_btn.grid(row=3, column=1, sticky=E, pady=4)


def login_request():
    global user
    global access_token

    username = username_ntr.get()
    password = password_ntr.get()

    req = requests
    payload = {"username": username, "password": password}

    try:
        req = requests.post("http://espm.test/api/auth/login", json=payload, headers={
            "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest",
            "Accept": "application/json"
        })
    except:
        messagebox.showerror("Login", "login gagal")

    if req.status_code == requests.codes.ok:
        req_json = req.json()
        user = req_json["user"]
        access_token = req_json["access_token"]

        reset_form()
        import_form()
    else:
        messagebox.showerror("Login", "login gagal")


def main_window():
    global window

    window = Tk()
    window.title("Espm import 1.0")
    window.geometry("600x300")

    Label(window, text="Espm import 1.0", height="1", font=(
        "calibri", 20)).grid(row=0, column=0, sticky=W, pady=10)

    login_form()
    # import_request()

    window.mainloop()


main_window()
