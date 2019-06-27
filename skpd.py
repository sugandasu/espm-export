from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import requests
import pyodbc
import aiohttp
import asyncio
import csv
import os
from aiohttp import FormData
import xlsxwriter

main_loop = asyncio.get_event_loop()
index = 0


def validated_row(row):
    for i in range(0, len(row)):
        if row[i] == None:
            row[i] = ''

    return row


def get_data(query, querycount):
    conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=.\SQLEXPRESS;"
                          "Database=Simda_2019;"
                          "Trusted_Connection=yes;")

    cursor = conn.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.execute(querycount)
    (rowscount,) = cursor.fetchone()

    return rows, rowscount


def reset_file(filename):
    if os.path.exists(filename):
        os.remove(filename)
        return False


async def send_request(url, filename):
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    data = FormData()
    data.add_field('file',
                   open(filename, 'rb'),
                   filename=filename,
                   content_type='application/vnd.ms-excel')

    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data, headers=headers) as response:
            data = {
                "message": "failed"
            }
            if response.status == 200:
                data = await response.json()

            print(await response.json())


async def send_request_file(url, filename, rows, rowscount, fieldnames):
    reset_file(filename)

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    col = 0
    for fieldname in fieldnames:
        worksheet.write(0, col, fieldname)
        col += 1

    row = 1
    rowcount = 1
    for columns in rows:
        col = 0
        columns = validated_row(columns)
        for column in columns:
            worksheet.write(row, col, column)
            col += 1

        # kalau data sudah 1000 dan sudah sampai akhir row
        if (row % 1000 == 0) or (rowcount == rowscount):
            workbook.close()
            await send_request(url, filename)

            reset_file(filename)
            workbook = xlsxwriter.Workbook(filename)
            worksheet = workbook.add_worksheet()

            col = 0
            for fieldname in fieldnames:
                worksheet.write(0, col, fieldname)
                col += 1

            row = 0

        row += 1
        rowcount += 1


async def send_spm_request_file(url, rows, rowscount, fieldnames):
    filename = "spmfile.xlsx"
    reset_file(filename)

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    col = 0
    for fieldname in fieldnames:
        worksheet.write(0, col, fieldname)
        col += 1

    row = 1
    rowcount = 1
    for columns in rows:
        col = 0
        columns = validated_row(columns)
        for column in columns:
            worksheet.write(row, col, column)
            col += 1

        await import_spm_detail(columns[1])
        await import_spm_tax(columns[1])

        # kalau data sudah 1000 dan sudah sampai akhir row
        if (row % 1000 == 0) or (rowcount == rowscount):
            Label(window, text=str(rowcount) +
                  " data").grid(row=6, column=2, sticky=E)
            workbook.close()

            await send_request(url, filename)

            reset_file(filename)
            workbook = xlsxwriter.Workbook(filename)
            worksheet = workbook.add_worksheet()

            col = 0
            for fieldname in fieldnames:
                worksheet.write(0, col, fieldname)
                col += 1

            row = 0

        row += 1
        rowcount += 1

    messagebox.showinfo("Import", "Import selesai")


async def import_spm():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(skpd["kd_sub"])
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(skpd["kd_sub"])
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm"
    fieldnames = [
        "tahun", "no_spm", "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "no_spp", "jn_spm", "tgl_spm", "uraian",
        "nm_penerima", "bank_penerima", "rek_penerima", "npwp",
        "bank_pembayar", "nm_penandatangan", "nip_penandatangan",
        "jbt_penandatangan", "kd_edit"
    ]

    data = await send_spm_request_file(url, rows, rowscount, fieldnames)
    return data


async def import_spm_detail(no_spm):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Rinc] WHERE no_spm='" + \
        str(no_spm) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Rinc] WHERE no_spm='" + str(
        no_spm) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm-detail"
    filename = "spmdetail.xlsx"
    fieldnames = [
        "tahun", "no_spm", "no_id", "kd_urusan", "kd_bidang",
        "kd_unit", "kd_sub", "kd_prog", "id_prog", "kd_keg", "kd_rek_1",
        "kd_rek_2", "kd_rek_3", "kd_rek_4", "kd_rek_5", "nilai"
    ]

    data = await send_request_file(url, filename, rows, rowscount, fieldnames)

    for row in rows:
        spmdetail = [row[0], row[3], row[4], row[5], row[6], row[7], row[8], row[9]]
        await import_spm_detail_event(spmdetail)

    return data


async def import_spm_detail_event(spmdetail):
    query = ("SELECT * FROM [Simda_2019].[dbo].[Ta_Kegiatan]" +
        " WHERE tahun=" + str(spmdetail[0]) +
        " AND kd_urusan=" + str(spmdetail[1]) +
        " AND kd_bidang=" + str(spmdetail[2]) +
        " AND kd_unit=" + str(spmdetail[3]) +
        " AND kd_sub=" + str(spmdetail[4]) +
        " AND kd_prog=" + str(spmdetail[5]) +
        " AND id_prog=" + str(spmdetail[6]) +
        " AND kd_keg=" + str(spmdetail[7]))
    querycount = ("SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_Kegiatan]" +
        " WHERE tahun=" + str(spmdetail[0]) +
        " AND kd_urusan=" + str(spmdetail[1]) +
        " AND kd_bidang=" + str(spmdetail[2]) +
        " AND kd_unit=" + str(spmdetail[3]) +
        " AND kd_sub=" + str(spmdetail[4]) +
        " AND kd_prog=" + str(spmdetail[5]) +
        " AND id_prog=" + str(spmdetail[6]) +
        " AND kd_keg=" + str(spmdetail[7]))
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/event"
    filename = "spmevent.xlsx"
    fieldnames = [
        "tahun", "kd_urusan", "kd_bidang", "kd_unit", "kd_sub",
        "kd_prog", "id_prog", "kd_keg", "ket_kegiatan"
    ]

    await send_request_file(url, filename, rows, rowscount, fieldnames)


async def import_spm_tax(no_spm):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Pot] WHERE no_spm='" + \
        str(no_spm) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Pot] WHERE no_spm='" + str(
        no_spm) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm-tax-account"
    filename = "spmtax.xlsx"
    fieldnames = [
        "tahun", "no_spm", "kd_pot_rek", "nilai"
    ]

    return await send_request_file(url, filename, rows, rowscount, fieldnames)


def import_spm_request():
    future = asyncio.create_task(import_spm())
    try:
        main_loop.run_until_complete(future)
    except:
        pass


def reset_form():
    username_lbl.destroy()
    username_ntr.destroy()
    password_lbl.destroy()
    password_ntr.destroy()
    login_btn.destroy()


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
    global skpd
    global access_token

    username = username_ntr.get()
    password = password_ntr.get()

    payload = {"username": username, "password": password}
    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json"
    }

    try:
        req = requests.post("http://espm.test/api/auth/login",
                            json=payload, headers=headers)
    except:
        messagebox.showerror("Login", "login gagal")

    if req.status_code == 200:
        req_json = req.json()

        role = req_json["user"]["roles"][0]["abbr"]

        if role != "skpd":
            messagebox.showerror("Login", "login gagal")
        else:
            user = req_json["user"]
            skpd = req_json["user"]["skpds"][0]
            access_token = req_json["access_token"]

            reset_form()
            import_form()
    else:
        messagebox.showerror("Login", "login gagal")


def import_form():
    global user_lbl
    global import_btn
    global spm_ntr

    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    user_lbl = Label(window, text="Pengguna: " + user["username"])
    user_lbl.grid(row=1, column=0, sticky=W, padx=4)

    try:
        req = requests.get(
            "http://espm.test/api/import-app/meta", headers=headers)

        req_json = req.json()

        skpd_name = req_json["skpd_name"]
        spm_count = req_json["spm_count"]
        last_import = req_json["last_import"]

        skpd_name_text = "SKPD: " + str(skpd_name)
        Label(window, text=skpd_name_text).grid(
            row=2, column=0, sticky=W, padx=4)

        spm_count_text = "Jumlah Spm: " + str(spm_count)
        Label(window, text=spm_count_text).grid(
            row=3, column=0, sticky=W, padx=4)

        last_import_text = "Import Terakhir: " + str(last_import)
        Label(window, text=last_import_text).grid(
            row=4, column=0, sticky=W, padx=4)

        Label(window, text="Import data spm", font=("calibri", 12)).grid(
            row=5, column=0, sticky=W, padx=4)
        Label(window, text="Import").grid(row=6, column=0, sticky=W, padx=4)
        import_btn = Button(window, text="import", command=import_spm_request)
        import_btn.grid(row=6, column=1, sticky=E, pady=4)
    except:
        messagebox.showerror("Login", "login gagal")


async def run_tk(root):
    try:
        while True:
            root.update()
            await asyncio.sleep(.01)
    except Tk.TclError as e:
        if "application has been destroyed" not in e.args[0]:
            raise

if __name__ == "__main__":
    global window
    window = Tk()

    window.title("Espm import 1.0")
    window.geometry("700x400")

    Label(window, text="Espm import SKPD 1.0", height="1", font=(
        "calibri", 20)).grid(row=0, column=0, sticky=W, pady=10)

    login_form()

    main_loop.run_until_complete(run_tk(window))
