from dotenv import load_dotenv
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from aiohttp import FormData

import requests
import pyodbc
import aiohttp
import asyncio
import csv
import os
import xlsxwriter

main_loop = asyncio.get_event_loop()
load_dotenv()

def read_server_file():
    global server_file_frame
    global login_frame

    server_file_exist = os.path.exists('sqlserver.txt')
    if server_file_exist:
        file = open("sqlserver.txt", "r")
        line = file.readline()
        server_db = line.split()
        server = server_db[0]
        db = server_db[1]
        test_connection(server, db)
        if test_connection(server, db):
            login_frame = login_view()
            login_frame.pack(fill="both", expand=True)
        else:
            server_file_frame = server_file_view()
            server_file_frame.pack(fill="both", expand=True)
    else:
        server_file_frame = server_file_view()
        server_file_frame.pack(fill="both", expand=True)


def test_connection_input():
    server_in = server_ent.get()
    db_in = db_ent.get()

    if test_connection(server_in, db_in):
        messagebox.showinfo("Koneksi", "Koneksi berhasil")
        save_server_file()
    else:
        messagebox.showerror("Koneksi", "Koneksi gagal")


def test_connection(server, db):
    global conn
    server=os.getenv(server)
    database=os.getenv(database) 
    username=os.getenv(username)
    password=os.getenv(password)
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

    if conn:
        return True
    else:
        return False


def save_server_file():
    server_in = server_ent.get()
    db_in = db_ent.get()

    if os.path.exists("sqlserver.txt"):
        os.remove("sqlserver.txt")

    file = open("sqlserver.txt", "w")
    file.write(server_in)
    file.write(" ")
    file.write(db_in)
    file.close()

    server_file_frame.destroy()
    login_frame = login_view()
    login_frame.pack(fill="both", expand=True)


def reset_file(filename):
    if os.path.exists(filename):
        os.remove(filename)
        return False

# DATABASE


def get_data(query, querycount):
    cursor = conn.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.execute(querycount)
    (rowscount,) = cursor.fetchone()

    return rows, rowscount


def validated_row(row):
    for i in range(0, len(row)):
        if row[i] == None:
            row[i] = ''

    return row

# IMPORT


def validate_date(date):
    try:
        int(date)

        date_len = len(date)
        if date_len != 8:
            return False

        date_month = date[4:6]
        if int(date_month) > 12:
            return False

        date_day = date[6:8]
        if int(date_day) > 31:
            return False

        return True
    except:
        pass

    return False


def validate_import():
    import_index = import_cbx.current()
    import_data = import_cbx.get()
    begin_data = begin_ent.get()
    end_data = end_ent.get()
    spm_data = spm_ent.get()

    begin_data = begin_data.replace('-', '')
    end_data = end_data.replace('-', '')

    if (import_index >= 0 and import_index <= 2):
        if import_data == "Spm":
            if spm_data == "":
                if begin_data == "" or end_data == "":
                    messagebox.showerror(
                        "Tanggal", "Masukan tanggal mulai dan akhir")
                else:
                    if validate_date(begin_data) and validate_date(end_data):
                        future = asyncio.create_task(
                            import_spm_all(begin_data, end_data))
                    else:
                        messagebox.showerror(
                            "Tanggal", "Format tanggal mulai atau akhir salah")
            else:
                future = asyncio.create_task(import_spm_one(spm_data))
        elif import_data == "Anggaran":
            if begin_data == "" or end_data == "":
                messagebox.showerror(
                    "Tanggal", "Masukan tanggal mulai dan akhir")
            else:
                if validate_date(begin_data) and validate_date(end_data):
                    future = asyncio.create_task(
                        import_anggaran_all(begin_data, end_data))
                else:
                    messagebox.showerror(
                        "Tanggal", "Format tanggal mulai atau akhir salah")
        elif import_data == "Anggaran Triwulan":
            if begin_data == "" or end_data == "":
                messagebox.showerror(
                    "Tanggal", "Masukan tanggal mulai dan akhir")
            else:
                if validate_date(begin_data) and validate_date(end_data):
                    future = asyncio.create_task(
                        import_anggaran_tw_all(begin_data, end_data))
                else:
                    messagebox.showerror(
                        "Tanggal", "Format tanggal mulai atau akhir salah")

        try:
            main_loop.run_until_complete(future)
        except:
            pass
    else:
        messagebox.showerror("Import", "Pilihan import tidak valid")


async def import_anggaran_all(begin_data, end_data):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_RASK_Arsip] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND DateCreate >= '" + str(begin_data) + " 00:00:00.000'" + " AND DateCreate <= '" + str(
        end_data) + " 00:00:00.000'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_RASK_Arsip] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND DateCreate >= '" + str(begin_data) + " 00:00:00.000'" + " AND DateCreate <= '" + str(
        end_data) + " 00:00:00.000'"
    rows, rowscount = get_data(query, querycount)
    url = "https://espm.online/api/import-app/budget"
    filename = "anggaran.xlsx"
    fieldnames = [
        "tahun", "kd_perubahan", "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "kd_prog", "id_prog", "kd_keg", "kd_rek_1", "kd_rek_2",
        "kd_rek_3", "kd_rek_4", "kd_rek_5", "no_rinc", "no_id", "keterangan_rinc", "sat_1",
        "nilai_1", "sat_2", "nilai_2", "sat_3", "nilai_3", "satuan123",
        "jml_satuan", "nilai_rp", "total", "keterangan", "kd_ap_pub", "kd_sumber", "datecreate"
    ]

    if rowscount > 0:
        messagebox.showinfo("Import", "Import " + str(rowscount) + " data")
        data = await send_request_file(url, filename, rows, rowscount, fieldnames)
        messagebox.showinfo("Import", "Import selesai")
        return data
    else:
        messagebox.showinfo("Import", "Data kosong")


async def import_anggaran_tw_all(begin_data, end_data):
    begin_data = begin_data[0:4]
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_Rencana_Arsip] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND tahun = '" + str(begin_data) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_Rencana_Arsip] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND tahun = '" + str(begin_data) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "https://espm.online/api/import-app/budget-tw"
    filename = "anggaran-tw.xlsx"
    fieldnames = [
        "tahun", "kd_perubahan", "kd_urusan", "kd_bidang",
        "kd_unit", "kd_sub", "kd_prog", "id_prog",
        "kd_keg", "kd_rek_1", "kd_rek_2", "kd_rek_3",
        "kd_rek_4", "kd_rek_5", "jan", "feb",
        "mar", "apr", "mei", "jun",
        "jul", "agt", "sep", "okt",
        "nop", "des"
    ]

    if rowscount > 0:
        messagebox.showinfo("Import", "Import " + str(rowscount) + " data")
        data = await send_request_file(url, filename, rows, rowscount, fieldnames)
        messagebox.showinfo("Import", "Import selesai")
        return data
    else:
        messagebox.showinfo("Import", "Data kosong")


async def import_spm_all(begin_data, end_data):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND Tgl_SPM >= '" + str(begin_data) + " 00:00:00.000'" + " AND Tgl_SPM <= '" + str(
        end_data) + " 00:00:00.000'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(
        skpd["kd_sub"]) + " AND Tgl_SPM >= '" + str(begin_data) + " 00:00:00.000'" + " AND Tgl_SPM <= '" + str(
        end_data) + " 00:00:00.000'"

    rows, rowscount = get_data(query, querycount)

    if rowscount > 0:
        url = "https://espm.online/api/import-app/spm"
        fieldnames = [
            "tahun", "no_spm", "kd_urusan", "kd_bidang", "kd_unit",
            "kd_sub", "no_spp", "jn_spm", "tgl_spm", "uraian",
            "nm_penerima", "bank_penerima", "rek_penerima", "npwp",
            "bank_pembayar", "nm_penandatangan", "nip_penandatangan",
            "jbt_penandatangan", "kd_edit"
        ]

        data = await send_spm_all(url, rows, rowscount, fieldnames)
        return data
    else:
        messagebox.showinfo("Import", "Data kosong")


async def import_spm_one(spm_data):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(skpd["kd_sub"]) + " AND no_spm='" + str(spm_data) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM] " + "WHERE kd_urusan=" + str(skpd["kd_urusan"]) + " AND kd_bidang=" + str(
        skpd["kd_bidang"]) + " AND kd_unit=" + str(skpd["kd_unit"]) + " AND kd_sub=" + str(skpd["kd_sub"]) + " AND no_spm='" + str(spm_data) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "https://espm.online/api/import-app/spm"
    fieldnames = [
        "tahun", "no_spm", "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "no_spp", "jn_spm", "tgl_spm", "uraian",
        "nm_penerima", "bank_penerima", "rek_penerima", "npwp",
        "bank_pembayar", "nm_penandatangan", "nip_penandatangan",
        "jbt_penandatangan", "kd_edit"
    ]

    return await send_spm_all(url, rows, rowscount, fieldnames)


async def import_spm_detail(no_spm):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Rinc] WHERE no_spm='" + \
        str(no_spm) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Rinc] WHERE no_spm='" + str(
        no_spm) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "https://espm.online/api/import-app/spm-detail"
    filename = "spmdetail.xlsx"
    fieldnames = [
        "tahun", "no_spm", "no_id", "kd_urusan", "kd_bidang",
        "kd_unit", "kd_sub", "kd_prog", "id_prog", "kd_keg", "kd_rek_1",
        "kd_rek_2", "kd_rek_3", "kd_rek_4", "kd_rek_5", "nilai"
    ]

    data = await send_request_file(url, filename, rows, rowscount, fieldnames)

    event_rows = []
    event_rowscount = 0
    for row in rows:
        spmdetail = [row[0], row[3], row[4],
                     row[5], row[6], row[7], row[8], row[9]]

        new_event_rows, new_event_rowscount = import_spm_detail_event(
            spmdetail)
        event_rows += new_event_rows
        event_rowscount += new_event_rowscount

    event_url = "https://espm.online/api/import-app/event"
    event_filename = "spmevent.xlsx"
    event_fieldnames = [
        "tahun", "kd_urusan", "kd_bidang", "kd_unit", "kd_sub",
        "kd_prog", "id_prog", "kd_keg", "ket_kegiatan"
    ]

    await send_request_file(event_url, event_filename, event_rows, event_rowscount, event_fieldnames)

    return data


def import_spm_detail_event(spmdetail):
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
    return rows, rowscount


async def import_spm_tax(no_spm):
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Pot] WHERE no_spm='" + \
        str(no_spm) + "'"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Pot] WHERE no_spm='" + str(
        no_spm) + "'"
    rows, rowscount = get_data(query, querycount)
    url = "https://espm.online/api/import-app/spm-tax-account"
    filename = "spmtax.xlsx"
    fieldnames = [
        "tahun", "no_spm", "kd_pot_rek", "nilai"
    ]

    return await send_request_file(url, filename, rows, rowscount, fieldnames)


async def send_spm_all(url, rows, rowscount, fieldnames):
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
            workbook.close()

            await async_send_request(url, filename)

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

# REQUEST


def login_request():
    global user
    global skpd
    global access_token
    global import_frame

    username_in = username_ent.get()
    password_in = password_ent.get()

    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json"
    }
    url = "https://espm.online/api/auth/login"
    payload = {"username": username_in, "password": password_in}
    try:
        req = send_request(headers, url, payload)
        if req.status_code == 200:
            req_json = req.json()
            role = req_json["user"]["roles"][0]["abbr"]

            if role == "skpd":
                user = req_json["user"]
                skpd = req_json["user"]["skpds"][0]
                access_token = req_json["access_token"]

                login_frame.destroy()
                import_frame = import_view()
                import_frame.pack(fill="both", expand=True)
            else:
                messagebox.showerror("Login", "login gagal")
        else:
            messagebox.showerror("Login", "login gagal")
    except:
        messagebox.showerror("Login", "login gagal")


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
            await async_send_request(url, filename)

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


async def async_send_request(url, filename):
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


def send_request(headers, url, payload):
    return requests.post(url, json=payload, headers=headers)

# VIEW


def server_file_view():
    global server_ent
    global db_ent

    frame = Frame(window)

    Label(frame, text="Espm import SKPD 1.0", height="2", font=(
        "calibri", 20)).grid(column=0, row=0, sticky=W, padx=5)

    Label(frame, text="Koneksi Server", height="2", font=(
        "calibri", 15)).grid(column=0, row=1, sticky=W, padx=5)

    server_lbl = Label(frame, text="Server")
    server_lbl.grid(column=0, row=2, sticky=W, padx=5)
    server_ent = Entry(frame)
    server_ent.grid(column=0, row=3, sticky=W, padx=5)
    db_lbl = Label(frame, text="Database")
    db_lbl.grid(column=0, row=4, sticky=W, padx=5)
    db_ent = Entry(frame)
    db_ent.grid(column=0, row=5, sticky=W, padx=5)
    con_btn = Button(frame, text="Tes Koneksi", command=test_connection_input)

    Label(frame, text="", height="1").grid(column=0, row=6, sticky=W, padx=5)

    con_btn.grid(column=0, row=7, sticky=E, padx=5)
    save_btn = Button(frame, text="Simpan", command=save_server_file)
    save_btn.grid(column=1, row=7, sticky=W, padx=5)

    return frame


def login_view():
    global username_ent
    global password_ent

    frame = Frame(window)

    Label(frame, text="Espm import SKPD 1.0", height="2", font=(
        "calibri", 20)).grid(column=0, row=0, sticky=W, padx=5)

    Label(frame, text="Login", height="2", font=(
        "calibri", 15)).grid(column=0, row=1, sticky=W, padx=5)

    username_lbl = Label(frame, text="Username")
    username_lbl.grid(row=2, column=0, sticky=W, padx=5)
    username_ent = Entry(frame, width=20)
    username_ent.grid(row=3, column=0, sticky=W, padx=5)
    password_lbl = Label(frame, text="Password")
    password_lbl.grid(row=4, column=0, sticky=W, padx=5)
    password_ent = Entry(frame, width=20)
    password_ent.grid(row=5, column=0, sticky=W, padx=5)

    Label(frame, text="", height="1").grid(column=0, row=6, sticky=W, padx=5)

    login_btn = Button(frame, text="submit", command=login_request)
    login_btn.grid(row=7, column=0, sticky=E, pady=5)

    return frame


def import_view():
    global import_cbx
    global spm_ent
    global begin_ent
    global end_ent

    frame = Frame(window)

    Label(frame, text="Espm import SKPD 1.0", height="2", font=(
        "calibri", 20)).grid(column=0, row=0, sticky=W, padx=5)

    Label(frame, text="Import", height="2", font=(
        "calibri", 15)).grid(column=0, row=1, sticky=W, padx=5)

    username_text = "Username: " + user["username"]
    skpd_text = "SKPD: " + skpd["name"]

    username_lbl = Label(frame, text=username_text)
    username_lbl.grid(row=2, column=0, sticky=W, padx=5)
    skpd_lbl = Label(frame, text=skpd_text)
    skpd_lbl.grid(row=3, column=0, sticky=W, padx=5)

    Label(frame, text="", height="1").grid(column=0, row=4, sticky=W, padx=5)

    import_lbl = Label(frame, text="Pilih data yang ingin di Import")
    import_lbl.grid(row=5, column=0, sticky=W, padx=5)
    import_cbx = ttk.Combobox(frame)
    import_cbx["values"] = (
        "Anggaran",
        "Anggaran Triwulan",
        "Spm"
    )
    import_cbx.grid(row=6, column=0, sticky=W, padx=5)

    begin_lbl = Label(frame, text="Tanggal Awal *contoh: (2019-01-01)")
    begin_lbl.grid(row=7, column=0, sticky=W, padx=5)
    begin_ent = Entry(frame, width=50)
    begin_ent.grid(row=8, column=0, sticky=W, padx=5)

    end_lbl = Label(frame, text="Tanggal Akhir *contoh: (2019-12-31)")
    end_lbl.grid(row=9, column=0, sticky=W, padx=5)
    end_ent = Entry(frame, width=50)
    end_ent.grid(row=10, column=0, sticky=W, padx=5)

    spm_lbl = Label(frame, text="Nomor SPM (Optional)")
    spm_lbl.grid(row=11, column=0, sticky=W, padx=5)
    spm_ent = Entry(frame, width=50)
    spm_ent.grid(row=12, column=0, sticky=W, padx=5)

    Label(frame, text="", height="1").grid(column=0, row=13, sticky=W, padx=5)

    import_btn = Button(frame, text="import", command=validate_import)
    import_btn.grid(row=14, column=0, sticky=E, padx=5)

    return frame

# MAIN


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
    window.geometry("700x500")

    login_frame = login_view()
    login_frame.pack(fill="both", expand=True)

    main_loop.run_until_complete(run_tk(window))
