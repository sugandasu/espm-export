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


def reset_file():
    if os.path.exists('file.xlsx'):
        os.remove('file.xlsx')
        return False


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


async def send_request_file(url, rows, rowscount, fieldnames):
    reset_file()

    workbook = xlsxwriter.Workbook('file.xlsx')
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
            Label(window, text=str(rowcount) + " data").grid(row=6, column=1, sticky=E)
            workbook.close()
            await send_request(url)

            reset_file()
            workbook = xlsxwriter.Workbook('file.xlsx')
            worksheet = workbook.add_worksheet()
            
            col = 0
            for fieldname in fieldnames:
                worksheet.write(0, col, fieldname)
                col += 1
            
            row = 0

        row += 1
        rowcount += 1
    
    messagebox.showinfo("Import", "Import selesai")


async def send_request(url):
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    data = FormData()
    data.add_field('file',
                   open('file.xlsx', 'rb'),
                   filename='file.xlsx',
                   content_type='application/vnd.ms-excel')

    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data, headers=headers) as response:
            data = {
                "message": "failed"
            }
            if response.status == 200:
                data = await response.json()

            print(await response.json())


async def import_bank_account():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ref_Rek_5]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ref_Rek_5]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/bank-account"
    fieldnames = [
        "kd_rek_1", "kd_rek_2", "kd_rek_3",
        "kd_rek_4", "kd_rek_5", "nm_rek_5"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_budget():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_RASK_Arsip]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_RASK_Arsip]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/budget"
    fieldnames = [
        "tahun", "kd_perubahan", "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "kd_prog", "id_prog", "kd_keg", "kd_rek_1", "kd_rek_2",
        "kd_rek_3", "kd_rek_4", "kd_rek_5", "no_rinc", "no_id", "keterangan_rinc", "sat_1",
        "nilai_1", "sat_2", "nilai_2", "sat_3", "nilai_3", "satuan123",
        "jml_satuan", "nilai_rp", "total", "keterangan", "kd_ap_pub", "kd_sumber", "datecreate"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_budget_tw():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_Rencana_Arsip]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_Rencana_Arsip]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/budget-tw"
    fieldnames = [
        "tahun", "kd_perubahan", "kd_urusan", "kd_bidang",
        "kd_unit", "kd_sub", "kd_prog", "id_prog",
        "kd_keg", "kd_rek_1", "kd_rek_2", "kd_rek_3",
        "kd_rek_4", "kd_rek_5", "jan", "feb",
        "mar", "apr", "mei", "jun",
        "jul", "agt", "sep", "okt",
        "nop", "des"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_event():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_Kegiatan]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_Kegiatan]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/event"
    fieldnames = [
        "tahun", "kd_urusan", "kd_bidang", "kd_unit", "kd_sub",
        "kd_prog", "id_prog", "kd_keg", "ket_kegiatan"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_fund_source():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ref_Sumber_Dana]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ref_Sumber_Dana]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/fund-source"
    fieldnames = [
        "kd_sumber", "nm_sumber"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_skpd():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ref_Sub_Unit]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ref_Sub_Unit]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/skpd"
    fieldnames = [
        "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "nm_sub_unit"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_sp2d():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SP2D]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SP2D]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/sp2d"
    fieldnames = [
        "tahun", "no_sp2d", "no_spm", "tgl_sp2d",
        "kd_bank", "no_bku", "nm_penandatangan", "nip_penandatangan",
        "jbt_penandatangan", "keterangan"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_spm():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm"
    fieldnames = [
        "tahun", "no_spm", "kd_urusan", "kd_bidang", "kd_unit",
        "kd_sub", "no_spp", "jn_spm", "tgl_spm", "uraian",
        "nm_penerima", "bank_penerima", "rek_penerima", "npwp",
        "bank_pembayar", "nm_penandatangan", "nip_penandatangan",
        "jbt_penandatangan", "kd_edit"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_spm_detail():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Rinc]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Rinc]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm-detail"
    fieldnames = [
        "tahun", "no_spm", "no_id", "kd_urusan", "kd_bidang",
        "kd_unit", "kd_sub", "kd_prog", "id_prog", "kd_keg", "kd_rek_1",
        "kd_rek_2", "kd_rek_3", "kd_rek_4", "kd_rek_5", "nilai"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_spm_tax_account():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Pot]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ta_SPM_Pot]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/spm-tax-account"
    fieldnames = [
        "tahun", "no_spm", "kd_pot_rek", "nilai"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_tax():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ref_Pot_SPM]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ref_Pot_SPM]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/tax"
    fieldnames = [
        "kd_pot", "nm_pot", "kd_map",
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_tax_account():
    query = "SELECT * FROM [Simda_2019].[dbo].[Ref_Pot_SPM_Rek]"
    querycount = "SELECT COUNT(*) FROM [Simda_2019].[dbo].[Ref_Pot_SPM_Rek]"
    rows, rowscount = get_data(query, querycount)
    url = "http://espm.test/api/import-app/tax-account"
    fieldnames = [
        "kd_pot_rek", "kd_rek_1", "kd_rek_2", "kd_rek_3",
        "kd_rek_4", "kd_rek_5", "kd_pot"
    ]

    await send_request_file(url, rows, rowscount, fieldnames)


async def import_databases(import_data):
    print(import_data)
    Label(window, text="Sedang Import").grid(row=5, column=1, sticky=E)

    if import_data == "Anggaran":
        data = await import_budget()
    elif import_data == "Kegiatan":
        data = await import_event()
    elif import_data == "Pajak":
        data = await import_tax()
    elif import_data == "Pajak Spm":
        data = await import_spm_tax_account()
    elif import_data == "Rekening":
        data = await import_bank_account()
    elif import_data == "Rekening Pajak":
        data = await import_tax_account()
    elif import_data == "Rencana Anggaran":
        data = await import_budget_tw()
    elif import_data == "Skpd":
        data = await import_skpd()
    elif import_data == "Sp2d":
        data = await import_sp2d()
    elif import_data == "Spm":
        data = await import_spm()
    elif import_data == "Spm Detail":
        data = await import_spm_detail()
    elif import_data == "Sumber Dana":
        data = await import_fund_source()

    Label(window, text="Import selesai").grid(row=5, column=1, sticky=E)

    return data


def import_request():
    current = import_combobox.current()
    import_data = import_combobox.get()

    if (current >= 0 and current <= 11):
        future = asyncio.create_task(import_databases(import_data))

        try:
            main_loop.run_until_complete(future)
        except:
            pass
    else:
        messagebox.showerror("Import", "Data import tidak valid")


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
        req = requests.post("https://arcane-hamlet-59445.herokuapp.com/api/auth/login",
                            json=payload, headers=headers)
    except:
        messagebox.showerror("Login", "login gagal")

    if req.status_code == 200:
        req_json = req.json()
        user = req_json["user"]
        access_token = req_json["access_token"]

        reset_form()
        import_form()
    else:
        messagebox.showerror("Login", "login gagal")


def import_form():
    global user_lbl
    global spm_number_lbl
    global last_import_lbl
    global import_combobox
    global import_btn

    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    user_lbl = Label(window, text="Pengguna: " + user["username"])
    user_lbl.grid(row=1, column=0, sticky=W, padx=4)

    try:
        # req = requests.get(
        #     "http://espm.test/api/import-app/meta", headers=headers)

        # req_json = req.json()
        # spm_count = req_json["spm_count"]
        # last_import = req_json["last_import"]

        # text = "Jumlah Spm: " + str(spm_count)
        # spm_number_lbl = Label(window, text=text)
        # spm_number_lbl.grid(row=2, column=0, sticky=W, padx=4)

        # text = "Tanggal Terakhir Import: " + str(last_import)
        # last_import_lbl = Label(window, text=text)
        # last_import_lbl.grid(row=3, column=0, sticky=W, padx=4)

        text = "Pilih Data yang ingin diimport"
        import_combobox_lbl = Label(window, text=text)
        import_combobox_lbl.grid(row=2, column=1, sticky=E)
        import_combobox = ttk.Combobox(window)
        import_combobox["values"] = (
            "Anggaran",
            "Kegiatan",
            "Pajak",
            "Pajak Spm",
            "Rekening",
            "Rekening Pajak",
            "Rencana Anggaran",
            "Skpd",
            "Spm",
            "Spm Detail",
            "Sumber Dana",
        )
        import_combobox.grid(row=3, column=1, sticky=E)

        import_btn = Button(window, text="import", command=import_request)
        import_btn.grid(row=4, column=1, sticky=E, pady=4)
    except:
        pass


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
    window.geometry("700x300")

    Label(window, text="Espm import 1.0", height="1", font=(
        "calibri", 20)).grid(row=0, column=0, sticky=W, pady=10)

    login_form()

    main_loop.run_until_complete(run_tk(window))
