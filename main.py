from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import requests
import pyodbc
import aiohttp
import asyncio

main_loop = asyncio.get_event_loop()
index = 0


def validated_row(row):
    for i in range(0, len(row)):
        if row[i] == None:
            row[i] = ''

    return row


async def send_request(session, url, headers, payload):
    async with session.post(url, json=payload, headers=headers) as response:
        data = {
            "message": "failed"
        }

        if response.status == 200:
            data = await response.json()

        print(await response.json())
        status["text"] = data["message"]


async def import_bank_account(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ref_Rek_5]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/bank-account"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)
            payload = {
                "kd_rek_1": str(row[0]),
                "kd_rek_2": str(row[1]),
                "kd_rek_3": str(row[2]),
                "kd_rek_4": str(row[3]),
                "kd_rek_5": str(row[4]),
                "nm_rek_5": str(row[5]),
            }

            await send_request(session, url, headers, payload)


async def import_budget(conn, headers):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM [Simda_2019].[dbo].[Ta_RASK_Arsip] ORDER BY [DateCreate]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/budget"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "kd_perubahan": str(row[1]),
                "kd_urusan": str(row[2]),
                "kd_bidang": str(row[3]),
                "kd_unit": str(row[4]),
                "kd_sub": str(row[5]),
                "kd_prog": str(row[6]),
                "id_prog": str(row[7]),
                "kd_keg": str(row[8]),
                "kd_rek_1": str(row[9]),
                "kd_rek_2": str(row[10]),
                "kd_rek_3": str(row[11]),
                "kd_rek_4": str(row[12]),
                "kd_rek_5": str(row[13]),
                "no_rinc": str(row[14]),
                "no_id": str(row[15]),
                "keterangan_rinc": str(row[16]),
                "sat_1": str(row[17]),
                "nilai_1": str(row[18]),
                "sat_2": str(row[19]),
                "nilai_2": str(row[20]),
                "sat_3": str(row[21]),
                "nilai_3": str(row[22]),
                "satuan123": str(row[23]),
                "jml_satuan": str(row[24]),
                "nilai_rp": str(row[25]),
                "total": str(row[26]),
                "keterangan": str(row[27]),
                "kd_ap_pub": str(row[28]),
                "kd_sumber": str(row[29]),
                "datecreate": str(row[30])
            }

            await send_request(session, url, headers, payload)


async def import_budget_tw(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_Rencana_Arsip]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/budget-tw"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "kd_perubahan": str(row[1]),
                "kd_urusan": str(row[2]),
                "kd_bidang": str(row[3]),
                "kd_unit": str(row[4]),
                "kd_sub": str(row[5]),
                "kd_prog": str(row[6]),
                "id_prog": str(row[7]),
                "kd_keg": str(row[8]),
                "kd_rek_1": str(row[9]),
                "kd_rek_2": str(row[10]),
                "kd_rek_3": str(row[11]),
                "kd_rek_4": str(row[12]),
                "kd_rek_5": str(row[13]),
                "jan": str(row[14]),
                "feb": str(row[15]),
                "mar": str(row[16]),
                "apr": str(row[17]),
                "mei": str(row[18]),
                "jun": str(row[19]),
                "jul": str(row[20]),
                "agt": str(row[21]),
                "sep": str(row[22]),
                "okt": str(row[23]),
                "nop": str(row[24]),
                "des": str(row[25]),
            }

            await send_request(session, url, headers, payload)


async def import_event(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_Kegiatan]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/event"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "kd_urusan": str(row[1]),
                "kd_bidang": str(row[2]),
                "kd_unit": str(row[3]),
                "kd_sub": str(row[4]),
                "kd_prog": str(row[5]),
                "id_prog": str(row[6]),
                "kd_keg": str(row[7]),
                "ket_kegiatan": str(row[8])
            }

            await send_request(session, url, headers, payload)


async def import_fund_source(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ref_Sumber_Dana]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/fund-source"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "kd_sumber": str(row[0]),
                "nm_sumber": str(row[1]),
            }

            await send_request(session, url, headers, payload)


async def import_skpd(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ref_Sub_Unit]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/skpd"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "kd_urusan": str(row[0]),
                "kd_bidang": str(row[1]),
                "kd_unit": str(row[2]),
                "kd_sub": str(row[3]),
                "nm_sub_unit": str(row[4])
            }

            await send_request(session, url, headers, payload)


async def import_sp2d(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_SP2D]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/sp2d"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "no_sp2d": str(row[1]),
                "no_spm": str(row[2]),
                "tgl_sp2d": str(row[3]),
                "kd_bank": str(row[4]),
                "no_bku": str(row[5]),
                "nm_penandatangan": str(row[6]),
                "nip_penandatangan": str(row[7]),
                "jbt_penandatangan": str(row[8]),
                "keterangan": str(row[9]),
            }

            await send_request(session, url, headers, payload)


async def import_spm(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_SPM]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/spm"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "no_spm": str(row[1]),
                "kd_urusan": str(row[2]),
                "kd_bidang": str(row[3]),
                "kd_unit": str(row[4]),
                "kd_sub": str(row[5]),
                "no_spp": str(row[6]),
                "jn_spm": str(row[7]),
                "tgl_spm": str(row[8]),
                "uraian": str(row[9]),
                "nm_penerima": str(row[10]),
                "bank_penerima": str(row[11]),
                "rek_penerima": str(row[12]),
                "npwp": str(row[13]),
                "bank_pembayar": str(row[14]),
                "nm_penandatangan": str(row[16]),
                "nip_penandatangan": str(row[17]),
                "jbt_penandatangan": str(row[18]),
                "kd_edit": str(row[19])
            }

            await send_request(session, url, headers, payload)


async def import_spm_detail(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Rinc]")

    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/spm-detail"

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "no_spm": str(row[1]),
                "no_id": str(row[2]),
                "kd_urusan": str(row[3]),
                "kd_bidang": str(row[4]),
                "kd_unit": str(row[5]),
                "kd_sub": str(row[6]),
                "kd_prog": str(row[7]),
                "id_prog": str(row[8]),
                "kd_keg": str(row[9]),
                "kd_rek_1": str(row[10]),
                "kd_rek_2": str(row[11]),
                "kd_rek_3": str(row[12]),
                "kd_rek_4": str(row[13]),
                "kd_rek_5": str(row[14]),
                "nilai": str(row[15])
            }

            await send_request(session, url, headers, payload)


async def import_spm_tax_account(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ta_SPM_Pot]")
    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/spm-tax-account"
    responses = []

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "tahun": str(row[0]),
                "no_spm": str(row[1]),
                "kd_pot_rek": str(row[2]),
                "nilai": str(row[3])
            }

            responses.append(await send_request(session, url, headers, payload))
    return responses


async def import_tax(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ref_Pot_SPM]")
    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/tax"
    responses = []

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "kd_pot": str(row[0]),
                "nm_pot": str(row[1]),
                "kd_map": str(row[2]),
            }

            responses.append(await send_request(session, url, headers, payload))
    return responses


async def import_tax_account(conn, headers):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM [Simda_2019].[dbo].[Ref_Pot_SPM_Rek]")
    url = "https://arcane-hamlet-59445.herokuapp.com/api/import-app/tax-account"
    responses = []

    async with aiohttp.ClientSession() as session:
        for row in cursor:
            row = validated_row(row)

            payload = {
                "kd_pot_rek": str(row[0]),
                "kd_rek_1": str(row[1]),
                "kd_rek_2": str(row[2]),
                "kd_rek_3": str(row[3]),
                "kd_rek_4": str(row[4]),
                "kd_rek_5": str(row[5]),
                "kd_pot": str(row[6])
            }

            responses.append(await send_request(session, url, headers, payload))

    return responses


async def import_databases(import_data):
    conn = pyodbc.connect("Driver={SQL Server};"
                          "Server=.\SQLEXPRESS;"
                          "Database=Simda_2019;"
                          "Trusted_Connection=yes;")

    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Authorization": "Bearer " + access_token
    }

    print(import_data)

    if import_data == "Anggaran":
        data = await import_budget(conn, headers)
    elif import_data == "Kegiatan":
        data = await import_event(conn, headers)
    elif import_data == "Pajak":
        data = await import_tax(conn, headers)
    elif import_data == "Pajak Spm":
        data = await import_spm_tax_account(conn, headers)
    elif import_data == "Rekening":
        data = await import_bank_account(conn, headers)
    elif import_data == "Rekening Pajak":
        data = await import_tax_account(conn, headers)
    elif import_data == "Rencana Anggaran":
        data = await import_budget_tw(conn, headers)
    elif import_data == "Skpd":
        data = await import_skpd(conn, headers)
    elif import_data == "Sp2d":
        data = await import_sp2d(conn, headers)
    elif import_data == "Spm":
        data = await import_spm(conn, headers)
    elif import_data == "Spm Detail":
        data = await import_spm_detail(conn, headers)
    elif import_data == "Sumber Dana":
        data = await import_fund_source(conn, headers)

    Label(window, text="Import selesai").grid(
        row=3, column=2, sticky=W)
    messagebox.showinfo("Import", "Import berhasil")

    return data


def import_request():
    global status
    status = Label(window)
    status.grid(row=2, column=2, sticky=W)

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
        req = requests.get(
            "https://arcane-hamlet-59445.herokuapp.com/api/import-app/meta", headers=headers)

        req_json = req.json()
        spm_count = req_json["spm_count"]
        last_import = req_json["last_import"]

        text = "Jumlah Spm: " + str(spm_count)
        spm_number_lbl = Label(window, text=text)
        spm_number_lbl.grid(row=2, column=0, sticky=W, padx=4)

        text = "Tanggal Terakhir Import: " + str(last_import)
        last_import_lbl = Label(window, text=text)
        last_import_lbl.grid(row=3, column=0, sticky=W, padx=4)

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
            "Sp2d",
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
