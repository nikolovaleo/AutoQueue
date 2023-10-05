# from get_que import get_available_users
# import numpy as np
import pandas as pd
import numpy as np

from datetime import date, datetime
from math import isnan, trunc


from tkinter.filedialog import askopenfilename
from os import system, name, getcwd, getlogin
from shutil import rmtree

# import os

# import win32com.client as win32
from win32com.client.gencache import EnsureDispatch

from pathlib import Path


from unicodedata import normalize, combining

# import subprocess
from subprocess import run


import tkinter as tk
from tkinter.filedialog import askopenfilename

# ----------------------------------------------------------------------------------
# -------------- Empiezan las Funciones --------------------------------------------
# ----------------------------------------------------------------------------------


def clear():
    system("cls" if name == "nt" else "clear")
    # os.system('cls' if os.name == 'nt' else 'clear')


def delete_chache():
    actual_directory = getcwd()

    user_id = actual_directory.strip().split("\\")[2]
    path_to_del = f"C:/Users/{user_id}/AppData/Local/Temp/gen_py"

    try:
        rmtree(path_to_del)
    except:
        pass


def path_selector_rawdata():
    clear()
    input(
        "\n\n\nSelect the RawData File:   \n\n\n \t\tNew Master Data for Pizarra (xx).xlsx \n\n\n\nPress ENTER to continue...     "
    )

    filename_rawdata = askopenfilename()

    clear()
    replay = str(
        input(
            f"\n\n\nFile Path selected: \n\n \t[[{filename_rawdata}]] \n\n\n Do you want to change it?     (y/n):                      or press ENTER to continue\n\t "
        )
    )

    if replay == "y" or replay == "yes":
        input(
            "\n\n\nSelect the RawData File:   \n\n\n \t\tNew Master Data for Pizarra (xx).xlsx \n\n\n\nPress ENTER to continue...     "
        )
        filename_rawdata = askopenfilename()

        return filename_rawdata

    else:
        return filename_rawdata


def path_selector_schedule():
    clear()
    input(
        "\n\n\nSelect the Schedule File:   \n\n\n \t\t2023 EECR Schedule R1.xlsx \n\n\n\nPress ENTER to continue...     "
    )
    filename_schedule = askopenfilename()

    clear()
    replay = str(
        input(
            f"\n\n\nFile Path selected: \n\n \t[[{filename_schedule}]] \n\n\n Do you want to change it?     (y/n):                      or press ENTER to continue\n\t "
        )
    )

    #

    if replay == "y" or replay == "yes":
        input(
            "\n\n\nSelect the Schedule File:   \n\n\n \t\t2022 EECR Schedule R1.xlsx \n\n\n\nPress ENTER to continue...     "
        )
        filename_schedule = askopenfilename()
        return filename_schedule

    else:
        return filename_schedule


def save_filename_rawdata(filename_rawdata):
    f = open("./data/path_rawdata.txt", "w")
    f.write(filename_rawdata)
    f.close()


def load_list(path):
    res = []
    lines = open(path, encoding="utf-8").readlines()
    for line in lines:
        x_list = line.strip().split(" ")

        first_name = x_list[0]

        last_name = x_list[1]
        res.append(f"{first_name} {last_name}")
    return res


def load_path_txt(path):
    lines = open(path, encoding="utf-8").readlines()
    for line in lines:
        x_list = line.strip().split("\n")

        path = x_list[0]
    return str(path)


def remove_accents(input_str):
    nfkd_form = normalize("NFKD", input_str)
    return "".join([c for c in nfkd_form if not combining(c)])


def encontrar_y_agregar_old(name_on_table, available_users, Que_desordenado):
    estado_segundo_nombre = False

    if " [AUTOSOL/PWS/CR]" in name_on_table:
        name_on_table = name_on_table.replace(" [AUTOSOL/PWS/CR]", "")
    elif " [AUTOSOL/PWS/CORI]" in name_on_table:
        name_on_table = name_on_table.replace(" [AUTOSOL/PWS/CORI]", "")
    elif " [EMR/SYSS/PWS/GUAC]" in name_on_table:
        name_on_table = name_on_table.replace(" [EMR/SYSS/PWS/GUAC]", "")

    name_on_table = remove_accents(name_on_table)

    apellido = name_on_table.split(", ")[0]
    nombre_completo = name_on_table.split(", ")[1].split(" ")

    if len(nombre_completo) == 1:
        primer_nombre = nombre_completo[0]

    elif len(nombre_completo) == 2:
        primer_nombre = nombre_completo[0]
        segundo_nombre = nombre_completo[1]

        estado_segundo_nombre = True

    for user in available_users:
        # print('user: ', user)
        if estado_segundo_nombre:
            if (
                (apellido in remove_accents(user))
                and (primer_nombre in remove_accents(user))
            ) or (
                (apellido in remove_accents(user))
                and (segundo_nombre in remove_accents(user))
            ):
                # print('name_on_table: ', primer_nombre, segundo_nombre, apellido, '\n')
                Que_desordenado.append(user)
                return Que_desordenado
        elif not (estado_segundo_nombre):
            if (apellido in remove_accents(user)) and (
                primer_nombre in remove_accents(user)
            ):
                # print('name_on_table: ',primer_nombre, apellido, '\n')
                Que_desordenado.append(user)
                return Que_desordenado
    return False


def encontrar_y_agregar(name_on_table, available_users, Que_desordenado):
    estado_segundo_nombre = False
    estado_segundo_apellido = True

    name_on_table = name_on_table.split("]")[0]
    if " [AUTOSOL/PWS/CR" in name_on_table:
        name_on_table = name_on_table.replace(" [AUTOSOL/PWS/CR", "")
    elif " [AUTOSOL/PWS/CORI" in name_on_table:
        name_on_table = name_on_table.replace(" [AUTOSOL/PWS/CORI", "")
    elif " [EMR/SYSS/PWS/GUAC" in name_on_table:
        name_on_table = name_on_table.replace(" [EMR/SYSS/PWS/GUAC", "")

    name_on_table = remove_accents(name_on_table)

    apellido = name_on_table.split(", ")[0]
    nombre_completo = name_on_table.split(", ")[1]  # .split(" ")

    df_bd_names = pd.read_excel("./data/lists/names_db.xlsx")

    for db_usuario in df_bd_names.iterrows():
        db_nombres = remove_accents(
            db_usuario[1]["Nombre Completo"].split(", ")[1]
        )  # .split(" ")
        db_apellidos = remove_accents(
            db_usuario[1]["Nombre Completo"].split(", ")[0])

        if (apellido in db_apellidos) and (nombre_completo in db_nombres):
            split_db_nombres = db_nombres.split(" ")
            spliut_db_apellidos = db_apellidos.split(" ")

            if len(split_db_nombres) == 1:
                db_primer_nombre = split_db_nombres[0]

            elif len(split_db_nombres) == 2:
                db_primer_nombre = split_db_nombres[0]
                db_segundo_nombre = split_db_nombres[1]
                estado_segundo_nombre = True

            if len(spliut_db_apellidos) == 1:
                db_primer_apellido = spliut_db_apellidos[0]
                estado_segundo_apellido = False

            elif len(spliut_db_apellidos) == 2:
                db_primer_apellido = spliut_db_apellidos[0]
                db_segundo_apellido = spliut_db_apellidos[1]

            for user_disp in available_users:
                user_disp = remove_accents(user_disp)

                # print(user_disp, db_nombres, db_apellidos)
                if estado_segundo_nombre:
                    if (db_primer_nombre in user_disp) or (
                        db_segundo_nombre in user_disp
                    ):
                        if estado_segundo_apellido:
                            if (db_primer_apellido in user_disp) or (
                                db_segundo_apellido in user_disp
                            ):
                                if ("(" in user_disp) and (")" in user_disp):
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]} ({user_disp.split("(")[1]}'
                                    )
                                    return Que_desordenado
                                else:
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]}'
                                    )
                                    return Que_desordenado

                        elif not (estado_segundo_apellido):
                            if db_primer_apellido in user_disp:
                                if ("(" in user_disp) and (")" in user_disp):
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]} ({user_disp.split("(")[1]}'
                                    )
                                    return Que_desordenado
                                else:
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]}'
                                    )
                                    return Que_desordenado

                elif not (estado_segundo_nombre):
                    if db_primer_nombre in user_disp:
                        if estado_segundo_apellido:
                            if (db_primer_apellido in user_disp) or (
                                db_segundo_apellido in user_disp
                            ):
                                if ("(" in user_disp) and (")" in user_disp):
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]} ({user_disp.split("(")[1]}'
                                    )
                                    return Que_desordenado
                                else:
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]}'
                                    )
                                    return Que_desordenado

                        elif not (estado_segundo_apellido):
                            if db_primer_apellido in user_disp:
                                if ("(" in user_disp) and (")" in user_disp):
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]} ({user_disp.split("(")[1]}'
                                    )
                                    return Que_desordenado
                                else:
                                    Que_desordenado.append(
                                        f'{db_usuario[1]["Preferencia"]}'
                                    )
                                    return Que_desordenado

    return False


def move_eladio(Que_desordenado):
    for index, user in enumerate(Que_desordenado):
        if "Eladio" in user:
            aux = user
            Que_desordenado.pop(index)
            Que_desordenado.append(aux)
        else:
            pass
    return Que_desordenado


def check_OCWW_log(date_object):
    path = "./data/OCWW_log.txt"
    lines = open(path).readlines()

    for line in lines:
        if str(date_object) in line:
            # print('ITS TRUE')
            return True
        else:
            pass
    # print('ITS FALSE')
    return False


# ----------------------------------------------------------------------------------
# -------------- Terminan las Funciones --------------------------------------------
# ----------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------
# -------------- Se declara la funcion join() - une todo ---------------------------
# ----------------------------------------------------------------------------------


def join():
    # --------------------------------------------------------------------------
    # -------------- Fase 1: read_schedule -------------------------------------
    # --------------------------------------------------------------------------

    delete_chache()

    current_user = remove_accents(getlogin())
    # print(current_user)

    # filename_schedule = path_selector_schedule()     # se llama la funcion para pedirle al usuario la ubicacion del archivo SCHEDULE
    filename_schedule = rf"C:\Users\{current_user}\Downloads\2023 EECR Schedule R1.xlsx"

    # filename_rawdata = path_selector_rawdata()      # se llama la funcion para pedirle al usuario la ubicacion del archivo RAWDATA
    filename_rawdata = (
        rf"C:\Users\{current_user}\Downloads\New Master Data for Pizarra.xlsx"
    )

    df = pd.read_excel(
        filename_schedule, sheet_name="Calendar"
    )  # Se carga el horario del site

    clear()
    print("\n\n\tReading Schedule...")
    date_object = date.today()  # Se obtiene la fecha de hoy y se crea en objeto time
    yyyy = date_object.year  # Se descompone la fecha en año, mes y dia
    mm = date_object.month
    dd = date_object.day

    day_of_week = date_object.weekday()

    # time_date_object = datetime(yyyy, 9, 16, 0, 0, 0)
    time_date_object = datetime(yyyy, mm, dd, 0, 0, 0)

    df_just_date = (
        df[df.isin([time_date_object])].stack().unstack()
    )  # Se compara/busca en el df los que cumplen dicha fecha de hoy
    tdy_column_string = df_just_date.columns[
        0
    ]  # Se obtiene el nombre de la columna responsable del dia de hoy == colum_string

    get_start = load_path_txt("./data/start_table.txt")
    get_end = load_path_txt("./data/end_table.txt")

    # Se dejan solo las filas con nombres del site
    start_id = np.where(df["Unnamed: 2"] == get_start)[0][0]
    finish_id = np.where(df["Unnamed: 2"] == get_end)[0][0] + 1

    df_W_names = df.drop(df.index[range(0, start_id)])
    df_W_names = df_W_names.drop(
        df.index[range(finish_id, df.index.stop)]
    )  # el numero 76 se tiene que estar modificando cada vez que se agrega gente

    # print(df_W_names)
    """
    print(start_id)
    print(finish_id)

    
    print(df_W_names)
    x = input("wait: ")
    """
    # df = df.reset_index()  # make sure indexes pair with number of rows

    dict_h_clasificado = {}

    blacklist = load_list("./data/lists/blacklist.txt") + [
        "Katherine  Alvarado",
        "Jose Carlos Marin",
    ]

    training_list = load_list("./data/lists/training_list.txt")

    shadowing_list = load_list("./data/lists/shadowing_list.txt")

    senior_list = load_list("./data/lists/senior_list.txt")

    proySec_list = load_list("./data/lists/proySec_list.txt")

    # SE CREA LA FUNCION QUE LEE LAS LISTAS EXCEPTIONS

    RO_list = []
    M_list = []
    L_list = []
    nan_list = []
    senior_available = []

    for index, row in df_W_names.iterrows():
        # print(row['Unnamed: 2'], row[tdy_column_string])

        if (
            (row["Unnamed: 2"] not in blacklist)
            and (row["Unnamed: 2"] not in training_list)
            and (row["Unnamed: 2"] not in shadowing_list)
            and (row["Unnamed: 2"] not in shadowing_list)
            and (row["Unnamed: 2"] not in senior_list)
            and (row["Unnamed: 2"] not in proySec_list)
        ):
            if row[tdy_column_string] == "RO":
                RO_list.append(row["Unnamed: 2"])

            elif row[tdy_column_string] == "M" or row[tdy_column_string] == "M RO":
                # M_list.append(row['Unnamed: 2'])
                M_list.append(row["Unnamed: 2"] + " (early)")
            elif row[tdy_column_string] == "L" or row[tdy_column_string] == "L RO":
                # L_list.append(row['Unnamed: 2'])
                L_list.append(row["Unnamed: 2"] + " (late)")

            elif type(row[tdy_column_string]) == float:
                if (
                    isnan(row[tdy_column_string])
                    and (row["Unnamed: 2"] not in blacklist)
                    and (row["Unnamed: 2"] not in training_list)
                    and (row["Unnamed: 2"] not in shadowing_list)
                    and (row["Unnamed: 2"] not in shadowing_list)
                    and (row["Unnamed: 2"] not in senior_list)
                    and (row["Unnamed: 2"] not in proySec_list)
                ):
                    nan_list.append(row["Unnamed: 2"])

        if row["Unnamed: 2"] in senior_list:
            if row[tdy_column_string] == "RO":
                senior_available.append(row["Unnamed: 2"])

            elif row[tdy_column_string] == "M" or row[tdy_column_string] == "M RO":
                # M_list.append(row['Unnamed: 2'])
                senior_available.append(row["Unnamed: 2"] + " (early)")
            elif row[tdy_column_string] == "L" or row[tdy_column_string] == "L RO":
                # L_list.append(row['Unnamed: 2'])
                senior_available.append(row["Unnamed: 2"] + " (late)")

            elif type(row[tdy_column_string]) == float:
                if (
                    isnan(row[tdy_column_string])
                    and (row["Unnamed: 2"] not in blacklist)
                    and (row["Unnamed: 2"] not in training_list)
                    and (row["Unnamed: 2"] not in shadowing_list)
                    and (row["Unnamed: 2"] not in shadowing_list)
                    and (row["Unnamed: 2"] not in proySec_list)
                ):
                    senior_available.append(row["Unnamed: 2"])

    global_list = RO_list + M_list + L_list + nan_list

    # ----------------------------------------------------------------------------------------------

    # filename1 = askopenfilename()
    # filename1 = load_path_txt("./data/path_rawdata.txt") #f1_name = r'\data\New Master Data for Pizarra (79).xlsx'
    filename1 = filename_rawdata

    # create excel object

    # excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel = EnsureDispatch("Excel.Application")

    # excel can be visible or not
    excel.Visible = True  # False #True  # False

    # open workbooks
    f_path = Path.cwd()  # your path

    # f2_name = r'\data\New Pizarra (07-27-22).xlsx'
    f2_name = r"\data\New Pizarra (09-01-22).xlsx"

    # filename1 = str(f_path) + f1_name
    filename2 = str(f_path) + f2_name
    # clear()

    print("\n\n\tOpening Files and Replacing values...")
    # wb1 = excel.Workbooks.Open(r'C:\Users\e1426744\Desktop\auto_list\data\New Pizarra (07-27-22).xlsx')
    wb1_raw = excel.Workbooks.Open(filename1)
    wb2_board = excel.Workbooks.Open(filename2)

    sheetname1_raw = "Sheet1"
    df_raw = pd.read_excel(filename1, sheet_name=sheetname1_raw)
    # print(df_raw.shape)

    sheetname2_board = "Raw Data"
    sheetname2_board_graphics = "Graphics"
    df_raw_board = pd.read_excel(filename2, sheet_name=sheetname2_board)

    # Vacia el sheet raw data del new pizzara.
    wb2_board.Sheets(sheetname2_board).Range(
        f"A2:U{df_raw_board.shape[0] + 1}"
    ).Value = None

    wb1_raw.Sheets(sheetname1_raw).Range(f"A2:U{df_raw.shape[0] + 1}").Copy(
        Destination=wb2_board.Sheets(sheetname2_board).Range(
            f"A2:U{df_raw.shape[0] + 1}"
        )
    )
    # wb1_raw.Sheets(sheetname1_raw).Range(f"A2:U11").Copy(Destination = wb2_board.Sheets(sheetname2_board).Range(f"A2:U11"))
    # day_of_week = 1
    ###
    ###
    # SE ACTUALIZA EL OCWW SI ES MARTES! Y NO ESTA EN EL LOG!!
    if day_of_week == 1:
        try:
            # print("check1.0")
            ocww_change = check_OCWW_log(date_object)
        except:
            # print("check1.1")
            ocww_change = False

        if ocww_change == False:
            # print("check2")
            ocww_cel = ["B13", "E13", "B34"]
            for cel in ocww_cel:
                # print("check3")
                OCWW_value = wb2_board.Sheets(
                    sheetname2_board_graphics).Range(cel)
                suma = int(OCWW_value) + 1
                OCWW_value.Value = suma

            # write in the log
            # write in the log
            # write in the log

            # print("check4")
            path = "./data/OCWW_log.txt"
            k = open(path, "a+")
            new_occww = str(
                wb2_board.Sheets(sheetname2_board_graphics).Range(ocww_cel[0])
            )
            k.write(f"{date_object}\t{new_occww}\n")
            k.close()

    else:
        pass
    # m = input()
    print("\n\n\tCreating list...")
    wb2_board.RefreshAll()

    # wb1_raw.SaveAs() # save as new Workbook
    wb2_board.Save()

    wb1_raw.Close(True)
    # wb2_board.Close(True)

    # excel.Quit()

    # ----------------------------------------------------------------------------------------------------------------------------------------------------------------

    print("\n\n\tPlease wait...")
    # Se carga el horario del site
    df = pd.read_excel(
        "./data/New Pizarra (09-01-22).xlsx", sheet_name=sheetname2_board_graphics
    )  # _New Pizarra ------------ _New Pizarra ------------ _New Pizarra ------------ _New Pizarra ------------

    global_limit_top = 30
    global_limit_bott = 75

    # Delimita de manera aproximada el espacio de los empleados de SureService
    df_table = df.drop(df.index[range(0, global_limit_top)])
    df_table = df_table.drop(df.index[range(global_limit_bott, df.index.stop)])

    available_users = global_list
    start = False
    end = False

    Que_desordenado = []

    for index, row in df_table.iterrows():
        if ((row["Unnamed: 0"]) == "Row Labels") or (
            (row["Unnamed: 0"]) == "Etiquetas de fila"
        ):
            start = True
        elif (
            ((row["Unnamed: 0"]) == "Grand Total")
            or ((row["Unnamed: 0"]) == "Total general")
            or ("Grand Total" in str(row["Unnamed: 0"]))
        ):
            end = True
        elif start and not (end):
            # print(type(name_on_table), name_on_table)
            name_on_table = str(row["Unnamed: 0"])

            aux = encontrar_y_agregar(
                name_on_table, available_users, Que_desordenado)

    Que_desordenado = move_eladio(Que_desordenado)

    late_list = []
    late_position = 1
    print(Que_desordenado)

    for order, user in enumerate(Que_desordenado):
        if "(late)" in user:
            full_user = user
            user = f"(L{late_position})"
            user_4_list = f'{full_user.split("(")[0]}(L{late_position})'
            late_list.append(user_4_list)
            late_position += 1
        else:
            pass
        # print(order, user, string_que)
        if order == 0:
            string_que = f"{user}"

        else:
            string_que = f"{string_que} > {user}"

    # print(f'Que_desordenado: {Que_desordenado} \n\n')
    string_que = f"{string_que} \n\nStart at 10:"
    print("test # 5")
    for late in late_list:
        string_que = f"{string_que}\n{late}"

    string_que = remove_accents(string_que)

    clear()

    # print(f'\n\n\n\n Q list: \n\n\n\n{string_que} \n')
    # run("clip", universal_newlines=True, input=string_que)

    advance_support = ""

    if bool(senior_available):
        string_que = f"{string_que} \n\nAvailable for advanced support:"

        for senior in senior_available:
            if "Rolando" in senior:
                senior = "Rolando"
            elif "Nathalie" in senior:
                senior = "Nathalie"
            elif "Hidalgo" in senior:
                senior = "Hidalgo"
            else:
                pass
            string_que = f"{string_que}\n{senior}"

    print(f"\n\n\n\n Q list: \n\n\n\n{string_que} \n")

    run("clip", universal_newlines=True, input=string_que)

    print("\n\n\n\nQUE COPIED TO THE CLIPBOARD. \tUse (CTRL + V) to pase the list.")
    # %%

    cc = open("AutoQue_list.txt", "w")
    cc.write(string_que)
    cc.close()

    print("\n\tSAVED AS:       <<   AutoQue_list.txt  >>\n")
    print(
        "\n--------------------------------------- END --------------------------------------- \n\n"
    )
    # -------------------------------------------------------------------------------------------------------------------------------


clear()


def menu():
    print("\n\n")
    print("\t[1] Generate Automatic Que.")
    # print("\t[2] Add/delete users exceptions.")
    print("\t[9] Exit the program.")


menu()
option = str(input("\n\tSelect a option:   "))

while "9" not in option:
    if "1" in option or "" in option:
        # ejecutar el get_que.py
        # os.system('python join.py')
        join()

    elif "2" in option:
        # ejecutar las eliminacion o agregacion de nombres a la lista
        print("option2")

    else:
        print("Invalid option.")

    menu()
    option = str(input("\n\tSelect a option:   "))


print("\n    Thanks for using the AutoQue_program! :)")
