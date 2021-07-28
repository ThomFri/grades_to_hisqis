import json
import os
import subprocess

import pandas as pandas
import openpyxl
import xlwt
from xlrd import open_workbook
from os import listdir
from os.path import isfile, join
import dateutil.parser
from datetime import datetime
import numpy as np
from shutil import copyfile
import xlsxwriter
from xlutils.copy import copy
import math
from enum import Enum
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import easygui

#gui
#https://stackoverflow.com/questions/45441885/how-can-i-create-a-dropdown-menu-from-a-list-in-tkinter
#https://www.delftstack.com/de/howto/python-tkinter/how-to-create-dropdown-menu-in-tkinter/
#https://docs.python.org/3/library/dialog.html
#https://pythonbasics.org/tkinter-filedialog/


# headers etc.
class Hdrs(Enum):
    MNR = "mtknr"
    ABS = "abschl"
    STG = "stg"
    NNA = "nachname"
    VNA = "vorname"
    BEW = "bewertung"
    PDA = "pdatum"
    RES = "res1"
    PST = "pstatus"
    POR = "pordnr"
    LNR = "labnr"

class Grd(Enum):
    NAN = "NAN"
    KAN = "KAN"

class Edg(Enum):
    START = "startHISsheet"
    END = "endHISsheet"

key = Hdrs.MNR
req_cols = {
        Hdrs.MNR: {
            "name": "Matr.Nr."
        },
        Hdrs.BEW: {
            "name": "Bewertung"
        },
        Hdrs.PDA: {
            "name": "Prüfungsdatum",
            "special": "col/fixed"
        }
    }

use_tkinter = False
use_easygui = True
use_file_picker = True

def find_in_workbook(wb, needle, skiprows=0):
    result = []
    sheet = wb.sheet_by_index(0)
    for rowidx in range(skiprows, sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if cell.value == needle:
                tmp_result = {
                    "col": colidx,
                    "row": rowidx
                }
                result.append(tmp_result)

    return result

def get_input_int(text, range=None):
    print("" + str(text) + " ")
    result = input()
    while True:
        if not result.isnumeric():
            print("Eingabe ist keine (nicht negative) Zahl, bitte korrigieren!")
            result = input()
        else:
            result = int(result)
            if not range is None:
                if not result in range:
                    print("Eingabe ist nicht im Bereich von " + range_edges_to_string(range) + ", bitte korrigieren!")
                    result = input()
                else:
                    break
            else:
                break
                
    return result

def list_to_string_with_leading_index(list, start=0):
    result = ""
    for item in list:
        result += str(start) + "\t" + str(item) + "\n"
        start += 1

    result = result[:len(result)-1]

    return result

def join_non_strings(join_string, iteratable):
    result = ""
    for item in iteratable:
        result += join_string + str(item)

    result = result[len(join_string):]

    return result

def range_edges_to_string(range):
    result = ""
    range_list = list(range)
    min = range_list[0]
    max = range_list[0]

    for i in range:
        if min > i:
            min = i
        if max < i:
            max = i


    result = "["+str(min)+", "+str(max)+"]"
    return result

def file_selector(text, path=".", show_files=True, gui=use_file_picker):
    result = None

    while True:
        if use_file_picker:
            print(text)
            print("<<ENTER-Taste drücken um Dialog zu öffnen>>")
            tmp = input()


            if use_tkinter:
                # input_types = [('xls', '*.xls'), ('xlsx', '*.xlsx')]
                result = askopenfilename(title=text)
            if use_easygui:
                input_types = [["*.xls", "*.xlsx", "Excel Datei"]]
                result = easygui.fileopenbox(title=text, filetypes=input_types, multiple=False)


        else:
            files_in_path = [f for f in listdir(path) if isfile(join(path, f))]
            files_in_path.append("<<Keine hiervon>>")
            if show_files:
                print("Dateien im Verzeichnis \"" + path + "\": \n" + list_to_string_with_leading_index(files_in_path) + "\n")

            while True:
                file_index = get_input_int(text + " [Nummer auf der linken Seite!]:", range(len(files_in_path)))
                if file_index == len(files_in_path) - 1:
                    print("Geben Sie den gesamten Dateipfand an:")
                    result = input()
                else:
                    result = files_in_path[file_index]

                if not isfile(result):
                    print("Ungültiger Pfad!")
                else:
                    break

        if not result is None and isfile(result):
            break
        else:
            print("Keine gültige Datei gewählt!")

    return result


def clean_dataframe(df, cleaing_col, cleaning_set):
    to_delte = df[df[cleaing_col].isin(cleaning_set)].index
    df.drop(to_delte, inplace=True)
    df.reset_index(drop=True, inplace=True)

if __name__ == '__main__':
    print(
        "╔══════════════════════════╗" + "\n"
        "║  GRADES ==TO==> HISQIS   ║" + "\n"
        "╚══════════════════════════╝" + "\n\n\n"
    )

    if use_tkinter:
        gui = Tk()
        gui.withdraw()

    hq_file = file_selector("Bitte HisQis-Datei auswählen")
    print("\n")
    hq_wb = open_workbook(hq_file)
    tab_corners = {}
    tab_corners["start"] = find_in_workbook(hq_wb, Edg.START.value, skiprows=0)
    tab_corners["end"] = find_in_workbook(hq_wb, Edg.END.value, skiprows=0)


    skip_rows_till_tab = tab_corners["start"][0]["row"] + 1
    table_length = tab_corners["end"][1]["row"] - skip_rows_till_tab - 1
    hq_df = pandas.read_excel(hq_file, skiprows=skip_rows_till_tab, nrows=table_length)

    hq_key_name = key.value
    hq_default_key_col = 0
    if hq_key_name in hq_df.head():
        hq_df.set_index(hq_key_name)
        hq_index_col = hq_key_name
    else:
        hq_index_col = hq_df.columns[hq_default_key_col]
        hq_df.set_index([hq_index_col])

    hq_matrnr_set = set(hq_df[hq_index_col])


    own_file = file_selector("Bitte eigene Datei auswählen", show_files=False)
    own_df = pandas.read_excel(own_file, header=None)
    print("\n\n")


    print(own_df.head(10))

    print("\n")
    skip_rows_own = get_input_int("Bei welcher Zeilenzahl (links) beginnt Ihre Tabelle bzw. wo befindet sich der Tabellenkopf?")
    print("\n")

    own_df = pandas.read_excel(own_file, skiprows=skip_rows_own)
    print("Die Tabelle enhält folgende Spalten:")
    print(list_to_string_with_leading_index(own_df.columns))
    print("\n")

    own_cols = {}
    fixed_values = {}

    for req_col in req_cols:
        special = req_cols[req_col].get("special")
        ask_col = True

        if special == "col/fixed":
            special_input = get_input_int("Hat Ihre Tabelle eine \"" + req_cols[req_col]["name"] + "\"-Spalte? [1 = ja, 0 = nein]", [0,1])
            print("\n")

            if int(special_input) == 0:
                ask_col = False

                own_cols[req_col] = req_col

                print("Bitte geben Sie den Festwert für \"" + req_cols[req_col]["name"] + "\" ein:")
                fixed_values[req_col] = input()
                print("\n")

        if ask_col:
            col_index = get_input_int("Was ist die Nummer (links) Ihrer \"" + req_cols[req_col]["name"] + "\"-Spalte?", range(len(own_df.columns)))
            print("\n")
            own_cols[req_col] = own_df.columns[col_index]


    print(own_df.tail(10))
    last_rows_own = get_input_int("Bei welcher Zeilenzahl (links) endet Ihre Tabelle?", range(len(own_df)))
    
    nrows_own = last_rows_own + 1

    own_df = pandas.read_excel(own_file, skiprows=skip_rows_own, nrows=nrows_own)

    #renaming own_df
    inverted_own_cols = {}
    for k, v in own_cols.items():
        inverted_own_cols[v] = k.value

    own_df = own_df.rename(columns=inverted_own_cols)
    own_df.set_index(key.value)
    if Hdrs.PDA in fixed_values:
        own_df[Hdrs.PDA.value] = fixed_values[Hdrs.PDA]


    own_matrnr_set = set(own_df[key.value])
    set_diff = hq_matrnr_set ^ own_matrnr_set

    if not len(set_diff) == 0:
        add_hq = hq_matrnr_set - own_matrnr_set
        add_own = own_matrnr_set - hq_matrnr_set

        print("\n\n")
        print("╔══════════════════════════════════════════════════════════════════════╗")
        print("║ WARNUNG!!!                                                           ║")
        print("║ ----------                                                           ║")
        print("║ Matrikelnummern stimmen nicht überein!                               ║")
        print("║                                                                      ║")
        print("║ Zusätzliche in HisQis-Datei: " + join_non_strings(", ",add_hq))
        print("║ Zusätzliche in eigener Datei: " + join_non_strings(", ",add_own))
        print("╚══════════════════════════════════════════════════════════════════════╝")
        print("\n")

        ignore_options = [
            "Nur die Matr.-Nr. der HisQis-Datei berücksichtigen",
            "Nur die Matr.-Nr. der eigenen Datei berücksichtigen",
            "Nur die Matr.-Nr. berücksichtigen, die in beiden Dateien enthalten sind",
            "Die Matr.-Nr. aus beiden Dateien berücksichtigen",
            "Mehr Details anzeigen"
        ]

        while True:
            do_loop = False
            do_ignore = get_input_int("Wie soll hiermit verfahren werden?" + "\n" +
                                      list_to_string_with_leading_index(ignore_options),
                                      range(len(ignore_options))
                                      )

            if do_ignore == 0 or do_ignore == 2:
                clean_dataframe(own_df, key.value, add_own)
            if do_ignore == 1 or do_ignore == 2:
                clean_dataframe(hq_df, key.value, add_hq)
            if do_ignore == 4:
                print("\n\n")

                print("Zusätzlich in HisQis-Datei:")
                print("===========================")
                if len(add_hq) > 0:
                    print(hq_df[hq_df[key.value].isin(add_hq)].head())
                else:
                    print("Keine zusätzlichen Daten")

                print("\n\n")

                print("Zusätzlich in eigener Datei:")
                print("============================")
                if len(add_own) > 0:
                    print(own_df[own_df[key.value].isin(add_own)].head())
                else:
                    print("Keine zusätzlichen Daten")

                print("\n\n")

                do_loop = True

            if not do_loop:
                break

    print("\n" + "Daten abgleichen..." + "\n")
    original_header = hq_df.columns
    hq_df.drop(columns=[Hdrs.BEW.value, Hdrs.PDA.value], inplace=True)
    own_df = own_df[[key.value, Hdrs.BEW.value, Hdrs.PDA.value]]
    merged_dataframe = pandas.merge(hq_df, own_df, on=key.value, how="outer")
    merged_dataframe = merged_dataframe[original_header]

    merged_dataframe[Hdrs.BEW.value] = merged_dataframe[Hdrs.BEW.value].apply(
        lambda x: x.upper()
        if isinstance(x,str)
        else x
    )

    grade_mean = pandas.to_numeric(merged_dataframe[Hdrs.BEW.value], errors='coerce').mean()
    if grade_mean < 6:
        #needs multiplication!
        merged_dataframe[Hdrs.BEW.value] = merged_dataframe[Hdrs.BEW.value].apply(
            lambda x: x*100
            if not isinstance(x,str)
            else x
        )

    bewertung_contains_nan = (merged_dataframe[Hdrs.BEW.value].isnull().sum() > 0)

    if bewertung_contains_nan:
        bewertung_options = [
            "Ignorieren",
            "Durch \""+Grd.NAN.value+"\" ersetzen",
            "Durch \""+Grd.KAN.value+"\" ersetzen"
        ]
        do_bewertung = get_input_int("\"bewertung\"-Spalte enthälte leere Werte! Wie soll mit diesen verfahren werden?" + "\n" +
                                     list_to_string_with_leading_index(bewertung_options),
                                     range(len(bewertung_options))
                                     )
        print("\n")

        if do_bewertung == 1:
            merged_dataframe[Hdrs.BEW.value].replace(np.nan, Grd.NAN.value, regex=True, inplace=True)
        elif do_bewertung == 2:
            merged_dataframe[Hdrs.BEW.value].replace(np.nan, Grd.KAN.value, regex=True, inplace=True)

    merged_dataframe[Hdrs.PDA.value] = merged_dataframe[Hdrs.PDA.value].apply(
        lambda x: dateutil.parser.parse(str(x)).strftime("%d.%m.%Y")
        if (np.all(pandas.notnull(x)))
        else x
    )

    do_target = get_input_int("Ergebnis direkt in HisQis-Datei schreiben? [1 = ja, 0 = nein, Kopie anlegen]", [0,1])
    if do_target == 0:
        last_dot = hq_file.rfind('.')
        target_file = hq_file[:last_dot] + "_upload" + hq_file[last_dot:]

        if use_file_picker:
            if use_tkinter:
                save_types = [('xls', '*.xls')]
                target_file = asksaveasfilename(title="Upload-Datei speichern unter...", filetypes=save_types, defaultextension=save_types, initialfile=target_file)
            if use_easygui:
                save_types = [["*.xls", "Excel Datei"]]
                target_file = easygui.filesavebox(title="Upload-Datei speichern unter...", filetypes=save_types, default=target_file)

        else:
            pass

        copyfile(hq_file, target_file)

    else:
        target_file = hq_file


    write_start = tab_corners["start"][0]["row"] + 1
    write_end_row = write_start + len(merged_dataframe) + 3
    write_end_col = tab_corners["end"][0]["col"] + 1

    target_wb_tmp = open_workbook(target_file, formatting_info=True)
    target_sheet_tmp = target_wb_tmp.sheet_by_index(0)
    target_wb = copy(target_wb_tmp)
    target_sheet = target_wb.get_sheet(0)
    nan_format = xlwt.XFStyle()
    nan_pattern = xlwt.Pattern()
    nan_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    nan_pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
    nan_format.pattern = nan_pattern

    row_i = write_start
    #for col, val in enumerate(merged_dataframe.columns, start=0):
    #    target_sheet.write(row_i, col, val)
    row_i += 1

    for index, row_content in merged_dataframe.iterrows():
        for col, val in enumerate(row_content, start=0):
            if isinstance(val, float) and math.isnan(val):
                target_sheet.write(row_i, col, style=nan_format)
            else:
                target_sheet.write(row_i, col, val)
        row_i += 1

    real_end_row = row_i
    real_end_col = tab_corners["start"][0]["col"]

    if write_end_row > row_i:
        for row_i in range(row_i, write_end_row):
            for col_i in range(0, write_end_col):
                target_sheet.write(row_i, col_i)

    target_sheet.write(real_end_row, real_end_col, Edg.END.value)

    target_wb.save(target_file)

    print("\n" + "FERTIG!")
    print("Sie können die Datei \"" + target_file + "\" jetzt auf HisQis hochladen.")
    print("\n")

    do_open_file = get_input_int("Datei zur Kontrolle öffnen? [1 = ja, 0 = nein]", [0, 1])

    if do_open_file == 1:
        os.startfile(target_file)

    pass


