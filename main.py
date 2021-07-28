import json
import pandas as pandas
import openpyxl
from xlrd import open_workbook
from os import listdir
from os.path import isfile, join
import dateutil.parser
from datetime import datetime
import numpy as np
from shutil import copyfile
import xlsxwriter
from xlutils.copy import copy

#https://stackoverflow.com/questions/11348347/find-non-common-elements-in-lists ==> .difference()

#gui
#https://stackoverflow.com/questions/45441885/how-can-i-create-a-dropdown-menu-from-a-list-in-tkinter
#https://www.delftstack.com/de/howto/python-tkinter/how-to-create-dropdown-menu-in-tkinter/
#https://docs.python.org/3/library/dialog.html
#https://pythonbasics.org/tkinter-filedialog/

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
    print(text + "", end="")
    result = input()
    while True:
        if not result.isnumeric():
            print("Eingabe ist keine (nicht negative) Zahl, bitte korrigieren!", end="")
            result = input()
        else:
            result = int(result)
            if not range is None:
                if not result in range:
                    print("Eingabe ist nicht im Bereich von " + range_edges_to_string(range) + ", bitte korrigieren!", end="")
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

def file_selector(text, path="."):
    result = None

    files_in_path = [f for f in listdir(path) if isfile(join(path, f))]
    files_in_path.append("<<Keine hiervon>>")
    print("Dateien im Verzeichnis \"" + path + "\": \n" + list_to_string_with_leading_index(files_in_path))

    while True:
        file_index = get_input_int(text, range(len(files_in_path)))
        if file_index == len(files_in_path) - 1:
            print("Geben Sie den gesamten Dateipfand an:")
            result = input()
        else:
            result = files_in_path[file_index]

        if not isfile(result):
            print("Ungültiger Pfad!")
        else:
            break

    return result


def clean_dataframe(df, cleaing_col, cleaning_set):
    to_delte = df[df[cleaing_col].isin(cleaning_set)].index
    df.drop(to_delte, inplace=True)
    df.reset_index(drop=True, inplace=True)


if __name__ == '__main__':

    hq_file = file_selector("Bitte Nummer (links) der HisQis-Datei angeben:")
    print("\n")
    hq_wb = open_workbook(hq_file)
    tab_corners = {}
    tab_corners["start"] = find_in_workbook(hq_wb, "startHISsheet", skiprows=0)
    tab_corners["end"] = find_in_workbook(hq_wb, "endHISsheet", skiprows=0)


    skip_rows_till_tab = tab_corners["start"][0]["row"] + 1
    table_length = tab_corners["end"][1]["row"] - skip_rows_till_tab - 1
    hq_df = pandas.read_excel(hq_file, skiprows=skip_rows_till_tab, nrows=table_length)

    hq_key_name = "mtknr"
    hq_default_key_col = 0
    if hq_key_name in hq_df.head():
        hq_df.set_index(hq_key_name)
        hq_index_col = hq_key_name
    else:
        hq_index_col = hq_df.columns[hq_default_key_col]
        hq_df.set_index([hq_index_col])

    hq_matrnr_set = set(hq_df[hq_index_col])


    own_file = file_selector("Bitte Nummer (links) der eigenen Datei angeben:")
    own_df = pandas.read_excel(own_file, header=None)
    print(own_df.head(10))

    skip_rows_own = get_input_int("Bei welcher Zeilenzahl (links) beginnt Ihre Tabelle bzw. wo befindet sich der Tabellenkopf?")
    print("\n\n")

    own_df = pandas.read_excel(own_file, skiprows=skip_rows_own)
    print("Die Tabelle enhält folgende Spalten:")
    print(list_to_string_with_leading_index(own_df.columns))
    print("\n")

    own_cols = {}
    req_cols = {
        "mtknr": {
            "name": "Matr.Nr."
        },
        "bewertung": {
            "name": "Bewertung"
        },
        "pdatum": {
            "name": "Prüfungsdatum",
            "special": "col/fixed"
        }
    }
    fixed_values = {}

    for req_col in req_cols:
        special = req_cols[req_col].get("special")
        ask_col = True

        if special == "col/fixed":
            special_input = get_input_int("Hat Ihre Tabelle eine \"" + req_cols[req_col]["name"] + "\"-Spalte? [1 = ja, 0 = nein]", [0,1])

            if int(special_input) == 0:
                ask_col = False

                print("Bitte geben Sie den Festwert für \"" + req_cols[req_col]["name"] + "\" ein:", end="")
                own_cols[req_col] = req_col
                fixed_values[req_col] = input()

        if ask_col:
            col_index = get_input_int("Was ist die Nummer (links) Ihrer \"" + req_cols[req_col]["name"] + "\"-Spalte?", range(len(own_df.columns)))
            own_cols[req_col] = own_df.columns[col_index]


    print(own_df.tail(10))
    last_rows_own = get_input_int("Bei welcher Zeilenzahl (links) endet Ihre Tabelle?", range(len(own_df)))
    
    nrows_own = last_rows_own + 1

    own_df = pandas.read_excel(own_file, skiprows=skip_rows_own, nrows=nrows_own)

    #renaming own_df
    inverted_own_cols = {v: k for k, v in own_cols.items()}
    own_df = own_df.rename(columns=inverted_own_cols)
    own_df.set_index("mtknr")
    if "pdatum" in fixed_values:
        own_df["pdatum"] = fixed_values["pdatum"]


    own_matrnr_set = set(own_df["mtknr"])
    set_diff = hq_matrnr_set ^ own_matrnr_set

    if not len(set_diff) == 0:
        print("WARNUNG!!!")
        print("Matrikelnummern stimmen nicht überein!")

        add_hq = hq_matrnr_set - own_matrnr_set
        print("Zusätzliche in HisQis-Datei: " + join_non_strings(", ",add_hq))

        add_own = own_matrnr_set - hq_matrnr_set
        print("Zusätzliche in eigener Datei: " + join_non_strings(", ",add_own))

        ignore_options = [
            "Nur die Matr.-Nr. der HisQis-Datei berücksichtigen",
            "Nur die Matr.-Nr. der eigenen Datei berücksichtigen",
            "Die Matr.-Nr. aus beiden Datei berücksichtigen",
        ]
        do_ignore = get_input_int("Wie soll hiermit verfahren werden?" + "\n" +
                                  list_to_string_with_leading_index(ignore_options),
                                  range(3)
                                  )

        if do_ignore == 0:
            clean_dataframe(own_df, "mtknr", add_own)
        elif do_ignore == 1:
            clean_dataframe(hq_df, "mtknr", add_hq)
        elif do_ignore == 2:
            pass

    print("Merging data...")
    original_header = hq_df.columns
    hq_df.drop(columns=["bewertung", "pdatum"], inplace=True)
    own_df = own_df[["mtknr", "bewertung", "pdatum"]]
    merged_dataframe = pandas.merge(hq_df, own_df, on="mtknr", how="outer")
    merged_dataframe = merged_dataframe[original_header]

    grade_mean = merged_dataframe["bewertung"].mean()
    if grade_mean < 6:
        #needs multiplication!
        merged_dataframe["bewertung"] = merged_dataframe["bewertung"]*100
    merged_dataframe["bewertung"].replace(np.nan, 'KNA', regex=True)

    merged_dataframe["pdatum"] = merged_dataframe["pdatum"].apply(
        lambda x: dateutil.parser.parse(str(x)).strftime("%d.%m.%Y")
        if (np.all(pandas.notnull(x))) else x
    )

    do_target = get_input_int("Ergebnis direkt in HisQis-Datei schreiben? [1 = ja, 2 = nein, Kopie anlegen]")
    if do_target == 2:
        last_dot = hq_file.rfind('.')
        target_file = hq_file[:last_dot] + "_upload" + hq_file[last_dot:]
        copyfile(hq_file, target_file)
    else:
        target_file = hq_file


    write_start = tab_corners["start"][0]["row"] + 1
    write_end = write_start + len(merged_dataframe)

    target_wb_tmp = open_workbook(target_file, formatting_info=True)
    target_sheet_tmp = target_wb_tmp.sheet_by_index(0)
    target_wb = copy(target_wb_tmp)
    target_sheet = target_wb.get_sheet(0)


    row_i = write_start
    #for col, val in enumerate(merged_dataframe.columns, start=0):
    #    target_sheet.write(row_i, col, val)
    row_i += 1

    for index, row_content in merged_dataframe.iterrows():
        for col, val in enumerate(row_content, start=0):
            if val is np.nan:
                target_sheet.write(row_i, col)
            else:
                target_sheet.write(row_i, col, val)
        row_i += 1
    target_sheet.write(row_i, 0,"endHISsheet")

    if write_end > row_i:
        for row_i in range(row_i, write_end):
            target_sheet.write(row_i, 0)

    target_wb.save(target_file)

    #
    # target_sheet = open_workbook(target_file).sheets()[0]
    # row_i = 0
    # precols = []
    # for row_i in range(0, write_start):
    #     precols.append(target_sheet.row_values(row_i))
    #
    #
    # target_wb = xlsxwriter.Workbook(target_file)
    # target_sheet = target_wb.add_worksheet("First Sheet")
    #
    # row_i = write_start
    # target_sheet.write_row(row_i, 0, merged_dataframe.columns)
    # row_i +=  1
    # for index, row_content in merged_dataframe.iterrows():
    #     target_sheet.write_row(row_i, 0, row_content)
    #     row_i += 1
    # target_sheet.write(row_i, 0, "endHISsheet")
    #
    # if write_end > row_i:
    #     for row_i in range(row_i, write_end):
    #         target_sheet.write(row_i, 0, None)
    #
    # target_wb.close()

    pass


