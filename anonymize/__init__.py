import os
from uuid import uuid4
from shutil import copy2
import datetime
import pandas as pd


ROOT_DIR = os.path.abspath(os.curdir)
ORIGINAL_FOLDER_PATH = os.path.join(ROOT_DIR, "tests/original")

ANON_FOLDER_PATH = os.path.join(ROOT_DIR, "tests/anonymized")
ANON_ENTRIES_PATH = os.path.join(ANON_FOLDER_PATH, "anon_entries.xlsx")


def create_entries_excel():
    if not os.path.exists(ANON_ENTRIES_PATH):
        df = pd.DataFrame(data={"Nome completo": [], "ID": []})
        df.to_excel(ANON_ENTRIES_PATH, sheet_name="Entries", index=False)


def get_dir(path: str):
    return list(os.scandir(path))


def find_patient_name(file: pd.ExcelFile):
    sheets = file.sheet_names
    name = ""
    for sh in sheets:
        name = ""
        df = pd.read_excel(file, sheet_name=sh)
        r = df.columns.to_numpy().tolist()

        startConcat = False
        for cell in r:
            if cell == "Nome:":
                if startConcat == False:
                    startConcat = True
                else:
                    print("Weird error in finding name")
                    name = None
                    break
            elif startConcat == True:
                if str(cell).find(":") != -1:
                    if len(name.strip()) > 0:
                        return name.strip()
                    name = None
                    break
                if cell != None:
                    name += " " + str(cell)
    return name


def generate_UUID4():
    return "".join(str(uuid4()).split("-"))[:10]


def check_for_omonimy(patient_id: str):
    anon_file_name = "".join(
        [patient_id, "_", str(datetime.datetime.today().year), ".xlsx"]
    )
    destinationPath = os.path.join(ANON_FOLDER_PATH, anon_file_name)
    if os.path.exists(destinationPath):
        print(
            "[WARNING]: Possible omonimy detected for patient with id: "
            + patient_id
            + " File will not be saved!"
        )
        return True
    return False


def save_anon_patient_file(file: os.DirEntry, patient_id: str):
    anon_file_name = "".join(
        [patient_id, "_", str(datetime.datetime.today().year), ".xlsx"]
    )
    destinationPath = os.path.join(ANON_FOLDER_PATH, anon_file_name)
    copy2(file.path, destinationPath)


def get_patient_id(patient_full_name: str):
    df = pd.read_excel(ANON_ENTRIES_PATH)
    rows = df.to_numpy().tolist()

    for i, row in enumerate(rows):
        if row[0] == patient_full_name:
            print(
                "Found id: "
                + str(row[1])
                + " for patient "
                + str(patient_full_name)
                + " in row: "
                + str(i)
            )
            return str(row[1])
    return None


def is_file_valid(file: os.DirEntry):
    if str(file.name).find(".xlxs") == -1 and str(file.name).find(".xls") == -1:
        return False
    return True


def anonymize():
    os.makedirs(ANON_FOLDER_PATH, exist_ok=True)
    if not os.path.exists(ORIGINAL_FOLDER_PATH):
        print(
            "[WARNING] Original files folder was not found, one will be created for you in tests/original. This execution will be exited"
        )
        os.makedirs(ORIGINAL_FOLDER_PATH, exist_ok=True)
        return

    original_dir = get_dir(ORIGINAL_FOLDER_PATH)
    valid_files = list(filter(is_file_valid, original_dir))

    create_entries_excel()
    with pd.ExcelWriter(
        ANON_ENTRIES_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay"
    ) as writer:
        for f in valid_files:
            patient_full_name = find_patient_name(pd.ExcelFile(f.path))

            if patient_full_name is None:
                print(
                    "[ERROR]: File "
                    + f.name
                    + " will be skipped because it does not contain the name of a patient."
                )
                continue

            patient_id = get_patient_id(patient_full_name)

            if patient_id is None:
                patient_id = generate_UUID4()
                df = pd.DataFrame(
                    [[patient_full_name, patient_id]],
                    columns=["Nome completo", "ID"],
                )

                startrow = writer.sheets["Entries"].max_row
                df.to_excel(
                    writer,
                    startrow=startrow,
                    header=False,
                    index=False,
                    sheet_name="Entries",
                )
            if not check_for_omonimy(patient_id):
                save_anon_patient_file(f, patient_id)
