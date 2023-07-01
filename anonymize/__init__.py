from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment
import os
from collections.abc import Iterator
from uuid import uuid5, NAMESPACE_DNS
from dataclasses import dataclass
from shutil import copy2
import datetime


@dataclass
class Patient:
    name: str
    surname: str
    birthDate: str


ROOT_DIR = os.path.abspath(os.curdir)
ORIGINAL_FOLDER_PATH = os.path.join(ROOT_DIR, "tests/original")

ANON_FOLDER_PATH = os.path.join(ROOT_DIR, "tests/anonymized")
ANON_ENTRIES_PATH = os.path.join(ANON_FOLDER_PATH, "anon_entries.xlsx")


def get_anon_entries_wb():
    if not os.path.exists(ANON_ENTRIES_PATH):
        wb = Workbook()
        ws = wb.active
        if ws is not None:
            ws.title = "Entries"
            wss = wb.get_sheet_by_name("Entries")
            wss.append(["Nome", "Cognome", "Data di nascita", "ID"])
            for c in ["A1", "B1", "C1", "D1"]:
                wss[c].font = Font(size=20, bold=True)
                wss[c].alignment = Alignment(horizontal="center")
            wb.save(ANON_ENTRIES_PATH)
            return wb
    else:
        return load_workbook(ANON_ENTRIES_PATH)


def get_dir(path: str):
    return list(os.scandir(path))


def get_patients(dir: list[os.DirEntry]) -> list[Patient]:
    # For each file: name-surname-DD_MM__YYYY
    patients = []
    for f in dir:
        noExt = str(f.name).split(".")[0]
        [name, surname, birthDate] = noExt.strip().split("-")
        patients.append(Patient(name, surname, birthDate))
    return patients


def generate_UUID5_patient(patient: Patient):
    token = "".join([patient.name, patient.surname, patient.birthDate])
    return "".join(str(uuid5(NAMESPACE_DNS, token)).split("-"))[:10]


def save_anon_patient_file(file: os.DirEntry, patientID: str):
    anonFileName = "".join(
        [patientID, "_", str(datetime.datetime.today().year), ".xlsx"]
    )
    destinationPath = os.path.join(ANON_FOLDER_PATH, anonFileName)
    copy2(file.path, destinationPath)


def is_ID_in_ws(ws: Worksheet, ID: str):
    for i, row in enumerate(ws.values):
        if row[3] == ID:
            print("ID: " + ID + " found existing in row " + str(i))
            return True
    return False


def create_anon_entry(wb: Workbook, ws: Worksheet, patient: Patient, ID: str):
    if is_ID_in_ws(ws, ID):
        return

    ws.append([patient.name, patient.surname, patient.birthDate, ID])


def anonymize():
    os.makedirs(ANON_FOLDER_PATH, exist_ok=True)
    if not os.path.exists(ORIGINAL_FOLDER_PATH):
        print(
            "[WARNING] Original files folder was not found, one will be created for you in tests/original. This execution will be exited"
        )
        os.makedirs(ORIGINAL_FOLDER_PATH, exist_ok=True)
        return
    anonWb = get_anon_entries_wb()
    anonWs = None
    if anonWb is not None:
        anonWs = anonWb.get_sheet_by_name("Entries")
    originalDir = get_dir(ORIGINAL_FOLDER_PATH)
    patients = get_patients(originalDir)
    print("Found patients: " + str(patients))
    for i, f in enumerate(originalDir):
        id = generate_UUID5_patient(patients[i])
        save_anon_patient_file(f, id)
        if anonWb is not None and anonWs is not None:
            create_anon_entry(anonWb, anonWs, patients[i], id)
        else:
            print("[ERROR] Workbook for entries is None")

    if anonWb is not None:
        anonWb.save(ANON_ENTRIES_PATH)
