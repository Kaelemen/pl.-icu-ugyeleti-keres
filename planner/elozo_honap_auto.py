# -*- coding: utf-8 -*-
"""Beolvassa a megosztott Google Drive mappa legutóbb módosított Excel-fájlját (az
előző hónap véglegesített beosztása), és megkeresi, ki volt ügyeletben annak utolsó
napján - ők lépnek le az új hónap 1-jén. Az eredményt egy JSON fájlba írja, amit a
fő generáló script beolvas és beépít a "elozo_honap_lelepok" mezőbe.

Csak OLVASÁSI jogosultságot igényel (drive.readonly) - a szolgáltatásfiók csak a
vele megosztott mappát látja, semmi mást a Drive-on."""
import os
import io
import json
import sys
import datetime
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "").strip()
SA_JSON = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON", "").strip()
KIMENET = "elozo_honap_auto.json"

DUTY_KODOK = {"I", "A", "St"}


def van_ugyeletkod(cellertek):
    if not cellertek:
        return False
    return any(resz in DUTY_KODOK for resz in str(cellertek).split("/"))


def main():
    if not FOLDER_ID or not SA_JSON:
        print("Nincs beállítva GDRIVE_FOLDER_ID vagy GDRIVE_SERVICE_ACCOUNT_JSON - kihagyva.")
        json.dump({"elozo_honap_lelepok": []}, open(KIMENET, "w", encoding="utf-8"))
        return

    try:
        info = json.loads(SA_JSON)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        service = build("drive", "v3", credentials=creds)

        results = service.files().list(
            q=f"'{FOLDER_ID}' in parents and trashed = false",
            orderBy="modifiedTime desc",
            pageSize=10,
            fields="files(id, name, modifiedTime, mimeType)",
        ).execute()
        fajlok = results.get("files", [])
        # csak a valódi Excel (xlsx) fájlokat nézzük, a legutóbb módosítottat használjuk
        xlsx_fajlok = [f for f in fajlok if f["name"].lower().endswith(".xlsx")]
        if not xlsx_fajlok:
            print("Nem található .xlsx fájl a megosztott mappában - kihagyva.")
            json.dump({"elozo_honap_lelepok": []}, open(KIMENET, "w", encoding="utf-8"))
            return

        legutobbi = xlsx_fajlok[0]
        print(f"Legutóbbi fájl a Drive-mappában: {legutobbi['name']} ({legutobbi['modifiedTime']})")

        request = service.files().get_media(fileId=legutobbi["id"])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)

        wb = openpyxl.load_workbook(buf, data_only=True)
        lap_nev = None
        for jelolt in wb.sheetnames:
            if "Nyomtatási" in jelolt or "Változat" in jelolt:
                lap_nev = jelolt
                break
        if lap_nev is None:
            lap_nev = wb.sheetnames[0]
        ws = wb[lap_nev]
        print(f"Használt munkalap: {lap_nev}")

        # az utolsó napot tartalmazó oszlop megkeresése: a legmagasabb oszlopszámú
        # cella, aminek van tartalma bármelyik dolgozó sorában (2. oszloptól indul a
        # napok rácsa a mi formátumunkban)
        max_col = 2
        for row in ws.iter_rows(min_row=5, max_row=30, min_col=2, max_col=40):
            for cell in row:
                if cell.value not in (None, "") and cell.column > max_col:
                    max_col = cell.column

        lelepok = []
        for row in ws.iter_rows(min_row=5, max_row=30):
            nev = row[0].value
            if not nev:
                continue
            utolso_nap_cella = None
            for cell in row:
                if cell.column == max_col:
                    utolso_nap_cella = cell
                    break
            if utolso_nap_cella is not None and van_ugyeletkod(utolso_nap_cella.value):
                lelepok.append(nev)

        print(f"Az utolsó napon ({max_col - 1}.) ügyeletben lévők (ők lépnek le): {lelepok}")
        json.dump({"elozo_honap_lelepok": lelepok}, open(KIMENET, "w", encoding="utf-8"), ensure_ascii=False)

    except Exception as e:
        print(f"Hiba a Drive-olvasás közben ({e}) - üres listával folytatjuk, az admin felület kézi bevitele lesz a mérvadó.")
        json.dump({"elozo_honap_lelepok": []}, open(KIMENET, "w", encoding="utf-8"))


if __name__ == "__main__":
    main()
