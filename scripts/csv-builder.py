import os
import pandas as pd
from openpyxl import load_workbook
import warnings

# Ez a script a jelenléti ívekből kiszedi a túlóra adatokat és készít belőlük nevenként egy csv fájlt.

# Elnyomjuk az adatérvényesítés miatti figyelmeztetéseket
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Paraméterek
mappa = "I:\\#Jelenléti ív 2026"
start = "2026-01-01 00:00:00"
stop = "2026-01-31 23:59:59"
filename = "január"
start_time = pd.to_datetime(start)
stop_time = pd.to_datetime(stop)

# Ünnepnapok (beleértve az áthelyezett pihenőnapokat is)
unnepnapok = pd.to_datetime([
    "2026-01-01","2026-01-02","2026-03-15","2026-04-03","2026-04-06","2026-05-01",
    "2026-05-25","2026-08-20","2026-08-21","2026-10-23","2026-11-01","2026-12-24",
    "2026-12-25","2026-12-26"
])

# Áthelyezett munkanapok (szombatok, amik hétköznapnak számítanak)
moved_days = pd.to_datetime(["2026-01-10","2026-08-08","2026-12-12"])

# Segédfunkciók
def to_hours(td):
    if pd.isna(td):
        return 0.0
    return td.total_seconds() / 3600

def timedelta_to_hhmmss(hours):
    td = pd.to_timedelta(hours, unit="h")
    total_seconds = int(td.total_seconds())
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02}:{m:02}:{s:02}"

def timedelta_to_float_hours(value):
    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 3600
    return value or 0.0

for file in os.listdir(mappa):
    if file.endswith(".xlsx"):
        path = os.path.join(mappa, file)

        try:
            # Csak a szükséges oszlopok beolvasása
            df = pd.read_excel(path, usecols=[0,1,5,6,7])
            df["Munka megkezdésének időpontja"] = pd.to_datetime(
                df["Munka megkezdésének időpontja"], errors="coerce"
            )

            # --- SZŰRÉSEK (A .dt.normalize() biztosítja a pontos egyezést a listákkal) ---

            # 1. Hétköznap: Nem ünnep ÉS (Hétköznap VAGY ledolgozós szombat)
            weekdays = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (~df["Munka megkezdésének időpontja"].dt.normalize().isin(unnepnapok)) &
                ((df["Munka megkezdésének időpontja"].dt.weekday < 5) |
                 (df["Munka megkezdésének időpontja"].dt.normalize().isin(moved_days)))
            ].copy()

            # 2. Pihenőnap: Nem ünnep ÉS Hétvége ÉS Nem ledolgozós szombat
            weekends = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (~df["Munka megkezdésének időpontja"].dt.normalize().isin(unnepnapok)) &
                (df["Munka megkezdésének időpontja"].dt.weekday >= 5) &
                (~df["Munka megkezdésének időpontja"].dt.normalize().isin(moved_days))
            ].copy()

            # 3. Munkaszüneti nap: Benne van az ünnepnapok (és áthelyezett pihenőnapok) listájában
            holidays = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (df["Munka megkezdésének időpontja"].dt.normalize().isin(unnepnapok))
            ].copy()

            # --- SZÁMOLÁS A MEGFELELŐ OSZLOPOKBÓL ---

            # Hétköznap -> "Kiadandó/Megváltandó..." (1. oszlop)
            weekdays["Óra (dec)"] = pd.to_timedelta(
                weekdays["Kiadandó/Megváltandó rendkívüli munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            # Pihenőnap -> "Napi munkaidő" (2. oszlop)
            weekends["Óra (dec)"] = pd.to_timedelta(
                weekends["Napi munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            # Munkaszüneti nap -> "Napi munkaidő" (2. oszlop)
            holidays["Óra (dec)"] = pd.to_timedelta(
                holidays["Napi munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            # Összegzés
            weekdays_sum = timedelta_to_float_hours(weekdays["Óra (dec)"].sum())
            weekends_sum = timedelta_to_float_hours(weekends["Óra (dec)"].sum())
            holidays_sum = timedelta_to_float_hours(holidays["Óra (dec)"].sum())

            # Eredmény táblázat
            summary = pd.DataFrame([
                {"Kategória": "Hétköznap összesen", 
                 "Időtartam (óra)": f"{weekdays_sum:.2f}", 
                 "Időtartam (hh:mm:ss)": timedelta_to_hhmmss(weekdays_sum)},
                {"Kategória": "Pihenőnap összesen", 
                 "Időtartam (óra)": f"{weekends_sum:.2f}", 
                 "Időtartam (hh:mm:ss)": timedelta_to_hhmmss(weekends_sum)},
                {"Kategória": "Munkaszüneti nap összesen", 
                 "Időtartam (óra)": f"{holidays_sum:.2f}", 
                 "Időtartam (hh:mm:ss)": timedelta_to_hhmmss(holidays_sum)},
            ])

            # Név beolvasása B367 cellából
            wb = load_workbook(path, data_only=True)
            sheet = wb.active
            nev = sheet["B367"].value
            name_col = pd.DataFrame({ "Dolgozó": [nev] + [""] * (len(summary)-1) })
            summary_with_name = pd.concat([name_col, summary], axis=1)

            # Mentés
            output_name = os.path.splitext(file)[0] + "_" + filename + "_eredmeny.csv"
            output_path = os.path.join("../data", output_name)
            summary_with_name.to_csv(output_path, index=False, sep=";", encoding="utf-8-sig")

        except Exception as e:
            print(f"Hiba a(z) {file} feldolgozásakor: {e}")
