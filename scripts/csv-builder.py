import os
import pandas as pd
from openpyxl import load_workbook
import warnings

# Ez a script a jelenléti ívekből kiszedi a túlóra adatokat és készít belőlük nevenként egy csv fájlt.

# Elnyomjuk az adatérvényesítés miatti figyelmeztetéseket
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Paraméterek
mappa = "I:\\#Jelenléti ív 2025"
start = "2025-12-01 00:00:00"
stop = "2025-12-31 00:00:00"
filename = "december"
start_time = pd.to_datetime(start)
stop_time = pd.to_datetime(stop)

# Ünnepnapok
unnepnapok = pd.to_datetime([
    "2025-01-01","2025-03-15","2025-04-18","2025-04-21","2025-05-01","2025-06-09","2025-08-20","2025-10-23","2025-11-01","2025-12-25","2025-12-26"
])

# Áthelyezett munkanapok
moved_days = pd.to_datetime(["2025-05-17","2025-10-18","2025-12-13"])

# Timedelta objektumot órákra konvertál
def to_hours(td):
    if pd.isna(td):
        return 0.0
    return td.total_seconds() / 3600

# Órában megadott számot (float) átalakít HH:MM:SS formátumú stringgé
def timedelta_to_hhmmss(hours):
    td = pd.to_timedelta(hours, unit="h")
    total_seconds = int(td.total_seconds())
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02}:{m:02}:{s:02}"

# Timedelta vagy float órát float órára konvertál
def timedelta_to_float_hours(value):
    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 3600
    return value or 0.0  # NaN vagy None esetén 0.0

for file in os.listdir(mappa):
    if file.endswith(".xlsx"):
        path = os.path.join(mappa, file)

        try:
            df = pd.read_excel(path, usecols=[0,1,5,6,7])
            df["Munka megkezdésének időpontja"] = pd.to_datetime(
                df["Munka megkezdésének időpontja"], errors="coerce"
            )

            # Szűrések
            weekdays = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (~df["Munka megkezdésének időpontja"].isin(unnepnapok)) &
                ((df["Munka megkezdésének időpontja"].dt.weekday < 5) |
                 (df["Munka megkezdésének időpontja"].isin(moved_days)))
            ].copy()

            weekends = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (~df["Munka megkezdésének időpontja"].isin(unnepnapok)) &
                (~df["Munka megkezdésének időpontja"].isin(moved_days)) &
                (df["Munka megkezdésének időpontja"].dt.weekday >= 5)
            ].copy()

            holidays = df[
                (df["Munka megkezdésének időpontja"] >= start_time) &
                (df["Munka megkezdésének időpontja"] <= stop_time) &
                (df["Távollét oka"] == "Rendkívüli munkavégzés") &
                (df["Munka megkezdésének időpontja"].dt.normalize().isin(unnepnapok))
            ].copy()

            # Időtartamok órában (float)
            weekdays["Óra (dec)"] = pd.to_timedelta(
                weekdays["Kiadandó/Megváltandó rendkívüli munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            weekends["Óra (dec)"] = pd.to_timedelta(
                weekends["Napi munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            holidays["Óra (dec)"] = pd.to_timedelta(
                holidays["Napi munkaidő"].astype(str),
                errors="coerce"
            ).apply(to_hours)

            # Összegzés (float órában)
            weekdays_sum = timedelta_to_float_hours(weekdays["Óra (dec)"].sum())
            weekends_sum = timedelta_to_float_hours(weekends["Óra (dec)"].sum())
            holidays_sum = timedelta_to_float_hours(holidays["Óra (dec)"].sum())

            # Összegző DataFrame
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

            # Nevek kinyerése
            wb = load_workbook(path, data_only=True)  # data_only=True: csak a cella értéke
            sheet = wb.active  # az első munkalap
            nev = sheet["B367"].value
            name_col = pd.DataFrame({ "Dolgozó": [nev] + [""] * len(summary) })
            summary_with_name = pd.concat([name_col.reset_index(drop=True), summary.reset_index(drop=True)], axis=1)

            # Mentés CSV-be
            output_name = os.path.splitext(file)[0] + "_" + filename + "_eredmeny.csv"
            output_path = os.path.join("../data", output_name)
            summary_with_name.to_csv(output_path, index=False, sep=";", encoding="utf-8-sig")

        except Exception as e:
            print(f"Hiba a(z) {file} feldolgozásakor: {e}")
