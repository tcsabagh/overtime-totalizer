import os
import pandas as pd

# Ez a script összesíti az adott hónap adatait csv-ből.

mappa = "../data"
month = "december"
output_excel = f"../data/sum-{month}.xlsx"

# Oszlopok meghatározása
df = pd.DataFrame(columns=["Név", "Hónap", "Hétköznap összesen", "Pihenőnap összesen", "Munkaszüneti nap összesen"])

try:
    for file in os.listdir(mappa):
        if file.endswith(".csv"):
            csv_path = os.path.join(mappa, file)
            df_csv = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig")

            name = df_csv.iloc[0,0]
            weekday = df_csv.iloc[0,2]
            weekend = df_csv.iloc[1,2]
            holiday = df_csv.iloc[2,2]

            df.loc[len(df)] = [name, month, weekday, weekend, holiday]

    # Mentés Excelbe
    df.to_excel(output_excel, index=False)

except Exception as e:
    print(f"Hiba a(z) {file} feldolgozásakor: {e}")
