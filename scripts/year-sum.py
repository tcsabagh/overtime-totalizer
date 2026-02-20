import pandas as pd
import glob
import os

# Ez a script összesíti az xlsx-eket és kiszámolja az éves összes túlórát.

# 1. Beállítások
path = r'../data'
fajlok = glob.glob(os.path.join(path, "*.xlsx*"))

if not fajlok:
    print("Nem található Excel fájl a megadott mappában!")
else:
    # 2. Beolvasás és összefűzés
    lista = []
    for fajl in fajlok:
        df = pd.read_excel(fajl)
        lista.append(df)
    
    df_mind = pd.concat(lista, ignore_index=True)

    # 3. Soronkénti összesítés (A 3 kategória összege minden sorban)
    oszlopok = ['Hétköznap összesen', 'Pihenőnap összesen', 'Munkaszüneti nap összesen']
    # A .sum(axis=1) vízszintesen adja össze az értékeket
    df_mind['Összes túlóra'] = df_mind[oszlopok].sum(axis=1)

    # 4. Nevenkénti csoportosítás (Aggregálás)
    # Itt egyszerre adjuk össze a részletezett oszlopokat és a már kiszámolt összesen oszlopot
    final_data = df_mind.groupby('Név').agg({
        'Hétköznap összesen': 'sum',
        'Pihenőnap összesen': 'sum',
        'Munkaszüneti nap összesen': 'sum',
        'Összes túlóra': 'sum'
    }).reset_index()

    # 5. Végeredmény mentése
    output_name = '../eves-osszesito-vegleges.xlsx'
    final_data.to_excel(output_name, index=False)

    print(f"Siker! {len(fajlok)} fájl feldolgozva.")
    print(f"Az összesített táblázat mentve: {output_name}")
    
    # Ellenőrzés: írjuk ki az első 5 sort a terminálba is
    print("\nÍzelítő a végeredményből:")
    print(final_data.head())
