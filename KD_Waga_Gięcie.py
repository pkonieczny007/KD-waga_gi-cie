import pandas as pd
import glob

# Znajdź pliki Operacje* i Produkty*
operacje_files = glob.glob("Operacje*.xlsx")
produkty_files = glob.glob("Produkty*.xlsx")

if not operacje_files or not produkty_files:
    raise FileNotFoundError("Nie znaleziono plików Operacje* lub Produkty* w katalogu.")

# Załaduj pliki (zakładamy, że pierwszy plik znaleziony jest używany)
operacje_file = operacje_files[0]
produkty_file = produkty_files[0]

operacje_df = pd.read_excel(operacje_file)
produkty_df = pd.read_excel(produkty_file)

# Upewnij się, że wymagane kolumny istnieją
if 'Produkt' not in operacje_df.columns or 'Referencja' not in produkty_df.columns or 'Ciężar' not in produkty_df.columns:
    raise ValueError("Brakuje wymaganych kolumn w jednym z plików.")

# Dopasowanie ciężaru
def znajdz_ciezar(produkt):
    # Znajdź odpowiednią referencję w Produkty*
    match = produkty_df[produkty_df['Referencja'] == produkt]
    if not match.empty:
        return match['Ciężar'].values[0]
    return "brak"

# Dodaj nową kolumnę z ciężarem do pliku Operacje*
operacje_df['Nazwa produktu'] = operacje_df['Produkt'].apply(znajdz_ciezar)

# Zapisz wynikowy plik
wynik_file = "wynik.xlsx"
operacje_df.to_excel(wynik_file, index=False)
print(f"Zapisano wynik do pliku {wynik_file}")
