import pandas as pd
import matplotlib.pyplot as plt
import glob
from datetime import datetime


#------------------------------------------------------------------------------------------------------


##WYTYCZNE IMPORTOWANIA & SORTOWANIA TABELI
nazwa = 'Raport'
kolumny = {'Product_ID', 'Sale_Date', 'Sales_Amount'}
arkusz = 'Sheet1'
pliki_do_pominiecia = 'gotowa_tabela.xlsx'

#kolumna po której będzie sortowana tabela malejąco
kolumna_sort = 'Sale_Date'


#------------------------------------------------------------------------------------------------------


#IMPORTOWANIE TABEL 
def import_excel():
    pliki = glob.glob("*.xlsx")
    tabela = []

    if not pliki:
        print(f"Nie udało się znaleść żadnego pliku")
        return 

    print(f"Znaleziono pliki do zainportowania")


    for nazwa in pliki:

        if nazwa == pliki_do_pominiecia:
            print (f"Pominięto plik: {nazwa}")
            continue
         
        try:
            temp_tabela = pd.read_excel(nazwa, sheet_name = arkusz)
            

            #weryfikacja kolumn
            if not kolumny.issubset(temp_tabela.columns):
                raise ValueError (f"BŁĄD: Plik {nazwa} ma niepoprawne kolumny !")
            
            temp_tabela = temp_tabela[list(kolumny)]
        
            tabela.append(temp_tabela)
            print(f"Zainportowano poprawnie {nazwa}")
    

        except Exception as e:
            print (f"Problem z plikiem {nazwa}: {e}")
            raise

       
    print("Zakończono importowanie sukcesem")

    tabela =  pd.concat(tabela, ignore_index=True)
    tabela_podglad = tabela.sort_values(by=kolumna_sort, ascending = False)

    print(tabela_podglad)
    return tabela_podglad



#------------------------------------------------------------------------------------------------------


#Eksportowanie wykonaniej tabeli
def export_excel(dane_do_zapisu):
    if dane_do_zapisu is not None and not dane_do_zapisu.empty:
        
        dzisiaj = datetime.now().strftime("%Y-%m-%d")
        nazwa_tabeli = f"{dzisiaj} {nazwa}.xlsx"

        dane_do_zapisu.to_excel(nazwa_tabeli, index=False)
        print(f"Plik został wyeksportowany jako: {nazwa_tabeli}")
        return dane_do_zapisu
    
    else:
        print("Błąd: Brak danych do wyeksportowania (tabela jest pusta).")

if __name__ == "__main__":
    # NAJPIERW musimy pobrać dane z pierwszej funkcji
    tabela_wynikowa = import_excel() 
    
    # POTEM przekazujemy je do eksportu
    export_excel(tabela_wynikowa)        
