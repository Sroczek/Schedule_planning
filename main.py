from openpyxl.utils.exceptions import InvalidFileException
from sheet_search import SheetSearch, itod, cm


print("Schedule Planning V1.0\nAplikacja ułatwiająca przenoszenie zajęć\n")
sheet_name = input("Podaj ścieżkę do arkusza:\n>>")
if not sheet_name:
    sheet_name = "sheet.xlsx"
try:
    ss = SheetSearch(sheet_name)
except (InvalidFileException,FileNotFoundError):
    print("Ścieżka do pliku niepoprawna")
    exit(0)
print("Arkusz został załadowany. Podaj wiersz w formacie: <nazwa_zakladki> <nr_wiersza>")
query = input(">>")
while query != "q":

    argv = query.split(" ")
    if argv[0] == 'find':
        if argv[1] not in [cm['zima_s'], cm['zima_n'], cm['lato_s'], cm['lato_n']]:
            print("Pierwszy argument powinien być nazwą zakładki\n[%s | %s | %s | %s]" % (cm['zima_s'], cm['zima_n'], cm['lato_s'], cm['lato_n']))
        elif argv.__len__() < 3:
            print("Oczekiwano dwóch argumentów: nazwa zeszytu, numer wiersza")
        else:
            try:
                res = ss.find_possible_hours(argv[1], int(argv[2]))
            except KeyError as k:
                print(k)
                print("Nazwy kolumn uległy zmianie. Konieczna jest modyfikacja oprogramowania")
            except ValueError:
                print("Drugi argument powiniem być numerem wiersza")
            except IndexError:
                print("Wiersz o podanym numerze nie istnieje")
            else:
                for x in res:
                    print("{} {}    {} - {}   {}".format(x[0], x[1], itod(x[2].lower), itod(x[2].upper), x[2].upper - x[2].lower))
    else:
        print("Obslugiwane komendy: find")
    query = input(">>")

