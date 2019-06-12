from sheet_search import SheetSearch, itod

ss = SheetSearch("sheet.xlsx")

query = input(">>")
while query != "q":

    argv = query.split(" ")
    if argv[0] not in ['zima_s', 'zima_n', 'lato_s', 'lato_n']:
        print("Pierwszy argument powinien być nazwą zakładki\n[zima_s | zima_n | lato_s | lato_n]")
    elif argv.__len__() == 1:
        print("Oczekiwano dwóch argumnetów: nazwa zeszytu, numer wiersza")
    else:
        try:
            res = ss.find_possible_hours(argv[0], int(argv[1]))
        except KeyError:
             print("Nazwy kolumn uległy zmianie. Konieczna jest modyfikacja oprogramowania")
        except ValueError:
            print("Drugi argument powiniem być numerem wiersza")
        except IndexError:
            print("Wiersz o podanym numerze nie istnieje")
        else:
            for x in res:
                print("{} {}    {} - {}   {}".format(x[0], x[1], itod(x[2].lower), itod(x[2].upper), x[2].upper - x[2].lower))

    query = input(">>")

