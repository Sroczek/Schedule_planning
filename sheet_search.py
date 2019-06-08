from openpyxl import load_workbook
from collections import defaultdict


class SheetSearch:
    def __rows_generator(self, ws):
        for i, row in enumerate(ws.iter_rows(min_row=2, max_col=16)):
            yield [i + 2] + [cell.value for cell in row]

    def __get_titles(self, ws):
        for row in ws.iter_rows(min_row=1, max_row=1, max_col=16):
            return ["numer"] + [cell.value for cell in row]

    def __map_data(self, ws):
        titles = self.__get_titles(ws)
        for row in self.__rows_generator(ws):
            yield {titles[i]: cell for i, cell in enumerate(row)}

    def __map_rooms(self, ws):
        gen = self.__map_data(ws)
        for row in gen:
            yield {"nazwa": ((row["bud"] if row["bud"] is not None else "") + " " + (
                row["sala"] if row["sala"] is not None else "")).strip(),
                   "typ": row["typ"] if row["typ"] is not None else ""}

    def get_blocking_rows(self, ws, row_no):
        """Metoda, ktora zwraca wszystkie terminy, w ktorych prowadzacy lub dany rocznik jest zajety"""
        magic_row = list(filter(lambda x: x["numer"] == row_no, list(self.rdd_dict[ws])))[0]
        return list(
            set(map(lambda x: (x["dzien"], x["godz"], x["koniec"]), filter(lambda x: x["osoba"] == magic_row["osoba"] or
                                                                                     (x["studia"] == magic_row[
                                                                                         "studia"] and x["sem"] ==
                                                                                      magic_row["sem"]),
                                                                           self.rdd_dict[ws]))))

    def get_proper_rooms(self, ws, row_no):
        """Metoda, ktora zwraca nazwy wszystkich sal, w ktorych moga odbyc sie dane zajecia"""
        magic_row = list(filter(lambda x: x["numer"] == row_no, self.rdd_dict[ws]))[0]
        return list(
            set(map(lambda x: x["nazwa"], filter(lambda x: x["typ"] == magic_row["typ"], self.rdd_dict["sale"]))))

    def get_blocking_rows_for_rooms(self, ws, room_list):
        """Metoda, ktora zwraca slownik przypisujacy kazdej sali trojki: dzien,godzina,koniec,
         w ktorych odbywaja sie zajecia"""
        help = list(set(map(lambda x: (x["sala"], x["dzien"], x["godz"], x["koniec"]),
                            filter(lambda x: x["sala"] in room_list, self.rdd_dict[ws]))))
        result = defaultdict(list)
        for i, j, k, l in help:
            result[i].append((j, k, l))
        return result

    def __init__(self, path):
        wb = load_workbook(filename=path, read_only=True)
        self.rdd_dict = {
            "zima-s": [x for x in self.__map_data(wb["zima-s"])],
            "lato-s": [x for x in self.__map_data(wb["lato-s"])],
            "zima-n": [x for x in self.__map_data(wb["zima-n"])],
            "lato-n": [x for x in self.__map_data(wb["lato-n"])],
            "sale": [x for x in self.__map_rooms(wb["sale"])]
        }


# Przyklad uzycia


ss = SheetSearch("sheet.xlsx")
pr = ss.get_proper_rooms("zima-s", 109)
brr = ss.get_blocking_rows_for_rooms("zima-s", pr)
br = ss.get_blocking_rows("zima-s", 109)
print("Zajete terminy:")
for item in br:
    print(item)
print('\n')
print("Zajetosc sal:")
for key, value in brr.items():
    print(key)
    for v in value:
        print(v)
