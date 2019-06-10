import argparse
from sheet_search import SheetSearch
import intervals as I
import datetime

# parser = argparse.ArgumentParser(description="Program służy do wyszukiwania terminów na które można przełożyć dane zajęcia\n")
#
# parser.add_argument('-p', '--page', help="Nazwa zakładki arkusza", required=True)
# parser.add_argument('-r', '--row', help="Wiersz dla którego zamiany wyszukać", type=int, required=True)
#
#
# argv = parser.parse_args()
# row = argv.row
# semestr = argv.page

begin_classes_time = datetime.time(hour=8, minute=0)
end_classes_time = datetime.time(hour=19, minute=20)
days_of_week = ('Pn', 'Wt', 'Sr', 'Cz', 'Pt', 'Sb', 'Nd')
class_duration = 90


def parse_query(query):
    argv = query.split(" ")
    return argv[0], int(argv[1])


def dtoi(d: datetime.time):
    return d.hour*60 + d.minute


def itod(i: int):
    return datetime.time(hour=i//60, minute=i%60)


def blocking_row_to_interval(br):
    if br[0] not in days_of_week or br[1] is None: return []
    if br[2] is None: return [br[0], I.closedopen(dtoi(br[1]), dtoi(br[1]) + 90)]
    return [br[0], I.closedopen(dtoi(br[1]), dtoi(br[2]))]


def execute(ss, query):
    ws, row = parse_query(query)
    br = ss.get_blocking_rows(ws, row)
    br = list(filter(lambda x: x != [], map(lambda x: blocking_row_to_interval(x), br)))

    days_unreserved_hours = {day: I.closed(dtoi(begin_classes_time), dtoi(end_classes_time)) for day in days_of_week}

    for interval in br:
        days_unreserved_hours[interval[0]] = days_unreserved_hours[interval[0]] - interval[1]

    pr = ss.get_proper_rooms(ws, row)
    brr = ss.get_blocking_rows_for_rooms(ws, pr)

    result = []

    for room, intervals_when_blocked in brr.items():
        days_unreserved_room_hours = {day: I.closed(dtoi(begin_classes_time), dtoi(end_classes_time)) for day in days_of_week}

        intervals_when_blocked = list(filter(lambda x: x != [], map(lambda x: blocking_row_to_interval(x), intervals_when_blocked)))
        for interval in intervals_when_blocked:
            days_unreserved_room_hours[interval[0]] = days_unreserved_room_hours[interval[0]] - interval[1]

        for day in days_of_week:
            possible_hours = days_unreserved_hours[day] & days_unreserved_room_hours[day]
            result += [(room, day, interv) for interv in list(possible_hours)]

    result = list(map(lambda x: "{} {} {} {}".format(x[0], x[1], itod(x[2].lower), itod(x[2].upper)), filter(lambda y: y[2].upper - y[2].lower >= class_duration, result)))
    return result


ss = SheetSearch("sheet.xlsx")

query = input(">>")
while query != "q":
    for el in execute(ss, query):
        print(el)
    query = input(">>")


