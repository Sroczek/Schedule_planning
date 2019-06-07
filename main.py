from openpyxl import load_workbook
from pandas import DataFrame
wb = load_workbook(filename='2019-2020.xlsx', read_only=True)
ws = wb['egz']

data = ws.values
# Get the first line in file as a header line
columns = next(data)[0:]
# Create a DataFrame based on the second and subsequent lines of data
df = DataFrame(data, columns=columns)

print(df.head(100000))
df.select()


#
# # wb.get_sheet_by_name('egz')
# # print((wb.index(ws)))
#
# df = DataFrame(ws.values)
# data = ws.values
# cols = next(data)[1:]
# data = list(data)
# idx = [r[0] for r in data]
# data = (islice(r, 1, None) for r in data)
# df = DataFrame(data, index=idx, columns=cols)


# print(type(wb))
# print(type(ws))
# for row in ws.rows:
#     print(type(row))
#     print(row)
    # for cell in row:
    #     print(cell.value)



import argparse

parser = argparse.ArgumentParser(description="Program służy do wyszukiwania terminów na które można przełożyć dane zajęcia\n")

parser.add_argument('-p', '--page', help="Nazwa zakładki arkusza", required=True)
parser.add_argument('-r', '--row', help="Wiersz dla którego zamiany wyszukać", type=int, required=True)


argv = parser.parse_args()
row = argv.row
semestr = argv.page

print("row: {}\nsemestr: {}".format(row, semestr))


