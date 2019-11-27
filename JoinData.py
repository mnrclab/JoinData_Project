# MENGAMBIL DATA DARI JSON DAN CSV, LALU DITAMBAHKAN DALAM SHEET1 DAN SHEET2 DI FILE XLSX BARU

# AMBIL DATA DARI JSON
import json
with open('file.json', 'r') as file:
    data1 = json.load(file)

judul = list(data1[0].keys())
isi = []
for i in data1:
    isi.append(list(i.values()))

# AMBIL DATA DARI CSV
import csv
with open('file.csv', 'r') as x:
    y = csv.reader(x)
    data = list(y)

header = data[0]
value = data[1:]

# MEMBUAT XLSX BARU
import xlsxwriter
file_baru = xlsxwriter.Workbook('JoinData.xlsx')
sheet1 = file_baru.add_worksheet('SheetA')
sheet2 = file_baru.add_worksheet('SheetB')

# MENGISI SHEET1 DARI CSV
for i in range(len(value)+1):
    for j in range(len(header)):
        sheet1.write(i, j, data[i][j])

# MENGISI SHEET2 DARI JSON
# #write col
for i in judul:
    sheet2.write(0, judul.index(i), i)

# #write data
row = 1
for x, y, z, q in isi:
    sheet2.write(row, 0, x)
    sheet2.write(row, 1, y)
    sheet2.write(row, 2, z)
    sheet2.write(row, 3, q)
    row += 1

file_baru.close()