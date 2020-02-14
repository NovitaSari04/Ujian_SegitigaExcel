# Segitiga Excel

import xlsxwriter
book = xlsxwriter.Workbook("Ujian_SegitigaExcel.xlsx")
sheet = book.add_worksheet("Hasil")

def segitigaExcel(x):
    syarat = [1]
    awal = 1
    hasil = ''
    inisiasi = 0
    x = x.replace(' ', '')
    for i in range(2, len(x)):
        awal = awal + i
        syarat.append(awal)
    if len(x) in syarat :
        for i in range(syarat.index(len(x))+2):
            for j in range(i) :
                hasil = x[inisiasi]
                inisiasi += 1
                sheet.write(i-1,j, hasil)
    else :
        print("Mohon maaf, jumlah karakter tidak memenuhi syarat membentuk pola.")



segitigaExcel("Purwhadika")
segitigaExcel('Purwadhika Startup and Coding School @BSD')
segitigaExcel('kode')
segitigaExcel('kode python')
segitigaExcel('Lintang')

book.close()