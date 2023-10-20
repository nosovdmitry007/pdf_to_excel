import openpyxl
import PyPDF2
import os
import pandas as pd

put_in = 'чертежи\\excel1'


def creat_put(put_in):
    ex = []
    tx = []
    ydc = 0
    for root, dirs, fiels in os.walk(put_in):
        # print(len(fiels))
        ydc+=len(fiels)
        print(ydc)
        for i in fiels:

            if i[-4:] == 'xlsx':

                ex.append(root+'\\'+i)
            if i[-3:] =='txt':
                tx.append(root+'\\'+i)

    return ex, tx


ex, tx = creat_put(put_in)
stat ={}
for i in ex:
    wb = openpyxl.load_workbook(i)
    res = len(wb.sheetnames)

    pdf_put = i.split('\\')
    pdf_put[1] = 'Чертежи для Алрино'
    pdf_put[-1] = pdf_put[-1][:-4]+'pdf'
    path = '\\'.join(pdf_put)

    pdfReader = PyPDF2.PdfFileReader(path, strict=False)
    totalPages = pdfReader.numPages

    if res + 1 == totalPages:
        stat[i] = 'Полностью распознан'
    if res + 1 < totalPages:
        stat[i] = 'Частично распознано'
    if res + 1 > totalPages:
        stat[i] = 'Частично распознано'
for j in tx:
    stat[j] = 'Не распознано'

# print(stat)
# print(len(stat))
stat_path = 'чертежи\\Статистика.xlsx'
ras = pd.DataFrame.from_dict(stat, orient='index')
with pd.ExcelWriter(stat_path, mode='w') as writer:
        ras.to_excel(writer, sheet_name='Статистика')


