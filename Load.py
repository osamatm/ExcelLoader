import pandas as pd
import os

from tkinter import Tk, filedialog
from tkinter import messagebox

root = Tk()
root.withdraw()
files = filedialog.askopenfilenames(initialdir=".",
                                    title="파일을 선택 해 주세요",
                                    filetypes=(("*.xlsx", "*xlsx"), ("*.xls", "*xls")))
if files == '':
    messagebox.showwarning("경고", "파일을 추가 하세요")

print("Selected file name : ", files[0])

x = input('Do you want to add Delimiter between each item (Y/n): ')

delim = '|'

if x == 'n' or x == 'N':
    delim = ''

selectpath = os.path.dirname(files[0])
selectfile = os.path.basename(files[0])


data_pd = pd.read_excel('{}/{}'.format(selectpath, selectfile),
                        header=None, index_col=None, names=None, engine='openpyxl')

data_np = pd.DataFrame.to_numpy(data_pd)

print("Num of Data in excel file: ", len(data_pd))

# # print(data_pd.head(2))

# print('{0:<10}'.format(data_pd[0][1]),
#       '{0:<30}'.format(data_pd[2][1]),
#       '{0:<5}'.format(data_pd[3][1]),
#       '{0:<1}'.format(data_pd[4][1]),
#       sep='')

name, ext = os.path.splitext(selectfile)

f = open(name+'.txt', 'w')

for i in range(1, len(data_pd)):
    date = data_pd[3][i]
    datefield1 = date.strftime('%Y%m%d')
    date = data_pd[21][i]
    datefield2 = date.strftime('%Y%m%d')

    targetText = ('{0:<5}'.format(data_pd[0][i]) +
                  delim +
                  '{0:<7}'.format(data_pd[1][i]) +
                  delim +
                  '{0:<12}'.format(data_pd[2][i]) +
                  delim +
                  datefield1 +
                  delim +
                  '{0:<100}'.format(data_pd[4][i]) +
                  delim +
                  '{0:<100}'.format(data_pd[5][i]) +
                  delim +
                  '{0:<1}'.format(data_pd[6][i]) +
                  delim +
                  '{0:<100}'.format(data_pd[7][i]) +
                  delim +
                  '{0:<1}'.format(data_pd[8][i]) +
                  delim +
                  '{0:<1}'.format(data_pd[9][i]) +
                  delim +
                  '{0:<30}'.format(data_pd[10][i]) +
                  delim +
                  '{0:<1}'.format(data_pd[11][i]) +
                  delim +
                  '{0:<50}'.format(data_pd[12][i]) +
                  delim +
                  '{0:<100}'.format(data_pd[13][i]) +
                  delim +
                  '{0:<10.2f}'.format(data_pd[14][i]) +
                  delim +
                  '{0:<3}'.format(data_pd[15][i]) +
                  delim +
                  '{0:<30.2f}'.format(data_pd[16][i]) +
                  delim +
                  '{0:<30.2f}'.format(data_pd[17][i]) +
                  delim +
                  '{0:<30.2f}'.format(data_pd[18][i]) +
                  delim +
                  '{0:<10.1f}'.format(data_pd[19][i]) +
                  delim +
                  '{0:<20}'.format(data_pd[20][i]) +
                  delim +
                  datefield2 +
                  '\n'
                  )
    print(targetText)
    f.write(targetText)

f.close()
