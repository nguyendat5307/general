import pandas
import xlrd
import os

df_C = pandas.read_excel('C.xlsx')
df_Kf = pandas.read_excel('Kf.xlsx')
df_Kq = pandas.read_excel('Kq.xlsx')

df_Kf = df_Kf.drop('Nhom', axis=1).drop_duplicates(['F'])
df_Kq = df_Kq.drop('Nhom', axis=1).drop_duplicates(['QorV'])

# Tinh C
print('Chon thong so:')
print(df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True))
ThongSo = input()
ThongSo = int(ThongSo)
while not(ThongSo in df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True).index):
    print('Nhap sai! Nhap lai:')
    ThongSo = input()
    ThongSo = int(ThongSo)
ThongSo = str(df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True).iloc[ThongSo, 0])

_ = os.system('clear')
print('Nguon tiep nhan xa thai co dung cho sinh hoat khong? \n1: Co\n0: Khong')
Nhom = input()
Nhom = int(Nhom)
while not(Nhom == 1 or Nhom == 0):
    print('Nhap sai! Nhap lai:')
    Nhom = input()
    Nhom = int(Nhom)
if Nhom == 1:
    Nhom = 'A'
else:
    Nhom = 'B'
C = df_C[(df_C['ThongSo'] == ThongSo) & (df_C['Nhom'] == Nhom)]['C'].values[0]

# Tinh Kf
_ = os.system('clear')
print('Luu luong nguon xa thai F (m3/24h):')
print(df_Kf['F'])
F = input()
F = int(F)
while not(F in range(0, 3)):
    print('Nhap sai! Nhap lai:')
    F = input()
    F = int(F)
Kf = df_Kf.iloc[F, 1]

# Tinh Kq
_ = os.system('clear')
print('Nguon tiep nhan xa thai la nguon nuoc dong (song, suoi, kenh, rach, ...)? \n1: Co\n0: Khong')
QorV = input()
QorV = int(QorV)
while not(QorV == 1 or QorV == 0):
    print('Nhap sai! Nhap lai:')
    QorV = input()
    QorV = int(QorV)
if QorV == 1:
    print('Luu luong nguon tiep nhan xa thai Q (m3/s):')
    print(df_Kq.iloc[0:4, 0])
    Kq = input()
    Kq = int(Kq)
    while not(Kq in range(0, 4)):
        print('Nhap sai!, nhap lai:')
        Kq = input()
        Kq = int(Kq)
else:
    print('Dung tich nguon tiep nhan xa thai V (m3):')
    print(df_Kq.iloc[4:, 0])
    Kq = input()
    Kq = int(Kq)
    while not(Kq in range(4, 7)):
        print('Nhap sai!, nhap lai:')
        Kq = input()
        Kq = int(Kq)
Kq = df_Kq.iloc[Kq, 1]

# Tinh Cmax
_ = os.system('clear')
print('Cmax cua ' + ThongSo + ' la: ' + str(C*Kf*Kq))
