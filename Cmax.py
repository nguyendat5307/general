import pandas
import xlrd

df_C = pandas.read_excel('C.xlsx')
df_Kf = pandas.read_excel('Kf.xlsx')
df_Kq = pandas.read_excel('Kq.xlsx')

df_Kf = df_Kf.drop('Nhom', axis=1).drop_duplicates(['F'])
df_Kq = df_Kq.drop('Nhom', axis=1).drop_duplicates(['QorV'])

# Tinh C
print('Chon thong so:')
print(df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True))
ThongSo = input()
while not(ThongSo in df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True).index):
    print('Nhap sai! Nhap lai: ')
    ThongSo = input()
ThongSo = str(df_C[['ThongSo']].drop_duplicates('ThongSo').reset_index(drop=True).iloc[ThongSo, 0])

print('Nguon tiep nhan xa thai co dung cho sinh hoat khong? \n1: Co\n0: Khong')
Nhom = input()
while not(Nhom == 1 or Nhom == 0):
    print('Nhap sai! Nhap lai:')
    Nhom = input()
if Nhom == 1:
    Nhom = 'A'
else:
    Nhom = 'B'
C = df_C[(df_C['ThongSo'] == ThongSo) & (df_C['Nhom'] == Nhom)]['C'].values[0]

# Tinh Kq
print('Nguon tiep nhan xa thai la nguon nuoc dong (song, suoi, kenh, rach, ...)? \n1: Co\n0: Khong')
QorV = input()
while not(QorV == 1 or QorV == 0):
    print('Nhap sai! Nhap lai:')
    QorV = input()
if QorV == 1:
    print('Luu luong nguon tiep nhan xa thai Q (m3/s):')
    print(df_Kq.iloc[0:4, 0])
    Kq = input()
    while not(Kq in range(0, 4)):
        print('Nhap sai!, nhap lai:')
        Kq = input()
else:
    print('Dung tich nguon tiep nhan xa thai V (m3):')
    print(df_Kq.iloc[4:, 0])
    Kq = input()
    while not(Kq in range(4, 7)):
        print('Nhap sai!, nhap lai:')
        Kq = input()
Kq = df_Kq.iloc[Kq, 1]

# Tinh Kf
print('Luu luong nguon xa thai F (m3/24h):')
print(df_Kf['F'])
F = input()
while not(F in range(0, 3)):
    print('Nhap sai! Nhap lai:')
    F = input()
Kf = df_Kf.iloc[F, 1]

# Tinh Cmax
print('Cmax = ' + str(C*Kf*Kq))
