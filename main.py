from openpyxl import Workbook, load_workbook
import pymysql

connectionPublic = pymysql.connect(host=''
                                   , user=''
                                   , password=''
                                   , db='', charset='')
connectionProduct = pymysql.connect(host=''
                                    , user=''
                                    , password='x'
                                    , db='', charset='')

cur1 = connectionPublic.cursor()

sql = "SELECT user_wish.id" \
      ", user_wish.user_id" \
      ", user_wish_item_brand.brand_id " \
      "FROM user_wish" \
      ", user_wish_item_brand " \
      "WHERE user_wish.id = user_wish_item_brand.wish_id " \
      "AND user_wish.type = 'BRAND' " \
      "AND user_wish_item_brand.is_deleted = '0';"

cur1.execute(sql)
connectionPublic.commit()

datas1 = cur1.fetchall()

cur2 = connectionProduct.cursor()

sql = "SELECT brand.brand_id, brand.name_kor FROM brand;"

cur2.execute(sql)
connectionProduct.commit()

datas2 = cur2.fetchall()
# write_wb = Workbook()
# write_wb.save("/Users/trenbe/Desktop/파이썬/브랜드찜파일3.xlsx")


wb = load_workbook("/Users/jinjoolee/Desktop/Project/python/BrandWish.xlsx")
write_ws = wb['Sheet']
# write_ws.cell(1, 1, 'wish_id')
# write_ws.cell(1, 2, 'user_id')
# write_ws.cell(1, 3, 'brand_id')
# write_ws.cell(1, 4, 'brand_name')
# wb.save("/Users/trenbe/Desktop/파이썬/브랜드찜파일3.xlsx")
# result = []


result = []
# 숫자0부터 datas1의 row수 만큼 증가시키는 반복문(증가하는 숫자 n으로 표기)
for i in range(1, int(cur1.rowcount / 2)):
    # 숫자0부터 datas2의 row수 만큼 증가시키는 반복문(증가하는 숫자 n1으로 표기)
    for j in range(1, cur2.rowcount):
        if datas1[i][2] == datas2[j][0]:
            result.append(str(datas1[i][0]))
            result.append(str(datas1[i][1]))
            result.append(str(datas1[i][2]))
            result.append(str(datas2[j][1]))
            write_ws.append(result)
            result = []
            break;
wb.save("/Users/jinjoolee/Desktop/Project/python/BrandWish.xlsx")
write_ws = wb['Sheet']

for i in range(int((cur1.rowcount / 2) + 1), int(cur1.rowcount / 2)):
    # 숫자0부터 datas2의 row수 만큼 증가시키는 반복문(증가하는 숫자 n1으로 표기)
    for j in range(1, cur2.rowcount):
        if datas1[i][2] == datas2[j][0]:
            result.append(str(datas1[i][0]))
            result.append(str(datas1[i][1]))
            result.append(str(datas1[i][2]))
            result.append(str(datas2[j][1]))
            write_ws.append(result)
            result = []
            break;

wb.save("/Users/jinjoolee/Desktop/Project/python/BrandWish.xlsx")