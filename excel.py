import mysql.connector as mysql
import xlrd
import xlsxwriter
import os 


loc = ("C:\\Users\\asus\\Desktop\\nouri\\all.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)
t = []
for i in range(sheet.nrows):
    t.append (sheet.row_values(i))
print(t)
mydb = mysql.connect(
    host="localhost",
    user="root",
    password="",
    database="mohandes"
)
mycursor = mydb.cursor(buffered = True)
workbook = xlsxwriter.Workbook('new.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
output1 = []
for i in range (len(t)):
    barcode = t[i][1]
    mycursor.execute("SELECT `barcode`,`product_name`,`code_product`,`sell_price`,IFNULL((SELECT `def_commodity_entrance_member`.`buy_price` FROM `def_commodity_entrance_member` WHERE `def_commodity_entrance_member`.`barcode` = `def_product`.`barcode` ORDER BY `def_commodity_entrance_member`.`id` DESC LIMIT 1),0 ),`group_product` FROM `def_product` WHERE `def_product`.`barcode`={}  ".format(barcode))
    foroushgah = mycursor.fetchall()

    sd1 = foroushgah[1]
    sd2 = foroushgah[2]
    sell_price = foroushgah[3]
    buy_price = foroushgah[4]
    sd5 = foroushgah[5]
    tedad_mojode = t[i][3]
    tedad_moghayerat = t[i][4]
    tedad_ghorfe = t[i][5]
#    print(sd3,tedad)
    jgheymat_foroush = float(sell_price)*float(tedad_mojode)
    jgheymat_kharid = float(buy_price)*float(tedad_mojode) 
    jgheymat_foroush_moghayerat = float(sell_price)*float(tedad_moghayerat)
    jgheymat_kharid_moghayerat = float(buy_price)*float(tedad_moghayerat) 
    jgheymat_foroush_ghorfe = float(sell_price)*float(tedad_ghorfe)
    jgheymat_kharid_ghorfe = float(buy_price)*float(tedad_ghorfe) 
    A=[]
    A.append(barcode)
    A.append(sd1)
    A.append(sd2)
    A.append(sell_price)
    A.append(buy_price)
    A.append(sd5)
    A.append(tedad_mojode)
    A.append(tedad_moghayerat)
    A.append(tedad_ghorfe)
    A.append(int(jgheymat_foroush))
    A.append(int(jgheymat_kharid))
    A.append(int(jgheymat_foroush_moghayerat))
    A.append(int(jgheymat_kharid_moghayerat))
    A.append(int(jgheymat_foroush_ghorfe))
    A.append(int(jgheymat_kharid_ghorfe))
    
    output1.append(A)
print(output1)
for item, cost in enumerate(output1):
    row += 1
    for column_number, data in enumerate(cost):
        worksheet.write(row, column_number, data)
#
workbook.close()
print('x oheye')
os.startfile("new.xlsx")