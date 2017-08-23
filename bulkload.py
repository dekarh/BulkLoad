# -*- coding: utf-8 -*-

from mysql.connector import MySQLConnection, Error
import openpyxl
import sys
from datetime import datetime
from lib import read_config, l, lenl, get_filename


dbconfig = read_config(filename='bulkload.ini', section='mysql')
dbconn = MySQLConnection(**dbconfig)  # Открываем БД из конфиг-файла

read_cursor = dbconn.cursor()
read_cursor.execute("SELECT * FROM big WHERE id < 2")              # !!! имя таблицы !!!
a = read_cursor.fetchall()

print('\n'+ datetime.now().strftime("%H:%M:%S") +' Начинаем загрузку \n')

fields = []
sql = 'INSERT INTO big('                                            # !!! имя таблицы !!!
sql_end = ''
for i, q in enumerate(read_cursor.description):
    if i == 0:
        continue
    elif i == 1:
        sql += q[0]
        sql_end += '%s'
    else:
        fields.append(q[0])
        sql += ',' + q[0]
        sql_end += ',' + '%s'
sql += ') VALUES (' + sql_end + ')'

workbooks =  []
sheets = []
write_rows = []
for i, xlsx_file in enumerate(sys.argv):                              # Загружаем все xlsx файлы
    if i == 0:
        workbooks.append(None)
        sheets.append(None)
        continue
    xlsx_file_cut = get_filename(xlsx_file)
    workbooks.append(openpyxl.load_workbook(filename=xlsx_file, read_only=True))
    sheets.append(workbooks[i][workbooks[i].sheetnames[0]])
    print(datetime.now().strftime("%H:%M:%S") + ' Файл ' + xlsx_file_cut + ' открыт\n')
    sheet = sheets[i]
    for j, row in enumerate(sheet.rows):                              # Теперь строки
        if j == 0:
            continue
        omit = False
        write_row = (xlsx_file_cut[0:xlsx_file_cut.rfind('.xlsx')],)
        for k, cell in enumerate(row):
            if fields[k][2:] == 'date':
                try:
                    write_row += (datetime.strptime(cell.value, "%d.%m.%Y").date(),)
                except:
                    write_row += (datetime.strptime('11.11.1111', "%d.%m.%Y").date(),)
                    print(datetime.now().strftime("%H:%M:%S") + ' В файле ' + xlsx_file_cut + ' в строке '
                      + str(j+1) + ' в поле ' + fields[k] + ' значение ' + str(cell.value) + ' сброшено до 11.11.1111')
            elif fields[k] == 'snils':
                write_row += (l(cell.value),)
                if not (lenl(cell.value) < 12 and l(cell.value) > 100):
                    omit = True
                    print('СНИЛС ' + str(cell.value) + ' пропущен')
            elif fields[k] == 'from_tbl':
                write_row += (xlsx_file_cut[0:xlsx_file_cut.rfind('.xlsx')],)
            elif fields[k] == 'id':
                q = 0
            else:
                write_row += (cell.value,)
        if not omit:
            write_rows.append(write_row)
        if j % 10000 == 0:
            write_cursor = dbconn.cursor()
            write_cursor.executemany(sql, write_rows)
            dbconn.commit()
            write_rows = []
            print(datetime.now().strftime("%H:%M:%S") + ' 10k из файла '+ xlsx_file_cut +' загрузил')
    print('\n' + datetime.now().strftime("%H:%M:%S") + ' Файл '+ xlsx_file_cut +' загружен полностью\n')

write_cursor = dbconn.cursor()
write_cursor.executemany(sql, write_rows)
dbconn.commit()

print('\n'+ datetime.now().strftime("%H:%M:%S") +' Загрузка окончена \n')