import xlwt
import sys
import json
import DBManage as db
from datetime import date

'''
DB connection
'''
'''
connection = db.DBConnect()

cur = connection.cursor()



def query(location, data='*', order1=False, criterion='', order2=0):
    line = 'SELECT ' + data + ' FROM ' + location
    if order1:
        if order2 == 0:
            line += ' ORDER BY ' + criterion + ' ASC'
        else:
            line += ' ORDER BY ' + criterion + ' DESC'
    try:
        cur.execute(line)
    except Exception as e:
        exit(1)
        #raise e

    search = cur.fetchall()
    return search
'''


def widths(sheet):
    c1 = sheet.col(0)
    c2 = sheet.col(1)
    c3 = sheet.col(2)
    c4 = sheet.col(3)
    c5 = sheet.col(4)
    c6 = sheet.col(5)
    c7 = sheet.col(6)
    c8 = sheet.col(7)
    c9 = sheet.col(8)
    c10 = sheet.col(9)
    c11 = sheet.col(10)
    c12 = sheet.col(11)
    c13 = sheet.col(12)
    c1.width = 256 * 3
    c2.width = 256 * 10
    c3.width = 256 * 16
    c4.width = 256 * 40
    c5.width = 256 * 16
    c6.width = 256 * 16
    c7.width = 256 * 20
    c8.width = 256 * 16
    c9.width = 256 * 20
    c10.width = 256 * 16
    c11.width = 256 * 16
    c12.width = 256 * 16
    c13.width = 256 * 20
'''

def readJSON(fileName):
    try:
        with open(fileName) as f:
        config = json.load(f)
    except Exception as e:
        exit(6)

    if config["errors"] == "M":
        print("showing error messages")
        return True
    else:
        return False

'''


def writeData(tableName, fileName, columnNames):
    genfield = ''
    tableName = '"SVC".' + tableName
    # showError = readJSON('config.json')

    # excel workbook and sheet
    master = xlwt.Workbook()
    sheet1 = master.add_sheet('Report')

    # styles
    # bold and grey
    bg = xlwt.easyxf(
        "font: bold on; align: horiz center; pattern: pattern solid, fore_colour grey25")
    # bold
    b = xlwt.easyxf("font: bold on")
    # bold and size
    title = xlwt.easyxf("font: height 300; align: horiz center")
    # grey
    g = xlwt.easyxf(
        "pattern: pattern solid, fore_colour grey25; align: horiz left")

    # date
    today = str(date.today())

    # first lines
    sheet1.write_merge(
        1, 1, 0, 11, 'Listado General de Asociados', style=title)
    sheet1.write_merge(3, 3, 0, 1, 'Periodo:', style=b)
    sheet1.write(3, 2, today, style=b)
    sheet1.write(5, 0, 'No.', style=bg)
    widths(sheet1)

    # column names

    for i in range(len(columnNames)):
        sheet1.write(5, i + 1, columnNames[i], style=bg)

    # data

    master.save(fileName)
titles = ['Cuenta', 'Fecha Apertura', 'Usuario', 'Saldo Inicial', 'Ingresos', 'Ingresos Adicionales',
          'Egresos', 'Interes Proyectado', 'Auxilio Postumo', 'Interes Pagado', 'ISR Pagado', 'Saldo Final']
writeData('a', 'test.xls', titles)
