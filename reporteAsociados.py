import xlwt
import sys
import DBManage as db
from datetime import date

'''
DB connection
'''
DBName = 'db1'
DBUser = 'postgresii'
DBPass = 'a1'
DBHost = ''

connection = db.DBConnect(DBName, DBUser, DBPass, DBHost)

cur = connection.cursor()

'''
Function: query
Params:
    location = str, db's table's name
    data = str, what to search, default value = '*'
    order1 = bool, if ordered desired, default value = False
    order2 = 0-x, ascending or desending, default value = 0 (ASC)
    criterion = str, what to order by, default value = ''
Returns:
    array with tuples of all the data collected
sample:
    query('"SVC".asociados', 'id_asociado_id, primer_nombre', True, 'id_asociado_id', 1)

'''


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
        print('Unable to run query')
        raise e

    search = cur.fetchall()
    return search


'''
Function: widths
Params:
    fields = xlwt sheet, just to get custom widths
'''


def widths(sheet):
    c1 = sheet.col(0)
    c2 = sheet.col(1)
    c3 = sheet.col(2)
    c4 = sheet.col(3)
    c5 = sheet.col(4)
    c6 = sheet.col(5)
    c7 = sheet.col(7)
    c8 = sheet.col(8)
    c1.width = 256 * 3
    c2.width = 256 * 8
    c3.width = 256 * 32
    c4.width = 256 * 16
    c5.width = 256 * 10
    c6.width = 256 * 32
    c7.width = 256 * 20
    c8.width = 256 * 4

'''
Function: writeData
Params:
    tableName = str, table that contains the data in DB
    fileName = str, name of final excel
    columnNames = str array, params to write in the columns 
    fields = str array, params to search in DB
    orderCriterion = str DB column name, how to order
    order2 = int, 0 == ASC && 1 == DESC 
Returns:
    single excel file 
sample:
    writeData(filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'full_name', 'dpi', 'full_sexo', 'full_direccion', 'apellido_paterno'], orderCriteria, order2)
'''




def writeData(tableName, fileName, columnNames, fields, orderCriterion='', order2=''):
    genfield = ''
    tableName = '"SVC".'+tableName

    if orderCriterion != '':
        if orderCriterion not in fields:
            print('\n Error: ', orderCriterion, 'not in query fields\n')
            exit()
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
    sheet1.write_merge(1, 1, 0, 5, 'Listado General de Asociados', style=title)
    sheet1.write_merge(3, 3, 0, 1, 'Fecha:', style=b)
    sheet1.write(3, 2, today, style=b)
    sheet1.write(5, 0, 'No.', style=bg)

    widths(sheet1)

    # preparing fields for query
    for i in range(len(fields)):
        if i != len(fields) - 1:
            genfield += fields[i] + ', '
        else:
            genfield += fields[i]

    # query
    if orderCriterion:
        if order2:
            data = query(tableName, genfield, True, orderCriterion, order2)
        else:
            data = query(tableName, genfield, True, orderCriterion)
    else:
        data = query(tableName, genfield)

    # column names
    for i in range(len(columnNames)):
        sheet1.write(5, i + 1, columnNames[i], style=bg)
    # rest of data
    m, f = 0, 0
    dataLenght = len(data)
    for i in range(dataLenght):
        sheet1.write(6 + i, 0, i)
        sheet1.write(6 + i, 1, data[i][0])
        sheet1.write(6 + i, 2, data[i][1])
        sheet1.write(6 + i, 3, str(data[i][2]))
        sheet1.write(6 + i, 5, data[i][4])
        if data[i][3] == 'Masculino':
            sheet1.write(6 + i, 4, 'Masculino')
            m += 1
        else:
            sheet1.write(6 + i, 4, 'Femenino')
            f += 1

    # male - female counter
    sheet1.write(dataLenght + 10, 2, 'Asociados Femeninos', style=bg)
    sheet1.write(dataLenght + 11, 2, 'Asociados Masculinos', style=bg)
    sheet1.write(dataLenght + 10, 3, f, style=g)
    sheet1.write(dataLenght + 11, 3, m, style=g)

    # save work
    master.save(fileName)

'''
writeData(0, 't1.xls', ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'full_name', 'dpi', 'full_sexo', 'full_direccion', 'apellido_paterno'], 'id_asociado_id', 0)

'''

if __name__ == "__main__":
    arguments = len(sys.argv)
    if arguments == 5:
        location = sys.argv[1]
        filename = sys.argv[2]
        orderCriteria = sys.argv[3]
        order2 = int(sys.argv[4])
        writeData(location, filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'full_name', 'dpi', 'full_sexo', 'full_direccion', 'apellido_paterno'], orderCriteria, order2)
    elif arguments == 4:
        location = sys.argv[1]
        filename = sys.argv[2]
        orderCriteria = sys.argv[3]
        writeData(location, filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'full_name', 'dpi', 'full_sexo', 'full_direccion', 'apellido_paterno'], orderCriteria)
    elif arguments == 3:
        location = sys.argv[1]
        filename = sys.argv[2]
        writeData(location, filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'full_name', 'dpi', 'full_sexo', 'full_direccion', 'apellido_paterno'])
    else:
        exit()