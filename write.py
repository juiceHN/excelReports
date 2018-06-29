import xlwt
import DBManage as db

'''
Function: mergeData
Params:
	fields = str array, params to join in a str
Returns:
	single str 
sample:
	mergeData(['a','b','c'])
'''


def mergeData(fields):
    newValue = ''
    for i in range(len(fields)):
        if str(fields[i]) != 'None':
            newValue += str(fields[i]) + ' '
    return newValue

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
    c2.width = 256 * 6
    c3.width = 256 * 32
    c4.width = 256 * 16
    c5.width = 256 * 10
    c6.width = 256 * 32
    c7.width = 256 * 20
    c8.width = 256 * 4

'''
Function: writeData
Params:
	fileName = str, name of final excel
	columnNames = str array, params to write in the columns 
	fields = str array, params to search in DB
Returns:
	single excel file 
sample:
	mergeData('excelFile.xls', ['a'], ['queryA'])
'''

def writeData(fileName, columnNames, fields):
    genfield = ''
    # excel workbook and sheet
    master = xlwt.Workbook()
    sheet1 = master.add_sheet('Report')

    # bold
    style1 = "font: bold on"
    style = xlwt.easyxf(style1)

    # first lines
    sheet1.write(0, 0, 'Listado General de Asociados', style=style)
    sheet1.write(1, 0, 'No.', style=style)
    widths(sheet1)

    # preparing fields for query
    for i in range(len(fields)):
        if i != len(fields) - 1:
            genfield += fields[i] + ', '
        else:
            genfield += fields[i]
    # query
    data = db.query('"SVC".asociados', genfield)

    # column names
    for i in range(len(columnNames)):
        sheet1.write(1, i + 1, columnNames[i], style=style)
    # rest of data
    m, f = 0, 0
    for i in range(len(data)):
        name = mergeData([data[i][1], data[i][2], data[i][3], data[i][4]])
        address = mergeData([data[i][7], data[i][9], 'Zona', data[i][8]])
        sheet1.write(2 + i, 0, i)
        sheet1.write(2 + i, 1, data[i][0])
        sheet1.write(2 + i, 2, name)
        sheet1.write(2 + i, 3, str(data[i][5]))
        sheet1.write(2 + i, 5, address)
        if data[i][6] == 'M':
            sheet1.write(2 + i, 4, 'Masculino')
            m += 1
        else:
            sheet1.write(2 + i, 4, 'Femenino')
            f += 1

    # male - female counter
    sheet1.write(1, 7, 'Asociados Femeninos', style=style)
    sheet1.write(2, 7, 'Asociados Masculinos', style=style)
    sheet1.write(1, 8, f)
    sheet1.write(2, 8, m)

    #save work
    master.save(fileName)


writeData('reporte.xls', ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
          ['id_asociado_id', 'primer_nombre', 'segundo_nombre', 'apellido_paterno', 'apellido_materno', 'dpi', 'sexo', 'direccion', 'zona', 'barrio'])
