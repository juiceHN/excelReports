import xlwt
import sys
import json
import DBManage as db
from datetime import date

'''
DB connection
'''

connection = db.DBConnect()

cur = connection.cursor()



def query(lowerLimit, higherLimit):
    try:
        queryToExc= str('SELECT saldos_ahorro.tipo_cuenta,'
    'saldos_ahorro.correl_cuenta,  '
    'cuentas_ahorro.fecha_apertura,'
    'asociados.primer_nombre,'
    'asociados.segundo_nombre,'
    'asociados.apellido_paterno,'
    'asociados.apellido_materno,'
    'saldos_ahorro.saldo_inicial,' 
    'saldos_ahorro.intereses_proyectados, saldos_ahorro.intereses_finales, saldos_ahorro.ingresos_programados, saldos_ahorro.ingresos_adicionales, '
    'saldos_ahorro.egresos, saldos_ahorro.auxilio_postumo, saldos_ahorro.isr, saldos_ahorro.saldo_final'
    ' FROM "SVC".saldos_ahorro, "SVC".cuentas_ahorro, "SVC".asociados'
    ' WHERE asociados.id_asociado_id = cuentas_ahorro.asociado_id'
    ' AND cuentas_ahorro.asociado_id = saldos_ahorro.asociado_id'
    ' AND cuentas_ahorro.tipo_cuenta = saldos_ahorro.tipo_cuenta'
    ' AND cuentas_ahorro.correl_cuenta = saldos_ahorro.correl_cuenta'
    ' AND saldos_ahorro.mes_saldo >' + lowerLimit+ ' AND  saldos_ahorro.mes_saldo < '+ higherLimit +' ;')
        print queryToExc
        cur.execute(queryToExc)
    except Exception as e:
        print 'FUCK'
        exit(1)
        #raise e

    search = cur.fetchall()
    return search



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


def writeData(lowerLimit, higherLimit):
   # showError = readJSON('config.json')
    print 'ehhd'
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
    columnNames = ['Cuenta', 'Fecha Apertura', 'Usuario', 'Saldo Inicial', 'Ingresos', 'Ingresos Adicionales',
              'Egresos', 'Interes Proyectado', 'Auxilio Postumo', 'Interes Pagado', 'ISR Pagado', 'Saldo Final']
    
    for i in range(len(columnNames)):
        sheet1.write(5, i + 1, columnNames[i], style=bg)
    
    # data
    
    data = query(lowerLimit, higherLimit)
    m, f = 0, 0
    dataLenght = len(data)
    print str(dataLenght)
    for i in range(dataLenght):
        print data[i]
        sheet1.write(6 + i, 0, i)
        sheet1.write(6 + i, 1, data[i][0])
        sheet1.write(6 + i, 2, data[i][1])

    master.save('reporteCuentas12.xls')

if __name__ == "__main__":
    arguments = len(sys.argv)
    if arguments == 3:
        lowerLimit = sys.argv[1]
        higherLimit = sys.argv[2]
        writeData(lowerLimit, higherLimit)
    else:
        exit(2)

exit(0)