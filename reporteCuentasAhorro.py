import xlwt
import sys
import json
import DBManage as db
import calendar
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

    # first linesssssss
    sheet1.write_merge(
        1, 1, 0, 11, 'Reporte de Cuentas de Ahorro', style=title)
    sheet1.write_merge(3, 3, 0, 1, 'Periodo:', style=b)

    #sheet1.write(3, 2, today, style=b)
    #str(int(higherLimit[4:6]))
    print type(higherLimit[:4])
    sheet1.write(3, 2, str(lowerLimit)[:4] +'-'+str(lowerLimit)[4:6] + '-01 / ' 
        + str(higherLimit)[:4]+ '-' + str(higherLimit)[4:6] + '-' + str(calendar.monthrange(int(higherLimit[:4]),int(higherLimit[4:6])) [1] ), style=b)
    sheet1.write(5, 0, 'No.', style=bg)
    widths(sheet1)

    # column names
    columnNames = ['Cuenta', 'Fecha Apertura', 'Usuario', 'Saldo Inicial', 'Ingresos', 'Ingresos Adicionales',
              'Egresos', 'Interes Proyectado', 'Auxilio Postumo', 'Interes Pagado', 'ISR Pagado', 'Saldo Final']
    
    for i in range(len(columnNames)):
        sheet1.write(5, i + 1, columnNames[i], style=bg)
    
    # data
    
    data = query(lowerLimit, higherLimit)
    si, ing, ia, e, ip, ap, ipa, isrp, sf = 0,0,0,0,0,0,0,0,0
    dataLenght = len(data)
    print str(dataLenght)
    for i in range(dataLenght):
        sheet1.write(6 + i, 0, i+1)
        sheet1.write(6 + i, 1, str(data[i][0]) + '-' +  str(data[i][1]))
        sheet1.write(6 + i, 2, str(data[i][2]))
        #full_name = (data[i][3] + ' '  + data[i][4] + ' '  +data[i][5] + ' '  +data[i][6] )
        sheet1.write(6 + i, 3, (data[i][3] + ' '  + data[i][4] + ' '  +data[i][5] + ' '  +data[i][6] ).decode('utf8'))
        si += data[i][7]
        sheet1.write(6 + i, 4, data[i][7])
        ing += data[i][8]
        sheet1.write(6 + i, 5, data[i][8])
        ia += data[i][9]
        sheet1.write(6 + i, 6, data[i][9])
        e +=  data[i][10]
        sheet1.write(6 + i, 7, data[i][10])
        ip += data[i][11]
        sheet1.write(6 + i, 8, data[i][11])
        ap += data[i][12]
        sheet1.write(6 + i, 9, data[i][12])
        ipa += data[i][13]
        sheet1.write(6 + i, 10, data[i][13])
        isrp += data[i][14]
        sheet1.write(6 + i, 11, data[i][14])
        sf += data[i][15]
        sheet1.write(6 + i, 12, data[i][15])
    sheet1.write(6+dataLenght, 3, 'Totales' , style=bg);
    sheet1.write(6+dataLenght, 4, si, style=bg );
    sheet1.write(6+dataLenght, 5, ing,style=bg );
    sheet1.write(6+dataLenght, 6, ia, style=bg );
    sheet1.write(6+dataLenght, 7, e,style=bg );
    sheet1.write(6+dataLenght, 8, ip,style=bg );
    sheet1.write(6+dataLenght, 9, ap,style=bg );
    sheet1.write(6+dataLenght, 10, ipa,style=bg );
    sheet1.write(6+dataLenght, 11, isrp,style=bg );
    sheet1.write(6+dataLenght, 12, sf,style=bg );


    print 'bef'
    master.save('reporteCuentas12.xls')
    print 'aft'

if __name__ == "__main__":
    arguments = len(sys.argv)
    if arguments == 3:
        lowerLimit = sys.argv[1]
        higherLimit = sys.argv[2]
        writeData(lowerLimit, higherLimit)
    else:
        exit(2)

exit(0)