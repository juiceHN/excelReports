import xlwt
import DBManage as db


def writeData(fileName, columnNames, fields):
    genfield = ''
    master = xlwt.Workbook()
    sheet1 = master.add_sheet('Report')
    for i in range(len(fields)):
        if i != len(fields) - 1:
            genfield += fields[i] + ', '
        else:
            genfield += fields[i]

    data = db.query('"SVC".asociados', genfield)
    for i in range(len(columnNames)):
        sheet1.write(1, i+1, columnNames[i])
    for i in range(len(data)):
    	sheet1.write(i+2, 0, i)
    	for j in range(len(data[i])):
    		sheet1.write(i+2, j+1, data[i][j])

    master.save(fileName)

writeData('reporte.xls', ['codigo', 'nombre', 'sexo'],
          ['id_asociado_id', 'primer_nombre', 'sexo'])
