import sys
import write

if __name__ == "__main__":
    arguments = len(sys.argv)
    if arguments == 4:
        filename = sys.argv[1]
        orderCriteria = sys.argv[2]
        order2 = int(sys.argv[3])
        write.writeData(filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'primer_nombre', 'segundo_nombre', 'apellido_paterno', 'apellido_materno', 'dpi', 'sexo', 'direccion', 'zona', 'barrio'], orderCriteria, order2)
    elif arguments == 3:
        filename = sys.argv[1]
        orderCriteria = sys.argv[2]
        write.writeData(filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'primer_nombre', 'segundo_nombre', 'apellido_paterno', 'apellido_materno', 'dpi', 'sexo', 'direccion', 'zona', 'barrio'], orderCriteria)
    elif arguments == 2:
        filename = sys.argv[1]
        write.writeData(filename, ['Codigo', 'Nombre', 'DPI', 'Sexo', 'Direccion'],
                        ['id_asociado_id', 'primer_nombre', 'segundo_nombre', 'apellido_paterno', 'apellido_materno', 'dpi', 'sexo', 'direccion', 'zona', 'barrio'])
    else:
        print('''
Error: Wrong Usage

Usage: python main.py <filename>
optional: python main.py <filename> <order criteria> <order ASC = 0 or DESC = 1>

to save file in especific folder must be: 'C://foldername//filename.xls'
			''')
