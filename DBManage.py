import psycopg2 as pg

'''
DB connection
'''
DBName = 'db1'
DBUser = 'postgresii'
DBPass = 'a1'
DBvalues = 'dbname = ' + DBName + ' user = ' + DBUser + ' password = ' + DBPass

con = pg.connect(DBvalues)
cur = con.cursor()

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
    cur.execute(line)
    search = cur.fetchall()
    return search

