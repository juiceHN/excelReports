import psycopg2 as pg

'''
DB connection
'''
def DBConnect():
	DBName = 'DB02'
	DBUser = 'postgresii'
	DBPass = 'admin'
	DBHost = 'localhost'
	# string concat
	DBvalues = 'dbname = ' + DBName + ' user = ' + DBUser + ' password = ' + DBPass

	if DBHost:
	    DBvalues += ' host = ' + DBHost

	try:
	    con = pg.connect(DBvalues)
	    return con
	except: 
	    exit(1)
