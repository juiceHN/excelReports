import psycopg2 as pg

'''
DB connection
'''
def DBConnect(DBName, DBUser, DBPass, DBHost=''):
	# string concat
	DBvalues = 'dbname = ' + DBName + ' user = ' + DBUser + ' password = ' + DBPass

	if DBHost:
	    DBvalues += ' host = ' + DBHost

	try:
	    con = pg.connect(DBvalues)
	    return con
	except Exception as e:
	    print('Unable to connect to DB: ', DBName)
	    raise e
