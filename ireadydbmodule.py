# -*- coding: utf-8 -*-
#!/usr/bin/python

import pymysql.cursors  # sudo pip install PyMySQL


#
# Gets the db connection for our Bolt db on the local host.
#
def getiReadyDbConnection():
	# open a database connection
	# be sure to change the host IP address, username, password and database name to match your own
	connection = pymysql.connect(host = "localhost", user = "root", passwd = "", db = "lesson")

	# prepare a cursor object using cursor() method
	cursor = connection.cursor()

	# execute the SQL query using execute() method.
	cursor.execute("SELECT VERSION()")

	# fetch a single row using fetchone() method.
	row = cursor.fetchone()

	# print the row[0]
	# (Python starts the first row in an array with the number zero – instead of one)
	print "MySQL Server version:", row[0]

	# close the cursor object
	cursor.close()
	
	return connection


#
# Returns all the rows from the db that satisfy the given query.
#
def getRowsFromDbForQuery(connection, query):
	cursor = connection.cursor()

	rows = ""
	retVal = cursor.execute(query)
	rows = cursor.fetchall()
	
	cursor.close()
	
	return rows
	

#
# Returns value indicating if table was changed.
#
def changeTable(connection, schemaChangeCommand):
	cursor = connection.cursor()
	retVal = ''

	print "Executing: '" + schemaChangeCommand + "'."

	try:
		retVal = cursor.execute(schemaChangeCommand)
		connection.commit()
		retVal = retVal, ''
	except pymysql.Error, error:
		retVal = error
	finally:
		cursor.close()
	
	return retVal

#
# Returns value if all rows were changed or not.
#
def changeTableWithList(connection, schemaChangeCommands):
	retVal = ''
	retMsg = ''

	for cmd in schemaChangeCommands:
		retVal, retMsg = changeTable(connection, cmd)
		
		if (retVal <> 0 and retVal <> 1):
			print "Breaking out due to error " + str(retVal) + " and msg = " + retMsg
			break;
	
	return retVal, retMsg

