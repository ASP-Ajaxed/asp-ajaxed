<%
'**************************************************************************************************************

'' @CLASSTITLE:		Database
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2007-07-16 21:01
'' @CDESCRIPTION:	This class offers methods for database access. All of them are accessible
''					directly through "db" without creating an own instance. The AjaxedPage
''					offers a property "DBConnection" which automatically opens and closes the connection
''					within a page.
''					- the database type is automatically detected but it can also be set manually in the config (AJAXED_DBTYPE).
''					- if the database type could not be detected then the type is "unknown" and all operations are exectuted as it would be Microsoft SQL Server
'' @STATICNAME:		db
'' @COMPATIBLE:		Microsoft Sql Server, Microsoft Access, Sqlite, MySQL (not fully tested yet)
'' @REQUIRES:		-
'' @VERSION:		1.0

'**************************************************************************************************************
class Database

	'private members
	private p_numberOfDBAccess
	
	'public members
	public connection		''[ADODB.Connection] holds the database connection
	
	public property get numberOfDBAccess ''[int] gets the number which indicates how many database accesses has been made till now
		numberOfDBAccess = p_numberOfDBAccess
	end property
	
	public property get defaultConnectionString ''[string] gets the default connectionsstring which can be configured in the ajaxed config. if no configured then empty
		defaultConnectionString = lib.init(AJAXED_CONNSTRING, empty)
	end property
	
	public property get dbType ''[string] gets the type of the currently opened db. access, sqlite, mysql, mssql. unknown if the database type could not be detected. If your database could not be detected then its possible to set the type manually in the config using AJAXED_DBTYPE
		if connection is nothing then lib.throwError("Database.dbType needs an opened connection.")
		if isEmpty(p_dbType) then
			p_dbType = "unknown"
			typ = connection.properties("Extended Properties")
			if str.matching(typ, "ms access|microsoft access", true) then
				p_dbType = "access"
			elseif str.matching(typ, "sqlite", true) then
				p_dbType = "sqlite"
			end if
		end if
		dbType = p_dbType
	end property
	
	'******************************************************************************************************************
	'* constructor
	'******************************************************************************************************************
	public sub class_initialize()
		set connection = nothing
		p_numberOfDBAccess = 0
		p_dbType = lib.init(AJAXED_DBTYPE, empty)
		'in new line because it breaks the documentor!
		p_dbType = lCase(p_dbType)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opens a database connection with a given connection string
	'' @DESCRIPTION:	- The connection is available afterwards in the connection property
	''					- if the connection is already opened then it gets closed a the new is opened
	'' @PARAM:			connectionString [string]: a connection string
	'******************************************************************************************************************
	public sub open(connectionString)
		if not connection is nothing then close()
		set connection = server.createObject("ADODB.connection")
		if trim(connectionString) = "" then lib.throwError("Database.open() connectionstring cannot be an empty string.")
		connection.open(connectionString)
		debug("OPENING DB (" & dbType & ")")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opens the connection to the default database
	'' @DESCRIPTION:	- uses open() for it
	''					- throws an error if default connectionstring is NOT configured
	'******************************************************************************************************************
	public sub openDefault()
		if isEmpty(defaultConnectionString) then lib.throwError("No default connectionstring configured.")
		open(defaultConnectionString)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Closes the current database connection
	'******************************************************************************************************************
	public sub close()
		if not connection is nothing then
			if connection.state <> 0 then connection.close()
			debug("CLOSING DB (" & dbType & ")")
			set connection = nothing
			p_dbType = empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	deletes a record from a given database table.
	'' @DESCRIPTION:	- its required that the id column is named "id" if condition is used with int.
	''					- ID is parsed and only ID greater 0 are recognized
	'' @PARAM:			tablename [string]: the name of the table you want to delete the record from
	'' @PARAM:			condition [int], [string]: ID of the record or a condition e.g. "id = 20 AND cool = 1"
	''					- if condition is a string then you need to ensure sql-safety with str.sqlsafe yourself.
	'******************************************************************************************************************
	public sub delete(tablename, byVal condition)
		checkBeforeExec "db.delete", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.delete", "tablename cannot be empty"))
		if condition = "" then exit sub
		getRS "DELETE FROM " & str.sqlSafe(tablename) & getWhereClause(condition), empty
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	inserts a record into a given database table and returns the record ID
	'' @DESCRIPTION:	- primary key column must be named ID
	''					- the values are not type converted in any way. you need to do it yourself
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			data [array]: array which holds the columnames and its values. e.g. array("name", "jack johnson")
	''					- length must be even otherwise error is thrown
	'' @RETURN:			[int] ID of the inserted record
	'******************************************************************************************************************
	public function insert(tablename, data)
		checkBeforeExec "db.insert", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.insert", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		aRS.open tablename, connection, 1, 2, 2
		aRS.addNew()
		fillRSWithData aRS, data, "db.insert"
		aRS.update()
		if dbType = "sqlite" then
			insert = getScalar("SELECT last_insert_rowid();", 0)
		else
			insert = aRS("id")
		end if
		aRS.close()
		set aRS = nothing
		debug("inserted record into '" & tablename & "' with ID " & insert)
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	updates a record in a given database table
	'' @DESCRIPTION:	- primary key column must be named ID if condition is int
	''					- the values are not type converted in any way. you need to do it yourself
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			data [array]: array which holds the columnames and its values. e.g. array("name", "jack johnson")
	''					- length must be even otherwise error is thrown
	'' @PARAM:			condition [int], [string]: ID of the record or a condition e.g. "id = 20 AND cool = 1"
	''					- if condition is a string then you need to ensure sql-safety with str.sqlsafe yourself.
	'******************************************************************************************************************
	public sub update(tablename, data, byVal condition)
		checkBeforeExec "db.update", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.insert", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		aRS.open "SELECT * FROM " & str.sqlSafe(tablename) & getWhereClause(condition), connection, 1, 2
		fillRSWithData aRS, data, "db.update"
		aRS.update()
		aRS.close()
		set aRS = nothing
		debug("updated record in '" & tablename & "' with condition '" & condition & "'")
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end sub
	
	'******************************************************************************************************************
	'* fillRSWithData 
	'******************************************************************************************************************
	private sub fillRSWithData(byRef RS, dataArray, callingFunctionName)
		if (uBound(dataArray) + 1) mod 2 <> 0 then lib.throwError(array(100, callingFunctionName, "data length must be even. array(column, value, ...) "))
		for i = 0 to ubound(dataArray) step 2
			desc = ""
			col = dataArray(i)
			val = dataArray(i + 1)
			on error resume next
				RS(col) = val
				failed = err <> 0
				if failed then desc = err.description
			on error goto 0
			if failed then lib.throwError (array(100, callingFunctionName, "Error setting '" & col & "' column to value '" & val & "'. " & desc))
		next
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	gets the recordcount for a given table.
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			condition [string]: condition for the count. e.g. "deleted = 0". leave empty to get all
	'' @RETURN:			[int] number of records
	'******************************************************************************************************************
	public function count(tablename, byVal condition)
		checkBeforeExec "db.count", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.count", "tablename cannot be empty"))
		count = getScalar("SELECT COUNT(*) FROM " & str.SQLSafe(tablename) & lib.iif(condition <> "", " WHERE " & condition, ""), 0)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	toggles the state of a flag column. if the value is 1 its turned into 0 and vicaversa.
	'' @DESCRIPTION:	useful if you dont delete records but mark them deleted. e.g. toggle("user", "deleted", 10)
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			columnName [string]: name of the flag column. must be a numeric column accepting 1 and 0
	'' @PARAM:			condition [string], [int]: if number then treated as ID of the record otherwise condition for WHERE clause.
	'******************************************************************************************************************
	public sub toggle(tablename, columnName, byVal condition)
		checkBeforeExec "db.toggle", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.toggle", "tablename cannot be empty"))
		if trim(columnName) = "" then lib.throwError(array(100, "lib.toggle", "columnname cannot be empty"))
		sql = "UPDATE " & str.SQLSafe(tablename) & " SET " & str.SQLSafe(columnName) & " = 1 - " & str.SQLSafe(columnName) & getWhereClause(condition)
		getRS sql, empty
	end sub
	
	'******************************************************************************************************************
	'* getWhereClause - generates the where clause for SQL queries 
	'******************************************************************************************************************
	private function getWhereClause(byVal condition)
		getWhereClause = trim(condition)
		if isNumeric(condition) then
			rID = str.parse(condition, 0)
			if rID > 0 then getWhereClause = "id = " & rID
		end if
		if getWhereClause <> "" then getWhereClause = " WHERE " & getWhereClause
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Gets a LOCKED recordset from the currently opened database. Example of usage:
	''					set RS = lib.getRS("SELECT * FROM users WHERE name = '{0}'", "john")
	'' @PARAM:			sql [string]: Your SQL query. placeholder for params are {0}, {1}, ... check str.format() for details
	'' @PARAM:			params [array], [string]: parameters for the query which are used within the sql query. 
	''					- Parameters are made sql injection safe.
	''					- Leave empty if no params are needed
	''					- provide an array if you have more parameters in your sql
	''					- provide a string if you have only one parameter
	'' @RETURN:			[recordset] recordset with data matching the sql query
	'******************************************************************************************************************
	public function getRS(byVal sql, params)
		sql = parametrizeSQL(sql, params, "db.getRS")
		debug(sql)
		on error resume next
 		set getRS = connection.execute(sql)
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(101, "db.getRS", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Gets an UNLOCKED recordset from the currently opened database. check getRS() doc
	'' @PARAM:			sql [string]: check getRS() doc
	'' @PARAM:			params [array], [string]: check getRS() doc
	'' @RETURN:			[recordset] recordset with data matching the sql query
	'******************************************************************************************************************
	public function getUnlockedRS(byVal sql, params)
		sql = parametrizeSQL(sql, params, "db.getUnlockedRS")
		debug(sql)
		on error resume next
		set getUnlockedRS = server.createObject("ADODB.RecordSet")
		getUnlockedRS.cursorLocation = 3
		getUnlockedRS.cursorType = 3
		getUnlockedRS.open sql, connection
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(101, "db.getUnlockedRS", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'* parametrizeSQL 
	'******************************************************************************************************************
	private function parametrizeSQL(byVal sql, byVal params, callingFunction)
		checkBeforeExec callingFunction, sql, true
		parametrizeSQL = sql
		if not isEmpty(params) then
			if not isArray(params) then params = array(params)
			for i = 0 to uBound(params)
				params(i) = str.sqlSafe(params(i))
			next
			parametrizeSQL = str.format(parametrizeSQL, params)
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	executes a given sql-query and returns the first value of the first row.
	'' @DESCRIPTION:	if there is no record given then the alternative will be returned.
	''					the returned value (if available) will be converted to the type of which the alternative is.
	''					example: calling getScalar("...", 0) will convert the returned value into an integer. if no record
	''					then 0 will be returned
	'' @PARAM:			sql [string]: sql-query to be executed
	'' @PARAM:			alternative [variant]: what should be returned when there is no record returned
	'' @RETURN:			[variant] the first value of the result converted to the type of alternative
	''					or the alternative itself if no records available
	'******************************************************************************************************************
	public function getScalar(byVal sql, alternative)
		checkBeforeExec "db.getScalar", sql, true
		getScalar = alternative
		set aRS = getRS(sql, empty)
		if not aRS.eof then getScalar = str.parse(aRS(0) & "", alternative)
		set aRS = nothing
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	OBSOLETE! use getRS() instead.
	'' @DESCRIPTION:	Default method which should be always used to get a LOCKED recordset. Example for use:
	''					set RS = db.getRecordset("SELECT * FROM user")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[Recordset] recordset-Object
	'******************************************************************************************************************
	public function getRecordset(byVal sql)
		lib.logger.warn("Database.getRecordset(" & sql & ") is obsolete and Database.getRS() should be used instead.")
		set getRecordset = getRS(sql, empty)
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use getUnlockedRS() instead
	'' @DESCRIPTION: 	Default method which should be always used to get an UNLOCKED recordset. Example for use:
	''					set RS = db.getUnlockedRecordset("SELECT * FROM user")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[Recordset] returns a recordset object (adOpenStatic & adUseClient)
	'******************************************************************************************************************
	public function getUnlockedRecordset(byVal sql)
		lib.logger.warn("Database.getUnlockedRecordset(" & sql & ") is obsolete and Database.getUnlockedRS() should be used instead.")
		set getUnlockedRecordset = getUnlockedRS(sql, empty)
	end Function
	
	'******************************************************************************************************************
	'* checkBeforeExec - can be used to perform common checks before executing against the DB
	'******************************************************************************************************************
	private sub checkBeforeExec(callingFunc, sql, sqlRequired)
		if sqlRequired and trim(sql) = "" then lib.throwError(array(100, callingFunction, "SQL-Query cannot be empty in " & callingFunc))
		if connection is nothing then lib.throwError("connection is nothing. Configure/Open database connection before calling " & callingFunc)
	end sub
	
	'******************************************************************************************************************
	'* debug - debugs sql stuff
	'******************************************************************************************************************
	private sub debug(msg)
		lib.logger.log 1, msg, "0;36"
	end sub

end class
%>