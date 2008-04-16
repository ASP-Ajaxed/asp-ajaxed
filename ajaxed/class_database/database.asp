<%
'**************************************************************************************************************

'' @CLASSTITLE:		Database
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2007-07-16 21:01
'' @CDESCRIPTION:	This class offers methods for database access. All of them are accessible
''					directly through "db" without creating an own instance. The AjaxedPage
''					offers a property "DBConnection" which automatically opens and closes the connection
''					within a page.
'' @STATICNAME:		db
'' @REQUIRES:		-
'' @VERSION:		0.2

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
	
	'******************************************************************************************************************
	'* constructor
	'******************************************************************************************************************
	public sub class_initialize()
		set connection = nothing
		p_numberOfDBAccess = 0
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
		connection.open(connectionString)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opens the connection to the default database
	'' @DESCRIPTION:	- uses open() for it
	''					- throws an error if no default connectionstring is configured
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
			set connection = nothing
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
	public sub delete(tablename, condition)
		if trim(tablename) = "" then lib.throwError(array(100, "lib.delete", "tablename cannot be empty"))
		if condition = "" then exit sub
		getRecordset("DELETE FROM " & str.sqlSafe(tablename) & getWhereClause(condition))
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
		if trim(tablename) = "" then lib.throwError(array(100, "lib.insert", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		aRS.open tablename, connection, 1, 2, 2
		aRS.addNew()
		fillRSWithData aRS, data, "db.insert"
		aRS.update()
		insert = aRS("id")
		aRS.close()
		set aRS = nothing
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
	public sub update(tablename, data, condition)
		if trim(tablename) = "" then lib.throwError(array(100, "lib.insert", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		aRS.open "SELECT * FROM " & str.sqlSafe(tablename) & getWhereClause(condition), connection, 1, 2
		fillRSWithData aRS, data, "db.update"
		aRS.update()
		aRS.close()
		set aRS = nothing
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
	public function count(tablename, condition)
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
	public sub toggle(tablename, columnName, condition)
		if trim(tablename) = "" then lib.throwError(array(100, "lib.toggle", "tablename cannot be empty"))
		if trim(columnName) = "" then lib.throwError(array(100, "lib.toggle", "columnname cannot be empty"))
		sql = "UPDATE " & str.SQLSafe(tablename) & " SET " & str.SQLSafe(columnName) & " = not " & str.SQLSafe(columnName) & getWhereClause(condition)
		getRecordset(sql)
	end sub
	
	'******************************************************************************************************************
	'* getWhereClause - generates the where clause for SQL queries 
	'******************************************************************************************************************
	private function getWhereClause(condition)
		getWhereClause = trim(condition)
		if isNumeric(condition) then
			rID = str.parse(condition, 0)
			if rID > 0 then getWhereClause = "id = " & rID
		end if
		if getWhereClause <> "" then getWhereClause = " WHERE " & getWhereClause
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Default method which should be always used to get a LOCKED recordset. Example for use:
	''					set RS = db.getRecordset("SELECT * FROM user")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[Recordset] recordset-Object
	'******************************************************************************************************************
	public function getRecordset(sql)
		if trim(sql) = "" then lib.throwError("SQL query cannot be empty")
		if connection is nothing then lib.throwError("No connection is available.")
		on error resume next
 		set getRecordset = connection.execute(sql)
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(100, "Database.getRecordset", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Default method which should be always used to get an UNLOCKED recordset. Example for use:
	''					set RS = db.getUnlockedRecordset("SELECT * FROM user")
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[Recordset] returns a recordset object (adOpenStatic & adUseClient)
	'******************************************************************************************************************
	public function getUnlockedRecordset(sql)
		if trim(sql) = "" then lib.throwError("SQL query cannot be empty")
		if connection is nothing then lib.throwError("No connection available")
		on error resume next
		set getUnlockedRecordset = server.createObject("ADODB.Recordset")
		getUnlockedRecordset.cursorLocation = 3
		getUnlockedRecordset.cursorType = 3
		getUnlockedRecordset.open sql, connection
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(100, "Database.getUnlockedRecordset", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end Function
	
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
	public function getScalar(sql, alternative)
		if trim(sql) = "" then lib.throwError("SQL query cannot be empty")
		getScalar = alternative
		set aRS = getRecordset(sql)
		if not aRS.eof then getScalar = str.parse(aRS(0) & "", alternative)
		set aRS = nothing
	end function

end class
%>