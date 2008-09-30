<%
'**************************************************************************************************************

'' @CLASSTITLE:		Database
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2007-07-16 21:01
'' @CDESCRIPTION:	This class offers methods for database access. All of them are accessible
''					directly through <em>db</em> without creating an own instance. The AjaxedPage
''					offers a property <em>DBConnection</em> which automatically opens and closes the default connection
''					within a page.
''					- the database type is automatically detected but it can also be set manually in the config (<em>AJAXED_DBTYPE</em>).
''					- if the database type could not be detected then the type is <em>unknown</em> and all operations are exectuted as it would be Microsoft SQL Server
'' @STATICNAME:		db
'' @COMPATIBLE:		Microsoft Sql Server, Microsoft Access, Sqlite, MySQL
'' @REQUIRES:		-
'' @VERSION:		1.0

'**************************************************************************************************************
class Database

	'private members
	private p_numberOfDBAccess, p_dbType
	
	'public members
	public connection		''[ADODB.Connection] holds the database connection
	public clientCursor		''[bool] Indicates if the <em>insert()</em> and <em>update()</em> methods should use client cursor (<em>adUseClient</em>) for the used recordset.
							''This property exists due to a bug with MySQL and UTF-8 text columns. If you have columns of type <em>text</em> in any mySQL table and you want to update it then it will
							''result in strange chars. Using a client cursor solves this problem. Dont use it in other scenarios as mySQL behaves strange when using a client cursor.
							''default = FALSE (by default its <em>adUseServer</em>).<br/><br/>(Note: use this property only if you really experience this problem)
							''More info: http://bugs.mysql.com/bug.php?id=26985 or see also http://dev.mysql.com/tech-resources/articles/vb-cursors-and-locks.html<br/>
							''Don't forget to set this property back after you have used it.
	
	public property get numberOfDBAccess ''[int] gets the number which indicates how many database accesses has been made till now
		numberOfDBAccess = p_numberOfDBAccess
	end property
	
	public property get defaultConnectionString ''[string] gets the default connectionsstring which can be configured in the ajaxed config (<em>AJAXED_CONNSTRING</em>). if no configured then EMPTY
		defaultConnectionString = lib.init(AJAXED_CONNSTRING, empty)
	end property
	
	public property get dbType ''[string] gets the type of the currently opened db. access, sqlite, mysql, mssql. unknown if the database type could not be detected. If your database could not be detected then its possible to set the type manually in the config using <em>AJAXED_DBTYPE</em>
		if connection is nothing then lib.throwError("Database.dbType needs an opened connection.")
		if isEmpty(p_dbType) then
			p_dbType = "unknown"
			typ = connection.properties("Extended Properties")
			if str.matching(typ, "ms access|microsoft access", true) then
				p_dbType = "access"
			elseif str.matching(typ, "sqlite", true) then
				p_dbType = "sqlite"
			elseif str.matching(typ, "mysql", true) then
				p_dbType = "mysql"
			elseif str.matching(connection.properties("Provider Friendly Name"), "SQL Server", true) then
				p_dbType = "sqlserver"
			end if
		end if
		dbType = p_dbType
	end property
	
	public property get stringFieldTypes ''[array] gets the field type numbers of a recordset-field which represent a string value
		stringFieldTypes = array(8, 129, 200, 201, 202, 203, 130)
	end property
	
	'******************************************************************************************************************
	'* constructor
	'******************************************************************************************************************
	public sub class_initialize()
		set connection = nothing
		p_numberOfDBAccess = 0
		p_dbType = lib.init(AJAXED_DBTYPE, empty)
		'in new line because it breaks the documentor!
		if not isEmpty(p_dbType) then p_dbType = lCase(p_dbType)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	This function is used to create a legal SQL string that you can use in an SQL statement.
	'' @DESCRIPTION:	this is necessary to pass user input directly into a SQL Query
	''					e.g. on a basic login scenario: <code><% sql = "SELECT * FROM user WHERE login = " & db.SQLSafe(username) % ></code>
	''					- Note: It may sanitize differently for the different database types.
	''					- Warn: Its highliy recommended to additionally validate user input when using parameters directly within the SQL query
	'' @PARAM:			value [string]: the value which should be made "safe"
	'' @RETURN:			[string] safe value. e.g. <em>'</em> are escaped with <em>''</em>, etc.
	'******************************************************************************************************************
	public function SQLSafe(value)
		SQLSafe = replace(value & "", "'", "''")
		'mySql need to escape more. check http://dev.mysql.com/doc/refman/5.0/en/mysql-real-escape-string.html
		if dbType = "mysql" then SQLSafe = replace(SQLSafe, "\'", "\\'")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opens a database connection with a given connection string
	'' @DESCRIPTION:	- The connection is available afterwards in the <em>connection</em> property
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
	'' @DESCRIPTION:	- uses <em>open()</em> for it
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
	'' @DESCRIPTION:	- its required that the id column is named <em>id</em> if condition is used with an INT datatype.
	''					- <em>id</em> is parsed into a number and only <em>id</em> greater 0 are recognized
	'' @PARAM:			tablename [string]: the name of the table you want to delete the record from
	'' @PARAM:			condition [int], [string]: ID of the record or a condition e.g. <em>"id = 20 AND cool = 1"</em>
	''					- if condition is a string then you need to ensure sql safety with <em>db.SQLsafe</em> manually.
	'******************************************************************************************************************
	public sub delete(tablename, byVal condition)
		checkBeforeExec "Database.delete()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "lib.delete", "tablename cannot be empty"))
		if condition = "" then exit sub
		getRS "DELETE FROM " & sqlSafe(tablename) & getWhereClause(condition), empty
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	inserts a record into a given database table and returns the record ID
	'' @DESCRIPTION:	- primary key column must be named <em>id</em>
	''					- the values are not type converted in any way. you need to do it yourself
	''					- columns with string values are trimmed if they exceed the maximum allowed length. e.g. if column A only accepts 50 chars it will be trimmed to 50 if it exceeds the length of 50
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			data [array]: array which holds the columnames and its values. e.g. <em>array("name", "jack johnson")</em>
	''					- length must be even otherwise error is thrown
	'' @RETURN:			[int] ID of the inserted record
	'******************************************************************************************************************
	public function insert(tablename, data)
		checkBeforeExec "Database.insert()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "Database.insert()", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		if clientCursor then aRS.cursorLocation = 3 'adUseClient
		aRS.open tablename, connection, 1, 2, 2
		aRS.addNew()
		fillRSWithData aRS, data, "Database.insert()"
		aRS.update()
		if dbType = "sqlite" then
			insert = getScalar("SELECT last_insert_rowid();", 0)
		elseif dbType = "mysql" then
			insert = getScalar("SELECT last_insert_id();", 0)
		else
			insert = aRS("id")
		end if
		aRS.close()
		set aRS = nothing
		debug("inserted record into '" & tablename & "' with ID " & insert)
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Updates record(s) in a given database table. Return number of updated records.
	'' @DESCRIPTION:	- Primary key column must be named <em>id</em> if condition is int
	''					- The values are not type converted in any way. you need to do it yourself
	''					- Columns with string values are trimmed if they exceed the maximum allowed length. e.g. if column A only accepts 50 chars it will be trimmed to 50 if it exceeds the length of 50
	''					- If the condition matches more records then all are updated (batch updating)
	''					<code>
	''					<%
	''					'updates the firstname of all records of the table "person"
	''					to the value "leila"
	''					updated = db.update("person", array("firstname", "leila"), empty)
	''					str.writef("Updated {0} records", updated)
	''					% >
	''					</code>
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			data [array]: array which holds the columnames and its values. e.g. <em>array("name", "jack johnson")</em>
	''					- length must be even otherwise error is thrown
	'' @PARAM:			condition [int], [string]: ID of the record or a condition e.g. <em>"id = 20 AND cool = 1"</em>
	''					- If condition is a string then you need to ensure sql safety with <em>db.SQLsafe</em> manually.
	''					- Leave EMPTY if you want to update all records of the table (Note: this could take some time when there are a lot of records).
	'' @RETURN:			[int] Number of updated records. <em>0</em> if non updated (e.g. condition didnt match any records)
	'******************************************************************************************************************
	public function update(tablename, data, byVal condition)
		checkBeforeExec "Database.update()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "Database.update()", "tablename cannot be empty"))
		set aRS = server.createObject("ADODB.Recordset")
		sql = "SELECT * FROM " & sqlSafe(tablename) & getWhereClause(condition)
		if clientCursor then aRS.cursorLocation = 3 'adUseClient
		aRS.open sql, connection, 1, 2
		update = 0
		while not aRS.eof
			fillRSWithData aRS, data, "Database.update()"
			aRS.update()
			update = update + 1
			aRS.movenext()
		wend
		aRS.close()
		set aRS = nothing
		debug(array("updated record in '" & tablename & "' with condition '" & condition & "':", sql))
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Inserts (<em>insert()</em>) a record if it does not exist otherwise it updates the record(s) (<em>update()</em>).
	'' @DESCRIPTION:	For more details about the functionality refer to the <em>insert()</em> and <em>update()</em> methods. Those are used internally. 
	''					<strong>Note:</strong> Table must contain a column named <em>id</em>.
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			data [array]: array which holds the columnames and its values.
	'' @PARAM:			condition [int], [string]: ID of the record or a condition. Must be provided (cannot be EMPTY)
	'' @RETURN:			[int] ID of the record if inserted otherwise <em>0</em>
	'******************************************************************************************************************
	public function insertOrUpdate(tablename, data, byVal condition)
		checkBeforeExec "Database.insertOrUpdate()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "Database.insertOrUpdate()", "tablename cannot be empty"))
		if trim(condition) = "" then lib.throwError("Database.insertOrUpdate() must contain a condition")
		insertOrUpdate = str.parse(condition, 0)
		countCond = condition & ""
		if countCond = "0" or insertOrUpdate > 0 then countCond = "id = " & insertOrUpdate
		c = count(tablename, countCond)
		if c > 0 then
			update tablename, data, condition
			insertOrUpdate = 0
		else
			insertOrUpdate = insert(tablename, data)
		end if
	end function
	
	'******************************************************************************************************************
	'* fillRSWithData 
	'******************************************************************************************************************
	private sub fillRSWithData(byRef myRS, byRef dataArray, callingFunctionName)
		if (uBound(dataArray) + 1) mod 2 <> 0 then lib.throwError(array(100, callingFunctionName, "data length must be even. array(column, value, ...) "))
		set strFieldTypes = (new DataContainer)(stringFieldTypes)
		for i = 0 to ubound(dataArray) step 2
			desc = ""
			col = dataArray(i)
			val = dataArray(i + 1)
			size = myRS(col).definedSize
			on error resume next
				'we trim the value to the length which is allowed by the database
				if not isNull(val) then
					if strFieldTypes.contains(myRS(col).type) then val = left(val, size)
				end if
				myRS(col) = val
				failed = err <> 0
				if failed then desc = err.description
			on error goto 0
			if failed then lib.throwError (array(100, callingFunctionName, "Error setting '" & col & "' column to value '" & val & "'. " & desc))
		next
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	gets the recordcount for a given table.
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			condition [string]: condition for the count. e.g. <em>"deleted = 0"</em>. leave empty to get all
	'' @RETURN:			[int] number of records
	'******************************************************************************************************************
	public function count(tablename, byVal condition)
		checkBeforeExec "Database.count()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "Database.count()", "tablename cannot be empty"))
		count = getScalar("SELECT COUNT(*) FROM " & SQLSafe(tablename) & lib.iif(condition <> "", " WHERE " & condition, ""), 0)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	toggles the state of a flag column. if the value is 1 its turned into 0 and vicaversa.
	'' @DESCRIPTION:	useful if you dont delete records but mark them deleted. e.g. <em>toggle("user", "deleted", 10)</em>
	'' @PARAM:			tablename [string]: name of the table
	'' @PARAM:			columnName [string]: name of the flag column. must be a numeric column accepting 1 and 0
	'' @PARAM:			condition [string], [int]: if number then treated as ID of the record otherwise condition for WHERE clause.
	'******************************************************************************************************************
	public sub toggle(tablename, columnName, byVal condition)
		checkBeforeExec "Database.toggle()", empty, false
		if trim(tablename) = "" then lib.throwError(array(100, "Database.toggle()", "tablename cannot be empty"))
		if trim(columnName) = "" then lib.throwError(array(100, "Database.toggle()", "columnname cannot be empty"))
		sql = "UPDATE " & SQLSafe(tablename) & " SET " & SQLSafe(columnName) & " = 1 - " & SQLSafe(columnName) & getWhereClause(condition)
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
	'' @SDESCRIPTION:	Gets a locked recordset
	'' @DESCRIPTION: 	Gets a LOCKED recordset from the currently opened database. Example of usage:
	''					<code>set RS = lib.getRS("SELECT * FROM users WHERE name = '{0}'", "john")</code>
	'' @PARAM:			sql [string]: Your SQL query. placeholder for params are {0}, {1}, ... check <em>str.format()</em> for details
	'' @PARAM:			params [array], [string]: parameters for the query which are used within the sql query. 
	''					- Parameters are made sql injection safe.
	''					- Leave EMPTY if no params are needed
	''					- provide an array if you have more parameters in your sql
	''					- provide a string if you have only one parameter
	'' @RETURN:			[recordset] recordset with data matching the sql query
	'******************************************************************************************************************
	public function getRS(byVal sql, params)
		sql = parametrizeSQL(sql, params, "Database.getRS()")
		debug(sql)
		on error resume next
 		set getRS = connection.execute(sql)
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(101, "Database.getRS()", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
		end if
		on error goto 0
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Gets an UNLOCKED recordset from the currently opened database. check <em>getRS()</em> documentation
	'' @PARAM:			sql [string]: check <em>getRS()</em> doc
	'' @PARAM:			params [array], [string]: check <em>getRS()</em> doc
	'' @RETURN:			[recordset] recordset with data matching the sql query
	'******************************************************************************************************************
	public function getUnlockedRS(byVal sql, params)
		sql = parametrizeSQL(sql, params, "Database.getUnlockedRS()")
		debug(sql)
		on error resume next
		set getUnlockedRS = server.createObject("ADODB.RecordSet")
		getUnlockedRS.cursorLocation = 3
		getUnlockedRS.cursorType = 3
		getUnlockedRS.open sql, connection
		if err <> 0 then
			errdesc = err.description
			on error goto 0
			lib.throwError(array(101, "Database.getUnlockedRS()", "Could not execute '" & sql & "'. Reason: " & errdesc, sql))
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
				params(i) = sqlSafe(params(i))
			next
			parametrizeSQL = str.format(parametrizeSQL, params)
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	executes a given sql-query and returns the first value of the first row.
	'' @DESCRIPTION:	If there is no record available then the alternative will be returned.
	''					The returned value (if available) will be converted to the type of <em>alternative</em>.
	''					Example:
	''					<code>set val = db.getScalar("SELECT sales FROM table", 0)</code>
	''					This will set the variable <em>val</em> to an integer for sure. If no records found then 0 will be returned otherwise the <em>sales</em> column value of the first record.
	'' @PARAM:			sql [string]: sql-query to be executed
	'' @PARAM:			alternative [variant]: what should be returned when there is no record returned
	'' @RETURN:			[variant] the first value of the result converted to the type of alternative
	''					or the alternative itself if no records available
	'******************************************************************************************************************
	public function getScalar(byVal sql, alternative)
		checkBeforeExec "Database.getScalar()", sql, true
		getScalar = alternative
		set aRS = getRS(sql, empty)
		if not aRS.eof then getScalar = str.parse(aRS(0) & "", alternative)
		set aRS = nothing
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	OBSOLETE! use <em>getRS()</em> instead.
	'' @DESCRIPTION:	Default method which should be always used to get a LOCKED recordset. Example for use:
	''					<code>set RS = db.getRecordset("SELECT * FROM user")</code>
	'' @PARAM:			sql [string]: Your SQL query
	'' @RETURN:			[Recordset] recordset-Object
	'******************************************************************************************************************
	public function getRecordset(byVal sql)
		lib.logger.warn("Database.getRecordset(" & sql & ") is obsolete and Database.getRS() should be used instead.")
		set getRecordset = getRS(sql, empty)
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use <em>getUnlockedRS()</em> instead
	'' @DESCRIPTION: 	Default method which should be always used to get an UNLOCKED recordset. Example for use:
	''					<code>set RS = db.getUnlockedRecordset("SELECT * FROM user")</code>
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