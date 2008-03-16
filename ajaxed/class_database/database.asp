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
'' @VERSION:		0.1

'**************************************************************************************************************
class Database

	'private members
	private p_numberOfDBAccess
	
	'public members
	public connection		''[ADODB.Connection] holds the database connection
	
	public property get numberOfDBAccess ''[int] gets the number which indicates how many database accesses has been made till now
		numberOfDBAccess = p_numberOfDBAccess
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
	'' @SDESCRIPTION: 	Closes the current database connection
	'******************************************************************************************************************
	public sub close()
		if not connection is nothing then
			if connection.state <> 0 then connection.close()
			set connection = nothing
		end if
	end sub
	
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
		p_numberOfDBAccess = p_numberOfDBAccess + 1
	end function

end class
%>