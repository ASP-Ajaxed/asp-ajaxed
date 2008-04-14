<%
'**************************************************************************************************************
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Library
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		12.09.2003
'' @STATICNAME:		lib
'' @CDESCRIPTION:	This class holds all general methods used within the library. They are accessible
''					through an already existing instance called "lib". It represents the Library itself somehow.
''					Thats why e.g. its possible to get the current version of the library using lib.version
'' @VERSION:		1.0

'**************************************************************************************************************

class Library

	'private members
	private uniqueID, p_browser, errorCaption
	
	'public members
	public page			''[AjaxedPage] holds the current executing page. Nothing if there is not page
	public fso			''[FileSystemObject] holds a filesystemobject instance for global use
	
	public property get browser	''[string] gets the browser (uppercased shortcut) which is used. version is ignored. e.g. FF, IE, etc. empty if unknown
		if p_browser = "" then
			agent = uCase(request.serverVariables("HTTP_USER_AGENT"))
			if instr(agent, "MSIE") > 0 then
				p_browser = "IE"
			elseif instr(agent, "FIREFOX") > 0 then
				p_browser = "FF"
			end if
		end if
		browser = p_browser
	end property
	
	public property get version ''[string] gets the version of the whole library
		version = "1.0"
	end property
	
	public property get env ''[string] gets the current environment. LIVE or DEV
		env = lib.init(AJAXED_ENVIRONMENT, "DEV")
		'always return development unless its really live
		if env <> "LIVE" then env = "DEV"
	end property
	
	public property get LIVE ''[bool] indicates if the environment is the live env (production)
		LIVE = env = "LIVE"
	end property
	
	public property get DEV ''[bool] indicates if the environment is the development evn
		DEV = env = "DEV"
	end property
	
	'***********************************************************************************************************
	'* constructor 
	'***********************************************************************************************************
	public sub class_Initialize()
		uniqueID = 0
		p_browser = ""
		set me.page = nothing
		errorCaption = init(AJAXED_ERRORCAPTION, "Erroro: ")
		set fso = server.createObject("scripting.filesystemobject")
	end sub
	
	'***********************************************************************************************************
	'* destructor
	'***********************************************************************************************************
	public sub class_terminate()
		set fso = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	checks if a given datastructure contains a given value
	'' @DESCRIPTION:	- returns false if the datastructure cannot be determined
	''					- only array supported yet.
	'' @PARAM:			data [array]: the array which should be checked against.
	'' @RETURN:			[bool] true if it contains the value
	'**********************************************************************************************************
	public function contains(data, val)
		contains = true
		if isArray(data) then
			for each d in data
				if d & "" = val & "" then exit function
			next
		end if
		contains = false
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	generates an array for a range of values which are defined by its start and end.
	'' @PARAM:			startingWith [float], [int]: the start of the range (incl)
	'' @PARAM:			endsWith [float], [int]: the end of the range (incl)
	'' @PARAM:			interval [float], [int]: the step for the incremental increase of the starting value
	'' @RETURN:			[array] array with numbers where each value is a value between the boundaries (incl)
	'**********************************************************************************************************
	public function range(startsWith, endsWith, interval)
		if interval = 0 then lib.throwError("interval cannot be 0")
		arr = array()
		for i = startsWith to endsWith step interval
			redim preserve arr(uBound(arr) + 1)
			arr(uBound(arr)) = i
		next
		range = arr
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	calls a given function/sub if it exists
	'' @DESCRIPTION:	tries to call a given function/sub with the given parameters.
	''					the scope is the scope when calling exec. 
	'' @PARAM:			params [variant]: you choose how you provide your params. provide empty to call a procedure without parameters
	'' @RETURN:			[variant] whatever the function returns
	'**********************************************************************************************************
	public function exec(functionName, params)
		set func = getFunction(functionName)
		if func is nothing then exit function
		if isEmpty(params) then
			exec = func
		else
			exec = func(params)
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a reference to a function/sub by a given name.
	'' @DESCRIPTION:	if function was found it can be executed afterwards. eg. set f = getFunction("test") : f
	'' @RETURN:			[object] reference to the function/sub or nothing if not found
	'**********************************************************************************************************
	public function getFunction(functionName)
		set getFunction = nothing
		on error resume next
		set getFunction = getRef(functionName)
		on error goto 0
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a new dictionary filled with a list of values
	'' @PARAM:			values [array]: values to fill into the dictionary. array( array(key, value), arrray(key, value) )
	''					if the fields are not arrays (valuepairs) then the key is generated automatically. if no array
	''					provided then an empty dictionary is returned
	'' @RETURN:			[dictionary] dictionary with values.
	'**********************************************************************************************************
	public function newDict(values)
		set newDict = server.createObject("scripting.dictionary")
		if not isArray(values) then exit function
		for each v in values
			if isArray(v) then
				newDict.add v(0), v(1)
			else
				newDict.add getUniqueID(), v
			end if
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	throws an ASP runtime Error which can be handled with on error resume next
	'' @DESCRIPTION:	if you want to throw an error where just the user should be notified use lib.error instead
	'' @PARAM:			args [array], [string]: 
	''					- if array then fields => number, source, description
	''					- The number range for user errors is 512 (exclusive) - 1024 (exclusive)
	''					- if args is a string then its handled as the description and an error is raised with the
	''					number 1024 (highest possible number for user defined VBScript errors)
	'**********************************************************************************************************
	public sub throwError(args)
		if isArray(args) then
			if ubound(args) < 2 then me.throwError("To less arguments for throwError. must be (number, source, description)")
			if args(0) <= 0 then me.throwError("Error number must be greater than 0.")
			nr = 512 + args(0)
			source = args(1)
			description = args(2)
		else
			nr = 1024
			description = args
			source = request.serverVariables("SCRIPT_NAME")
			if not page is nothing then
				if page.QS(empty) <> "" then source = source & "?" & QS(empty)
			end if
		end if
		'user errors start after 512 (VB spec)
		err.raise nr, source, description
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes an error and ends the response. for unexpected errors
	'' @DESCRIPTION:	this error should be used for any unexpected errors. In comparison to throwError
	''					it should only be used if a common error message should be displayed to the user instead of
	''					raising a real ASP error (throwError).
	''					- if buffering is turned off then the already written response won't be cleared.
	'' @PARAM:			msg [string]: error message
	'******************************************************************************************************************
	public sub [error](msg)
		if response.buffer then response.clear()
		str.writeln(errorCaption)
		str.writeln(str.HTMLEncode(msg))
		str.end()
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets a new globally unique Identifier.
	'' @RETURN:			[string]: new guid without hyphens. (hexa-decimal)
	'**********************************************************************************************************
	public function getGUID()
		getGUID = ""
	    set typelib = server.createObject("scriptlet.typelib")
	    getGUID = typelib.guid
	    set typelib = nothing
	    getGUID = mid(replace(getGUID, "-", ""), 2, 32)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	initializes a variable with a default value if the variable is not set set (isEmpty)
	'' @PARAM:			var [variant]: some variable
	'' @PARAM:			default [variant]: the default value which should be taken if the var is not set
	'' @RETURN:			[variant] if var is set then the var otherwise the default value.
	'**********************************************************************************************************
	public function init(var, default)
		init = var
		if isEmpty(var) then init = default
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sleeps for a specified time. for a while :)
	'' @PARAM:			seconds [int]: how many seconds. Minmum 1, Maximum 20. Value will be autochanged if value
	''					is incorrect. 
	'******************************************************************************************************************
	public sub sleep(seconds)
	    if seconds < 1 then
		    seconds = 1
	    elseIf seconds > 20 then
		    seconds = 20
	    end if
		
	    x = timer()
	    do until x + seconds = timer()
		    'sleep
	    loop
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Returns an unique ID for each pagerequest, starting with 1
	'' @RETURN:			uniqueID [int]: the unique id
	'******************************************************************************************************************
	public function getUniqueID()
		uniqueID = uniqueID + 1
		getUniqueID = uniqueID
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opposite of server.URLEncode
	'' @PARAM:			- endcodedText [string]: your string which should be decoded. e.g: Haxn%20Text (%20 = Space)
	'' @RETURN:			[string] decoded string
	'' @DESCRIPTION: 	If you store a variable in the queryString then the variables input will be automatically
	''					encoded. Sometimes you need a function to decode this %20%2H, etc.
	'******************************************************************************************************************
	public function URLDecode(endcodedText)
    	decoded = endcodedText
	    
		set oRegExpr = server.createObject("VBScript.RegExp")
	    oRegExpr.pattern = "%[0-9,A-F]{2}"
	    oRegExpr.global = true
	    set matchCollection = oRegExpr.execute(endcodedText)
		
	    for each match in matchCollection
	    	decoded = replace(decoded, match.value, chr(cint("&H" & right(match.value, 2))))
	    next
		
	    URLDecode = decoded
    end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	This will replace the IIf function that is missing from the intrinsic functions of ASP
	'' @PARAM:			i [variant]: condition
	'' @PARAM:			j [variant]: expression 1
	'' @PARAM:			k [variant]: expression 2
	'' @RETURN:			[string]
	'******************************************************************************************************************
	public function iif(i, j, k)
    	if i then iif = j else iif = k
	end function

end class
%>
