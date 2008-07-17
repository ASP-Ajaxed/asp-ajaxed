<%
'**************************************************************************************************************

'' @CLASSTITLE:		Library
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		12.09.2003
'' @STATICNAME:		lib
'' @CDESCRIPTION:	This class holds all general methods used within the library. They are accessible
''					through an already existing instance called <em>lib</em>. It represents the Ajaxed Library itself somehow ;)
''					Thats why e.g. its possible to get the current version of the library using <em>lib.version</em>
''					- environment specific configs are loaded when an instance of Library is created (thus its possible to override configs dependent on the environment). just place a sub called <em>envDEV</em> to override configs for the <em>dev</em> environment. <em>envLIVE</em> for the <em>live</em>. Note: The config var must be defined outside the sub in order to work:
''					<code>
''					<%
''					'by default we disable the logging
''					AJAXED_LOGLEVEL = 0
''					sub envDEV()
''					.	'but we enable it on the dev environment
''					.	AJAXED_LOGLEVEL = 1
''					end sub
''					% >
''					</code>
'' @VERSION:		1.0

'**************************************************************************************************************

class Library

	'private members
	private uniqueID, p_browser, libraryLocation
	
	'public members
	public page			''[AjaxedPage] holds the current executing page. Nothing if there is not page
	public fso			''[FileSystemObject] holds a filesystemobject instance for global use
	public logger		''[Logger] holds a logger instance which is ready-to-use for logging.
	
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
		version = "1.1"
	end property
	
	public property get env ''[string] gets the current environment. <em>LIVE</em> or <em>DEV</em>
		env = uCase(init(AJAXED_ENVIRONMENT, "DEV"))
		'always return development unless its really live
		if env <> "LIVE" then env = "DEV"
	end property
	
	public property get live ''[bool] indicates if the environment is the live env (production)
		LIVE = env = "LIVE"
	end property
	
	public property get dev ''[bool] indicates if the environment is the development env
		DEV = env = "DEV"
	end property
	
	'***********************************************************************************************************
	'* constructor 
	'***********************************************************************************************************
	public sub class_Initialize()
		'loads environment specific configs (only if the sub exists)
		exec "env" & env, empty
		uniqueID = 0
		p_browser = ""
		libraryLocation = init(AJAXED_LOCATION, "/ajaxed/")
		set me.page = nothing
		set fso = server.createObject("scripting.filesystemobject")
		set me.logger = nothing
	end sub
	
	'***********************************************************************************************************
	'* destructor
	'***********************************************************************************************************
	public sub class_terminate()
		set fso = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	gets the virtual path of the ajaxed library root folder.
	'' @DESCRIPTION:	useful for static files like e.g. javascript, css which are used internally by the library
	'' @PARAM:			filename [string]: some file which should be appended to the root path. e.g. <em>class_rss/rss.asp</em> would return <em>/ajaxed/class_rss/rss.asp</em> leave EMPTY if not required
	'' @RETURN:			[string] path of ajaxed root folder. e.g. <em>/ajaxed/</em>
	'**********************************************************************************************************
	public function path(filename)
		path = libraryLocation & filename
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Requires a given class to be existing. error is raised if it does not exist
	'' @PARAM:			classname [string]: name of the class you want to be required
	'' @PARAM:			caller [string]: name of the object requiring the class
	'**********************************************************************************************************
	public sub require(classname, caller)
		'we have to check the classname first for a legal classname otherwise it would be a risk bunging it into the execute
		if not str.matching(classname, "^[a-z0-9_]*$", true) then lib.throwError("lib.require 'classname' is not a valid classname.")
		on error resume next : execute("set tryToInstantiate = new " & classname)
		if err <> 0 then failed = true
		on error goto 0
		if failed then throwError array(850, caller, "'" & caller & "' requires '" & classname & "' to be included.")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Detects the first loadable server component from a given list. 
	'' @PARAM:			components [array]: names of the components you want to try to detect
	'' @RETURN:			[string] name of the component which could be loaded first or EMPTY if no one could be loaded
	'**********************************************************************************************************
	public function detectComponent(components)
		detectComponent = empty
		for each c in components
			on error resume next
				server.createObject(c)
				failed = err <> 0
			on error goto 0
			if not failed then
				detectComponent = c
				exit for
			end if
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! Checks if a given datastructure contains a given value. Use <em>DataContainer.contains()</em> instead.
	'' @DESCRIPTION:	returns FALSE if the datastructure cannot be determined
	'' @PARAM:			data [array], [dictionary]: the data structure which should be checked against.
	''					if its a dictionary then the key is used for comparison.
	'' @RETURN:			[bool] true if it contains the value
	'**********************************************************************************************************
	public function contains(byRef data, byVal val)
		logger.warn("lib.contains() is obsolete. Use DataContainer.contains() instead.")
		contains = false
		set dt = (new DataContainer)(data)
		if not dt is nothing then contains = dt.contains(val)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	generates an array for a range of values which are defined by its start and end.
	'' @PARAM:			startingWith [float], [int]: the start of the range (incl)
	'' @PARAM:			endsWith [float], [int]: the end of the range (incl)
	'' @PARAM:			interval [float], [int]: the step for the incremental increase of the starting value
	'' @RETURN:			[array] array with numbers where each value is a value between the boundaries (incl)
	'**********************************************************************************************************
	public function range(startsWith, endsWith, interval)
		if interval = 0 then throwError("interval cannot be 0")
		arr = array()
		decimals = len(str.splitValue(startsWith, ",", -1))
		decimalsE = len(str.splitValue(endsWith, ",", -1))
		decimalsI = len(str.splitValue(interval, ",", -1))
		if decimalsE > decimals then decimals = decimalsE
		if decimalsI > decimals then decimals = decimalsI
		for i = startsWith to endsWith step interval
			redim preserve arr(uBound(arr) + 1)
			i = round(i, decimals)
			arr(uBound(arr)) = i
		next
		range = arr
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	calls a given function/sub if it exists
	'' @DESCRIPTION:	tries to call a given function/sub with the given parameters.
	''					the scope is the scope where exec is called. 
	'' @PARAM:			params [variant]: you choose how you provide your params. provide EMPTY to call a procedure without parameters
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
	'' @DESCRIPTION:	if function was found it can be executed afterwards. e.g. 
	''					<code><% set f = getFunction("test") : f % ></code>
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
	'' @PARAM:			values [array]: values to fill into the dictionary. <em>array( array(key, value), arrray(key, value) )</em>.
	''					if the fields are not arrays (name value pairs) then the key is generated automatically. if no array
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
	'' @DESCRIPTION:	if you want to throw an error where just the user should be notified use <em>lib.error</em> instead
	'' @PARAM:			args [array], [string]: 
	''					- if ARRAY then fields => number, source, description
	''					- The number range for user errors is 512 (exclusive) - 1024 (exclusive)
	''					- if <em>args</em> is a string then its handled as the description and an error is raised with the
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
				if page.QS(empty) <> "" then source = source & "?" & page.QS(empty)
			end if
		end if
		logger.error(description)
		'user errors start after 512 (VB spec)
		err.raise nr, source, description
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes an error and ends the response. for unexpected errors
	'' @DESCRIPTION:	this error should be used for any unexpected errors. In comparison to <em>throwError</em>
	''					it should only be used if a common error message should be displayed to the user instead of
	''					raising a real ASP error (<em>throwError</em>).
	''					- if buffering is turned off then the already written response won't be cleared.
	''					- error is logged into the log on the <em>dev</em> env
	'' @PARAM:			msg [string]: error message
	'******************************************************************************************************************
	public sub [error](msg)
		if response.buffer then response.clear()
		str.writeln(init(AJAXED_ERRORCAPTION, "Erroro: "))
		str.writeln(str.HTMLEncode(msg))
		me.logger.error(msg)
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
	'' @SDESCRIPTION:	initializes a variable with a default value if the variable is not set set (<em>isEmpty</em>)
	'' @PARAM:			var [variant]: some variable
	'' @PARAM:			default [variant]: the default value which should be taken if the var is not set (if it is EMPTY)
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
	'' @DESCRIPTION:	It is not assured that the returned id is always the same. But it is assured that it is unique on the current page request.
	'' @RETURN:			uniqueID [int]: the unique id
	'******************************************************************************************************************
	public function getUniqueID()
		uniqueID = uniqueID + 1
		getUniqueID = uniqueID
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Opposite of <em>server.URLEncode</em>
	'' @PARAM:			- endcodedText [string]: your string which should be decoded. e.g: <em>Haxn%20Text</em> (%20 = Space)
	'' @RETURN:			[string] decoded string
	'' @DESCRIPTION: 	If you store a variable in the queryString then the variables input will be automatically
	''					encoded. Sometimes you need a function to decode this <em>%20%2H</em>, etc.
	'******************************************************************************************************************
	public function URLDecode(endcodedText)
    	decoded = endcodedText
		set rex = new Regexp
	    rex.pattern = "%[0-9,A-F]{2}"
	    rex.global = true
	    set matchCollection = rex.execute(endcodedText)
	    for each match in matchCollection
	    	decoded = replace(decoded, match.value, chr(cint("&H" & right(match.value, 2))))
	    next
	    URLDecode = decoded
    end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	This will replace the <em>IIf</em> function that is missing from the intrinsic functions of VBScript
	'' @DESCRIPTION:	Makes code more readable by placing a conditional in one line:
	''					<code><% cssClass = lib.iif(i mod 2 = 0, "even", "odd") % ></code>
	'' @PARAM:			condition [variant]: condition (must return TRUE or FALSE)
	'' @PARAM:			expression1 [variant]: expression which is returned if the condition is TRUE
	'' @PARAM:			expression2 [variant]: expression which is returned if the condition is FALSE
	'' @RETURN:			[string] returns <em>expression1</em> if <em>condition</em> is TRUE otherwise <em>expression2</em>
	'******************************************************************************************************************
	public function iif(condition, expression1, expression2)
    	if condition then iif = expression1 else iif = expression2
	end function

end class
%>
