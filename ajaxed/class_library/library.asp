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
		version = "2.0"
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
	'' @SDESCRIPTION:	Returns a random number within a given range
	'' @PARAM:			min [int]: The minimum possible number (inclusive)
	'' @PARAM:			max [int]: The maximum possible number (inclusive)
	'' @RETURN:			[int] Number within the desired range
	'**********************************************************************************************************
	public function random(min, max)
		randomize
		random = int((max - min + 1) * rnd() + min)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Creates an OPTIONSHASH out of an ARRAY.
	'' @ALIAS:			["O"]()
	'' @DESCRIPTION:	OPTIONSHASH is the term for a hash with name value pairs. Its main idea is to simulate optional parameters
	''					for VBScript methods and transform an ARRAY into an associative ARRAY (collection of unique keys and a collection of values, where each key is associated with one value).
	''					With this its possible to extend a methods input without changing the methods signature.
	''					The OPTIONSHASH is represented with a DICTIONARY (The idea of options has been taken from Ruby on Rails as those guys
	''					also often use a hash as the last argument of a function. And it works great ;)). Example of usage:
	''					<code>
	''					<%
	''					'a function which accepts options
	''					function doSomething(options)
	''					.	lib.options array("a", "b"), options, 0
	''					.	'or you can also use the alias
	''					.	["O"] array("a", "b"), options, 0
	''					.	'now the variable options has been transformed into an optionshash...
	''					.	str.write(options("a"))
	''					.	str.write(options("b"))
	''					end function
	''					
	''					'now calling the function in various ways..
	''					doSomething(array("a", 1, "b", 2))
	''					'prints: 12
	''					doSomething(array("b", "cool"))
	''					'prints: cool
	''					doSomething(empty)
	''					'prints: 00
	''					% >
	''					</code>
	''					Assurances for an OPTIONSHASH:
	''					- Keys of the DICTIONARY represent the option names and items represent the option values. One option name incl its corresponding value is called <em>option</em>.
	''					- Each OPTIONSHASH can contain <em>1-n</em> options
	''					- Each OPTIONSHASH is always an instance of a DICTIONARY and is never NOTHING.
	''					- Option names are always <strong>lowercased</strong> strings and contain at least one character.
	''					- The order of options within an OPTIONSHASH is arbitrary.
	'' @PARAM:			optionNames [array], [string]: Name(s) of the options which are part of the OPTIONSHASH.
	'' @PARAM:			actualOptions [array]: Contains the options which will be converted into an OPTIONSHASH. It must be an ARRAY with an even amount of fields or EMPTY (if no options).
	''					Each odd field represents the name of the options and each even field its value. Example:
	''					<code>
	''					<%
	''					'valid options
	''					array("option1", "value1", "option2", "value2")
	''					array("option1", empty)
	''					array()
	''					'invalid options
	''					array("option1", "value1", "option2")
	''					% >
	''					</code>
	''					<strong>Important:</strong> This variable is an OPTIONSHASH (DICTIONARY) afterwards!
	'' @PARAM:			defaultValues [variant], [array]: The default value(s) for each option which does not exist within the <em>options</em>. Usually EMPTY.
	''					Its possible to provide an ARRAY and give each option a different default value. If there are more options then default values then the last
	''					defaut value is used. Example with more default values:
	''					<code>
	''					<%
	''					["O"] array("a", "b", "c"), empty, array(1, 0)
	''					'=> a=1, b=0, c=0
	''					["O"] array("a", "b", "c"), empty, array(1, 2, 3)
	''					'=> a=1, b=2, c=3
	''					% >
	''					</code>
	'' @RETURN:			[optionshash] Also returns the generated OPTIONSHASH. See the method details for the definition of an OPTIONSHASH
	'**********************************************************************************************************
	public function options(byVal optionNames, byRef actualOptions, byVal defaultValues)
		actualOptions = [](actualOptions)
		if (ubound(actualOptions) + 1) mod 2 <> 0 then throwError("Library.options() actualOptions argument must contain an even amount of fields.")
		set options = ["D"](empty)
		for i = 0 to uBound(actualOptions) step 2
			n = lCase(actualOptions(i))
			if n = "" then throwError("Library.options() cannot contain empty option names")
			options.add n, actualOptions(i + 1)
		next
		optionNames = [](optionNames)
		if uBound(optionNames) < 0 then throwError("Library.options() optionNames parameter must contain at least on field.")
		defaultValues = [](defaultValues)
		if ubound(defaultValues) = -1 then defaultValues = array(empty)
		for i = 0 to ubound(optionNames)
			name = lCase(optionNames(i))
			def = defaultValues(ubound(defaultValues))
			if ubound(defaultValues) >= i then def = defaultValues(i)
			if name = "" then throwError("Library.options() cannot contain empty option names")
			if not options.exists(name) then options.add name, def
		next
		set actualOptions = options
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Ensures that a given value will be returned as an ARRAY.
	'' @DESCRIPTION:	If the <em>value</em> is already
	''					an ARRAY then its just passed through otherwise an ARRAY is created and the first field contains the <em>value</em>.
	''					Unless: If the <em>value</em> is EMPTY then an empty ARRAY is returned.
	''					<code>
	''					<%
	''					'the same as array("x")
	''					arr = lib.arrayize("x")
	''					'the same as array("x", "y")
	''					arr = lib.arrayize(array("x", "y"))
	''					'or use the shortcut alias
	''					arr = ["O"](array(1, 2, 3))
	''					% >
	''					</code>
	'' @ALIAS:			[]()
	'' @PARAM:			value [variant]: The value which should be ensured to be an ARRAY
	'' @RETURN:			[array] Resulting ARRAY
	'**********************************************************************************************************
	public function arrayize(value)
		if isArray(value) then
			arrayize = value
		elseif isEmpty(value) then
			arrayize = array()
		else
			arrayize = array(value)
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Requests an URL and returns an <em>IXMLDOMDocument</em>.
	'' @DESCRIPTION:	Post parameters are always encoded with UTF-8.
	'' @PARAM:			method [string]: <em>GET</em> or <em>POST</em>
	'' @PARAM:			url [int]: URL for the request. Only full URLs are supported (starting with <em>http://</em>, <em>https://</em>, etc.) 
	''					or virtual URLs (starting with <em>/</em>) to request internal pages
	'' @PARAM:			params [array], [string]: Parameters for the request. Even fields of the array hold the names and the odd fields hold the corresponding values.
	''					Values are automatically <em>urlEncoded</em>.
	''					- if <em>POST</em> request then send as <em>POST</em> values (if querystring values needed then add them direclty to the URL).
	''					- if <em>GET</em> request then send via querystring.
	''					- provide EMPTY if no parameters are needed
	''					- if the value is a STRING then its just passed through to POST (as post variables) or GET (as querystring) request
	'' @PARAM:			timeout [int]: Timeout in seconds for the request. <em>0</em> means unlimited (not recommended!). Default is <em>3</em>
	'' @PARAM:			implementation [string]: Define a desired <em>IServerXMLHTTPRequest</em>. Default: <em>Msxml2.ServerXMLHTTP.3.0</em>
	'' @PARAM:			requestheader [array]: some additional request header which should be passed to the request. It must be an array where each even value represents the name and the odd values represent the value of the header.
	''					e.g. array("Accept-Language", "en")
	'' @RETURN:			[IServerXMLHTTPRequest] Returns an object which implements <em>IServerXMLHTTPRequest</em>.
	''					- It returns NOTHING if the page could not be requested at all (timout & network errors)
 	''					- You can check if request was successful by checking the <em>status</em> property (<em>200</em> means succesful).
	''					- If you need xml then use the <em>responseXML</em> property. Be sure to check if <em>responseXML.parseError.errorCode</em> is <em>0</em> before you proceed using xml.
	''					- Check your logfile to see more details for debugging
	''					- Check http://msdn.microsoft.com/en-us/library/ms754586(VS.85).aspx for more members of <em>IServerXMLHTTPRequest</em>
	'**********************************************************************************************************
	public function requestURL(method, url, params, options)
		if str.startsWith(url, "/") then
			protocol = lib.iif(lcase(request.serverVariables("HTTPS")) = "off", "http://", "https://")
			url = protocol & request.serverVariables("SERVER_NAME") & url
		end if
		if isArray(params) then
			if (uBound(params) + 1) mod 2 <> 0 then throwError("Library.requestURL() params must have an even length")
			for i = 0 to uBound(params) step 2
				pQS = pQS & params(i) & "=" & server.URLEncode(params(i + 1)) & "&"
			next
		else
			pQS = params
		end if
		["O"] array("timeout", "implementation", "requestheader"), options, array(3, "Msxml2.ServerXMLHTTP.3.0", array())
		if (uBound(options("requestheader")) + 1) mod 2 <> 0 then throwError("Library.requestURL() requestheader option must contain an even amount of fields.")
		
		set requestURL = server.createObject(options("implementation"))
		timeout = options("timeout") * 1000
		logStyle = "0;32"
		'resolve, connect, send, receive
		if timeout > 0 then requestURL.setTimeouts timeout, timeout, timeout, timeout
		
		method = uCase(method)
		if method = "GET" or method = "POST" then
			lib.logger.log 1, array(method & ": " & url, "params: " & pQS), logStyle
			if method = "GET" then
				if not str.endsWith(url, "?") and pQS <> "" then url = url & "?"
				url = url & pQS
				on error resume next
					requestURL.open "GET", url, false ', username, password can be appended if it uses basic authentication
					for i = 0 to uBound(options("requestheader")) step 2
						requestURL.setRequestHeader options("requestheader")(i), options("requestheader")(i + 1)
					next
					requestURL.send()
					if err <> 0 then desc = err.description
				on error goto 0
			else
				on error resume next
					requestURL.open "POST", url, false
					requestURL.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					requestURL.setRequestHeader "Encoding", "UTF-8"
					for i = 0 to uBound(options("requestheader")) step 2
						requestURL.setRequestHeader options("requestheader")(i), options("requestheader")(i + 1)
					next
					requestURL.send(pQS)
					if err <> 0 then desc = err.description
				on error goto 0
			end if
			if isEmpty(desc) then
				lib.logger.log 1, array("Response-code " & requestURL.status & ": ", requestURL.responseText), logStyle
			else
				lib.logger.log 1, array("Request failed: " & desc), logStyle
				set requestURL = nothing
				exit function
			end if
		else
			throwError("Library.requestURL() does not support '" & uCase(method) & "'")
		end if
	end function
	
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
	'' @SDESCRIPTION:	Gets a new dictionary filled with a list of values
	'' @ALIAS:			["D"]()
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
	''					- Error is logged into the logger
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
		on error goto 0
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes an error and ends the response. for unexpected errors
	'' @DESCRIPTION:	this error should be used for any unexpected errors. In comparison to <em>throwError</em>
	''					it should only be used if a common error message should be displayed to the user instead of
	''					raising a real ASP error (<em>throwError</em>).
	''					- if buffering is turned off then the already written response won't be cleared.
	''					- error is logged into the log
	'' @PARAM:			msg [string]: error message
	'******************************************************************************************************************
	public sub [error](msg)
		if response.buffer then response.clear()
		str.writeln(init(AJAXED_ERRORCAPTION, "Erroro: "))
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

function [](value)
	[] = lib.arrayize(value)
end function
function ["O"](optionNames, actualOptions, defaultValue)
	set ["O"] = lib.options(optionNames, actualOptions, defaultValue)
end function
function ["D"](values)
	set ["D"] = lib.newDict(values)
end function
%>
