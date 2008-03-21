<%
'**************************************************************************************************************
'* Michal Gabrukiewicz Copyright (C) 2007 									
'* For license refer to the license.txt    									
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		AjaxedPage
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		2007-06-28 20:32
'' @CDESCRIPTION:	Represents a page which Provides the functionality to call server-side ASP procedures
''					directly from client-side. When using the class be sure the draw() method is the first
''					which is called before any response has been done.
''					init(), main() and callback() need to be implemented within the page. init() is always
''					called first and allows preperation before any content is written to the response e.g.
''					security checks and stuff which is necessary for main() and callback() should be placed
''					into the init(). After the init whether main() or callback() is called. They are never
''					called both within one page execution.
''					main() = common state for the page which shows the user's presentation
''					callback() = handles all client requests
''					callback() needs to be defined with an parameter which holds the actual action to perform.
''					so the signature should be sub callback(action)
''					- REFER TO demo.asp FOR A SAMPLE USAGE.
''					- Requires Prototype JavaScript library (available at prototypejs.org)
'' @VERSION:		0.1

'**************************************************************************************************************
class AjaxedPage

	'private members
	private status, jason, callbackFlagName, loadedSources, componentLocation, loadingText, connectionString
	private ajaxHeaderDrawn
	
	'public members
	public loadPrototypeJS		''[bool] should protypeJS library be loaded. turn this off if you've done it manually. default = true
	public buffering			''[bool] turns the response.buffering on or off (no affect on callback). default = true
	public contentType			''[string] contenttype of the response. default = empty
	public debug				''[bool] turns debugging on of. default = false
	public DBConnection			''[bool] indicates if a database connection should be opened automatically for the page.
								''If yes then a connection with the configured connectionstring is established and can be used
								''within the init(), callback() and main() procedures. default = false
	public plain				''[bool] indicates if the page should be a plain page. means that no header(s) and footer(s) are included.
								''Useful for page which are actually only parts of pages and loaded with an XHR. default = false
	public title				''[string] title for the page (useful to use within the header.asp)
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set lib.page = me
		set jason = new JSON
		jason.toResponse = true
		loadPrototypeJS = lib.init(AJAXED_LOADPROTOTYPEJS, true)
		status = -1
		buffering = lib.init(AJAXED_BUFFERING, true)
		callbackFlagName = "PageAjaxed"
		set loadedSources = server.createObject("scripting.dictionary")
		componentLocation = lib.init(AJAXED_LOCATION, "/ajaxed/")
		loadingText = lib.init(AJAXED_LOADINGTEXT, "loading...")
		connectionString = lib.init(AJAXED_CONNSTRING, empty)
		DBConnection = lib.init(AJAXED_DBCONNECTION, false)
		debug = false
		plain = false
		ajaxHeaderDrawn = false
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws the page. Must be the first call on a page
	'' @DESCRIPTION:	The lifecycle is the following after executing draw().
	''					1. setting the HTTP Headers (response)
	''					2. init()
	''					3. main() or callback(action) - action is trimmed to 255 chars (for security reasons)
	''					- dont forget to provide your HTML, HEAD and BODY tags. The page does not do this for you ;)
	'**********************************************************************************************************
	public sub draw()
		setHTTPHeader()
		if DBConnection then
			if isEmpty(connectionString) then lib.throwError("No connectionstring configured.")
			db.open(connectionString)
		end if
		lib.exec "init", empty
		if isCallback() then
			writeln("{ ""root"": {")
			status = 0
			callback(left(RF(callbackFlagName), 255))
			if status = 0 then write(" null ")
			if status = 1 then
				writeln(vbcrlf & "} }")
			else
				writeln(vbcrlf & "}")
			end if
		else
			drawHeader()
			main()
			drawFooter()
		end if
		if DBConnection then db.close()
	end sub
	
	'******************************************************************************************************************
	'* drawHeader
	'******************************************************************************************************************
	private sub drawHeader()
		if not plain then %>
			<!--#include virtual="/ajaxedConfig/header.asp"--><%
			if not ajaxHeaderDrawn then lib.throwError("ajaxedHeader(params) must be called within the /ajaxedConfig/header.asp.")
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawFooter
	'******************************************************************************************************************
	private sub drawFooter()
		if not plain then %><!--#include virtual="/ajaxedConfig/footer.asp"--><% end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	draws all necessary headers for ajaxed (js, css, etc.)
	'' @DESCRIPTION:	call this method within the HEAD-tag of your header.asp
	'' @PARAM:			params [array]: not used yet (provide empty array)
	'******************************************************************************************************************
	public sub ajaxedHeader(params)
		if loadPrototypeJS then loadJSFile(loc("prototypejs/prototype.js"))
		loadJSFile(loc("class_ajaxedPage/ajaxed.js"))
		execJS(array(_
			"ajaxed.prototype.debug = " & lib.iif(debug, "true", "false") & ";",_
			"ajaxed.prototype.indicator.innerHTML = '" & loadingText & "';"_
		))
		ajaxHeaderDrawn = true
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns a value on callback()
	'' @DESCRIPTION:	those values are accessible on the javascript callback function defined in ajaxed.callback()
	'' @PARAM:			val [variant]: check JSON.toJSON() for more details
	'******************************************************************************************************************
	public sub return(val)
		if status < 0 then throwError("return() can only be called on a callback.")
		if status < 2 then
			status = 1
			response.clear()
			writeln("{ ""root"": ")
		end if
		jason.toJSON empty, val, true
		status = 3
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns a named value on callback(). call this within the callback() sub
	'' @DESCRIPTION:	this method can be called more than once because the value will be named and therefore more 
	''					values can be returned. 
	'' @PARAM:			name [string]: name of the value (accessible within the javascript callback)
	'' @PARAM:			val [variant]: refer to JSON.toJSON() method for details
	'******************************************************************************************************************
	public sub returnValue(name, val)
		if status < 0 then throwError("returnValue() can only be called on a callback.")
		if status > 2 then throwError("returnValue() cannot be called after return() has been called.")
		if name = "" then throwError("returnValue() requires a name.")
		if status > 0 then writeln(",")
		jason.toJSON name, val, true
		status = 1
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns wheather the page was already sent to server or not.
	'' @DESCRIPTION:	It should be the same as the isPostback() from asp.net.
	'' @RETURN:			[bool] posted back (true) or not (false)
	'******************************************************************************************************************
	public function isPostback()
		isPostback = (request.form.count > 0)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @DESCRIPTION:	just a wrapper to str.write
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function write(value)
		str.write(value)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a line to the output
	'' @DESCRIPTION:	just a wrapper to str.writeln
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function writeln(value)
		str.write(value)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	executes a given javascript. input may be a string or an array. each field = a line
	'' @PARAM:			JSCode [string]. [array]: your javascript-code. e.g. window.location.reload()
	'***********************************************************************************************************
	public sub execJS(JSCode)
		writeln("<script>")
		if isArray(JSCode) then
			for i = 0 to uBound(JSCode)
				writeln(JSCode(i))
			next
		else
			writeln(JSCode)
		end if
		writeln("</script>")
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use lib.error() instead.
	'******************************************************************************************************************
	public sub [error](msg)
		lib.error(msg)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value from a given form field after postback
	'' @DESCRIPTION:	just an equivalent for request.form.
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[string] value from the request-form-collection.
	'******************************************************************************************************************
	public function RF(name)
		RF = request.form(name)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value from a given form field and encodes the string into HTML.
	''					useful if you want the value be HTML encoded. e.g. inserting into value fields
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[string] value from the request-form-collection
	'******************************************************************************************************************
	public function RFE(name)
		RFE = str.HTMLEncode(RF(name))
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the value of a given form field and trims it automatically.
	'' @PARAM:			name [string]: name of the formfield you want to get
	'' @RETURN:			[string] value from the request-form-collection
	'******************************************************************************************************************
	public function RFT(name)
		RFT = trim(RF(name))
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns true if a given value exists in the request.form
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[bool] false if there is not value returned. true if yes
	'******************************************************************************************************************
	public function RFHas(name)
		RFHas = (trim(RF(name)) <> "")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for request.querystring. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it empty to get the whole querystring
	'' @RETURN:			[string] value from the request-querystring-collection
	'******************************************************************************************************************
	public function QS(name)
		if name = "" then
			QS = request.querystring
		else
			QS = request.querystring(name)
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified javscript-file
	'' @DESCRIPTION:	it will load the file only if it has not been already loaded on the page before.
	''					so you can load the file and dont have to worry if it will be loaded more than once.
	''					differentiation between the files is the filename (case-sensitive!).
	''					does not work on callback()
	'' @PARAM:			url [string]: url of your javascript-file
	'******************************************************************************************************************
	public sub loadJSFile(url)
		if isCallback() then exit sub
		sourceID = "JS" & url
		if not loadedSources.exists(sourceID) then
			writeln("<script type=""text/javascript"" language=""javascript"" src=""" & url & """></script>")
			loadedSources.add sourceID, empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'' @DESCRIPTION:	does not work on callback()
	'' @PARAM:			url [string]: url of your stylesheet
	'' @PARAM:			media [string]: what media is this stylesheet for. screen, etc.
	''					leave it blank if its for every media
	'******************************************************************************************************************
	public sub loadCSSFile(url, media)
		if isCallback() then exit sub
		sourceID = "CSS" & url & media
		if not loadedSources.exists(sourceID) then
			writeln("<link rel=""stylesheet"" type=""text/css""" & lib.iif(media <> "", " media=""" & media & """", empty) & " href=""" & url & """ />")
			loadedSources.add sourceID, empty
		end if
	end sub
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	OBSOLETE! use lib.iif instead
	'******************************************************************************************************************
	public function iif(i, j, k)
    	iif = lib.iif(i, j, k)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	gets the location of the page you are in. virtual, physical, file or the full URL of the page
	'' @PARAM:			format [string]: the format you want the location be returned: PHYSICAL (c:\web\f.asp), VIRTUAL (/web/f.asp),
	''					FULL (http://web/f.asp) or FILE (f.asp). Full takes the protocol into consideration (https or http)
	'' @PARAM:			withQS [bool]: should the querystring be appended or not?
	'' @RETURN:			[string] the location of the executing page in the wanted format
	'******************************************************************************************************************
	public function getLocation(byVal format, withQS)
		format = lCase(format)
		with request
			getLocation = .serverVariables("SCRIPT_NAME")
			select case format
				case "physical"
					getLocation = server.mapPath(getLocation)
				case "virtual"
				case "full"
					protocol = lib.iif(lcase(.serverVariables("HTTPS")) = "off", "http://", "https://")
					getLocation = protocol & .serverVariables("SERVER_NAME") & getLocation
				case else
					parts = split(getLocation, "/")
					getLocation = parts(uBound(parts))
			end select
			if format <> "physical" and withQS and QS("") <> "" then getLocation = getLocation & "?" & QS("")
		end with
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use lib.throwError instead
	'**********************************************************************************************************
	public sub throwError(args)
		lib.throwError(args)
	end sub
	
	'******************************************************************************************************************
	'* isCallback 
	'******************************************************************************************************************
	private function isCallback()
		isCallback = RFHas(callbackFlagName)
	end function
	
	'******************************************************************************************************************
	'* setHTTPHeader 
	'******************************************************************************************************************
	private sub setHTTPHeader()
		with response
			'UTF8 is necessary when working with prototype!
			.codePage = 65001
			.charset = "utf-8"
			.expires = 0
			if not isCallback() then .buffer = buffering
			if isCallback() then
				.contentType = "application/json"
			elseif contentType <> "" then
				.contentType = contentType
			end if
		end with
	end sub
	
	'**********************************************************************************************************
	'* loc 
	'**********************************************************************************************************
	private function loc(path)
		loc = componentLocation & path
	end function

end class
%>