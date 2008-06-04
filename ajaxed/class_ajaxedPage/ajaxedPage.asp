<%
'**************************************************************************************************************
'* Michal Gabrukiewicz Copyright (C) 2007 									
'* For license refer to the license.txt    									
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		AjaxedPage
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		2007-06-28 20:32
'' @CDESCRIPTION:	Represents a page which is located in physical file (e.g. index.asp). In 95% of cases
''					each of your pages will hold one instance of it which defines the page.
''					Furthermore it provides the functionality to call server-side ASP procedures
''					directly from client-side. Whenever using the class be sure the draw() method is the first
''					call before any other response is done. 
''					- use plain-property if your page is being used with in an XHR (there you dont need the whole basic structure)
''					- init(), main() and callback() need to be implemented within the page.
''					- init() is always called first and allows preperation before any content is written to the response e.g. security checks and stuff which is necessary for main() and callback() should be placed into the init().
''					- After the init always either main() or callback() is called. They are never called both within one page execution (request).
''					- main() = common state for the page which normally includes the page presentation (e.g. html elements)
''					- callback() = handles all client requests done with ajaxed.callback(). callback() needs to be defined with one parameter which holds the actual action to perform. so the signature should be sub callback(action)
''					- refer to /ajaxed/demo/ for a sample usage
''					- Requires Prototype JavaScript library (available at prototypejs.org) but can be turned on from the config
'' @VERSION:		0.1

'**************************************************************************************************************
class AjaxedPage

	'private members
	private status, jason, callbackFlagName, loadedSources, loadingText
	private ajaxHeaderDrawn, sessionCodePage
	
	'public members
	public loadPrototypeJS		''[bool] should protypeJS library be loaded. turn this off if you've done it manually. default = true
	public buffering			''[bool] turns the response.buffering on or off (no affect on callback). default = true
	public contentType			''[string] contenttype of the response. default = empty
	public debug				''[bool] turns debugging on of. default = false
	public DBConnection			''[bool] indicates if a database connection should be opened automatically for the page.
								''If yes then a connection with the configured connectionstring (ajaxed config) is established and can be used
								''within the init(), callback() and main() procedures. default = false
	public plain				''[bool] indicates if the page should be a plain page. means that no header(s) and footer(s) are included.
								''Useful for page which are actually only parts of pages and loaded with an XHR. default = false
	public title				''[string] title for the page (useful to use within the header.asp)
	public onlyDev				''[bool] should the page only be rendered on the development environment? default = false
	public defaultStructure		''[bool] should the header and the footer be generated automatically by the Page itself?
								''if true then it will load a default header.asp and footer.asp which holds the common HTML elements
								''like HTML, BODY, doctype, etc. This is useful when you dont really care about an own structure.
								''(e.g. the ajaxed console uses this). If false then the header.asp and footer.asp from the "ajaxedConfig"
								''is loaded. default = false
	
	private property get callbackAction
		callbackAction = left(RF(callbackFlagName), 255)
	end property
	
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
		loadingText = lib.init(AJAXED_LOADINGTEXT, "loading...")
		DBConnection = lib.init(AJAXED_DBCONNECTION, false)
		sessionCodePage = lib.init(AJAXED_SESSION_CODEPAGE, false)
		debug = false
		plain = false
		ajaxHeaderDrawn = false
		onlyDev = false
		defaultStructure = false
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws the page. Must be the first call on a page
	'' @DESCRIPTION:	The lifecycle is the following after executing draw().
	''					- 1. setting the HTTP Headers (response)
	''					- 2. init() (will be only called if exists in your page)
	''					- 3. main() or callback(action) - action is trimmed to 255 chars (for security reasons),
	''					ajaxed.callback(theAction, func, params, onComplete, url) is the Javascript function
	''					which can be used within your HTML to call serverside functions. The serverside execution can either
	''					return value(s) which are accessible as JSON after execution or it returns a page part (some HTML) to update the current page.
	''					The params are described as followed:
	''					- theAction: is the action which will be passed as parameter to the server-side callback function e.g. 'do'. if there is a sub starting with 'pagePart_' and the action name (e.g. for action 'do' it would be 'pagePart_do') then its treated as a page part and returns the subs content.
	''					- func: the client-side javascript function which will be called after server-side processing finished SUCCESSFULLY. e.g. done. It accepts one parameter which holds either the returned JSON (on cmmon callback) or the html of the page part. In case of a page part callback its possible to use the ID of the container which should be updated. e.g. ajaxed.callback('one', 'content') => this would update the element with ID 'content' with the response of server side 'pagePart_one()' sub
	''					- params (optional): javascript hash with POST parameters you want to pass to the callback function. They are accessible with e.g. lib.RF on within the callback function. e.g. {a:1,b:2}. if there is a form called "frm" with your page then all its values are passed as POST values to the callback (if you havent specified any params manually). 
	''					- onComplete (optional): client-side function which should be invoked when the server-side processing has been completed. This is called even if there was an error (same as with native XMLHttpRequest).
	''					- url (optional): you can specify the url of the page where the callback sub is located. Normally the page itself is called but you could also call callbacks of other pages. Note: use this also if you use callbacks on pages which are recognized as default in a folder. So when a user can call them without specifying the file itself. e.g. /demo/ (bug in iis5: http://support.microsoft.com/kb/216493)
	'**********************************************************************************************************
	public sub draw()
		logRequestDetails()
		if onlyDev and not lib.DEV then lib.error("Wrong environment.")
		setHTTPHeader()
		if DBConnection then db.openDefault()
		lib.exec "init", empty
		if isCallback() then
			set partCallback = lib.getFunction("pagePart_" & callbackAction)
			if not partCallback is nothing then
				doLog("Serving page part 'pagePart_" & callbackAction & "()'")
				writeln("pagePart:")
				partCallback
			else
				writeln("{ ""root"": {")
				status = 0
				set cb = lib.getFunction("callback")
				if cb is nothing then
					response.clear()
					lib.throwError("No callback defined. Define 'sub callback(action)' in you page in order to use ajaxed.callback")
				end if
				callback(callbackAction)
				if status = 0 then write(" null ")
				if status = 1 then
					writeln(vbcrlf & "} }")
				else
					writeln(vbcrlf & "}")
				end if
			end if
		else
			drawHeader()
			set mainFunc = lib.getFunction("main")
			if mainFunc is nothing then lib.throwError("no main() sub found.")
			mainFunc()
			drawFooter()
		end if
		if DBConnection then
			db.close()
			doLog(db.numberOfDBAccess & " database accesses. " & getLocation("FULL", true))
		end if
	end sub
	
	'******************************************************************************************************************
	'* logRequestDetails 
	'******************************************************************************************************************
	private sub logRequestDetails()
		if not lib.dev then exit sub
		method = request.serverVariables("request_method")
		if isCallback() then method = "CALLBACK (" & callbackAction & ")"
		pwd = "********"
		lines = array(method & " " & getLocation("FULL", true))
		for each f in request.form
			redim preserve lines(uBound(lines) + 1)
			'we hide all fields if they contain "password" in the fieldname
			lines(uBound(lines)) = "RF(""" & f & """): " & lib.iif(str.matching(f, "password", true), pwd, RF(f))
		next
		for each f in request.queryString
			redim preserve lines(uBound(lines) + 1)
			lines(uBound(lines)) = "QS(""" & f & """): " & lib.iif(str.matching(f, "password", true), pwd, QS(f))
		next
		doLog(lines)
	end sub
	
	'******************************************************************************************************************
	'* debug 
	'******************************************************************************************************************
	private sub doLog(msg)
		lib.logger.log 1, msg, 33
	end sub
	
	'******************************************************************************************************************
	'* drawHeader
	'******************************************************************************************************************
	private sub drawHeader()
		if not plain then
			if defaultStructure then %>
				<!--#include file="header.asp"-->
			<% else %>
				<!--#include virtual="/ajaxedConfig/header.asp"-->
			<% end if
			if not ajaxHeaderDrawn then lib.throwError("ajaxedHeader(params) must be called within the /ajaxedConfig/header.asp.")
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawFooter
	'******************************************************************************************************************
	private sub drawFooter()
		if not plain then
			if defaultStructure then %>
				<!--#include file="footer.asp"-->
			<% else %>
				<!--#include virtual="/ajaxedConfig/footer.asp"-->
			<% end if %>
		<% end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	draws all necessary headers for ajaxed (js, css, etc.)
	'' @DESCRIPTION:	call this method within the HEAD-tag of your header.asp
	'' @PARAM:			params [array]: not used yet (provide empty array)
	'******************************************************************************************************************
	public sub ajaxedHeader(params)
		if loadPrototypeJS then loadJSFile(lib.path("prototypejs/prototype.js"))
		loadJSFile(lib.path("class_ajaxedPage/ajaxed.js"))
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
			if sessionCodePage then
				session.codepage = 65001
			else
				.codePage = 65001
			end if
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
	'' @SDESCRIPTION:	OBSOLETE! use lib.throwError instead
	'**********************************************************************************************************
	public sub throwError(args)
		lib.logger.warn("AjaxedPage.throwError() is obsolete. lib.throwError() should be used instead.")
		lib.throwError(args)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use lib.error() instead.
	'******************************************************************************************************************
	public sub [error](msg)
		lib.logger.warn("AjaxedPage.error() is obsolete. lib.error() should be used instead.")
		lib.error(msg)
	end sub
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	OBSOLETE! use lib.iif instead
	'******************************************************************************************************************
	public function iif(i, j, k)
		lib.logger.warn("AjaxedPage.iif() is obsolete. lib.iif() should be used instead.")
    	iif = lib.iif(i, j, k)
	end function

end class
%>