<%
'**************************************************************************************************************
'* Michal Gabrukiewicz Copyright (C) 2007 									
'* For license refer to the license.txt    									
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		AjaxedPage
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		2007-06-28 20:32
'' @CDESCRIPTION:	Represents a page which is located in a physical file (e.g. index.asp). In 95% of cases
''					each of your pages will hold one instance of this class which defines the page.
''					For more details check the <a href="/ajaxed/demo/">/ajaxed/demo/</a> for a sample usage.<br><br>
''					Example of a simple page (e.g. default.asp):
''					<code>
''					<!--#include virtual="/ajaxed/ajaxed.asp"-->
''					<%
''					set page = new AjaxedPage
''					page.title = "my first ajaxed page"
''					page.draw()
''					sub main()
''					.	'the execution starts here
''					end sub
''					% >
''					</code>
''					Furthermore each page provides the functionality to call server side ASP procedures
''					directly from client-side (using the javascript function <em>ajaxed.callback()</em>). A server side procedure can either
''					return value(s) (all kind of datatypes e.g. BOOL, RECORDSET, STRING, etc.) or act as a page part which sends back whole HTML fragments.
''					In order to return values check <em>return()</em> and <em>returnValue()</em> for further details. Page parts can be used to update an existing page with some
''					HTML without using a conventional postback. Simply prefix a server side sub with <em>pagePart_</em> and call it with <em>ajaxed.callback(pagePartname, targetContainerID)</em>:
''					<code>
''					<% sub pagePart_one() % >
''					.	<strong>some bold text</strong>
''					<% end sub
''					sub main() % >
''					.	<div id="container"></div>
''					.	<button onclick="ajaxed.callback('one', 'container')">load page part one</button>
''					<% end sub % >
''					</code>
''					- <strong>Refer to the draw() method for more details about the ajaxed.callback() javascript function.</strong>
''					- Whenever using the class be sure the <em>draw()</em> method is the first call before any other response is done.
''					- The <em>main()</em> sub must be implemented within each page. You have to implement <em>callback(action)</em> sub if callbacks are used.
''					- <em>init()</em> is always called first (before <em>main()</em> and <em>callback()</em>) and allows preperations to be done before any content is written to the response e.g. security checks and stuff which is necessary before <em>main()</em> and <em>callback()</em>.
''					- After the <em>init()</em> always either <em>main()</em> or <em>callback()</em> is called. They are never called both within one page execution (request).
''					- <em>main()</em> represents the common state of the page which normally includes the page presentation (e.g. html elements)
''					- <em>callback()</em> handles all client requests done with <em>ajaxed.callback()</em>. It needs to be defined with <strong>one parameter</strong> which holds the actual action to perform. As a result the signature should be <em>callback(action)</em>. Example of a valid callback sub
''					<code>
''					<%
''					sub callback(action)
''					.	if action = "add" then page.return(2 + 3)
''					end sub
''					% >
''					</code>
''					- Requires Prototype JavaScript library (available at prototypejs.org) but can be loaded automatically by ajaxed (turn on in the config)
''					- access querystring and form fields using page methods:
''					<code>
''					<%
''					sub main()
''					.	id = page.QS("id") ' = request.queryString("id")
''					.	name = page.RF("name") ' = request.form("name")
''					.	save = page.RFHas("save") 'checks if "save" is not empty
''					.	'automatically try to parse a querystring value into an integer.
''					.	id = str.parse(page.QS("id"), 0)
''					end sub
''					% >
''					</code>
''					- The page also supports a mechanism called "page parts". 
''					- use <em>plain</em> property if your page is being used with in an XHR (there you dont need the whole basic structure)
'' @VERSION:		0.1

'**************************************************************************************************************
class AjaxedPage

	'private members
	private status, jason, loadedSources, loadingText
	private ajaxHeaderDrawn, sessionCodePage, p_callbackType
	
	'public members
	public loadPrototypeJS		''[bool] should protypeJS library be loaded. Turn this off if you've done it manually (e.g. in your <em>header.asp</em>). default = TRUE
	public buffering			''[bool] turns the <em>response.buffering </em>on or off (no affect on callback). default = TRUE
	public contentType			''[string] contenttype of the response. default = EMPTY
	public debug				''[bool] turns debugging on of. default = FALSE
	public DBConnection			''[bool] indicates if a database connection should be opened automatically for the page.
								''If yes then a connection for the default database (default database is the one configured with <em>AJAXED_CONNSTRING</em> config) will be established and can be used
								''within the <em>init()</em>, <em>callback()</em> and <em>main()</em> procedures. default = FALSE
	public plain				''[bool] indicates if the page should be a plain page. means that no header and footer is rendered.
								''Useful for page which are actually only parts of pages and loaded with an XHR. default = FALSE
	public title				''[string] title for the page (useful to use within the header.asp)
	public onlyDev				''[bool] should the page only be rendered on the development environment? default = FALSE
	public defaultStructure		''[bool] Should the header and the footer be generated automatically by the Page itself?
								''if TRUE then it will <strong>always</strong> (ignoring all other header/footer settings) load a default <em>header.asp</em> and <em>footer.asp</em> which holds the common HTML elements
								''like <em>html</em>, <em>body</em>, a <em>doctype</em>, etc. This is useful when you dont really care about an own structure but want a valid page structure.
								''e.g. the ajaxed console uses this. If false then the <em>header.asp</em> and <em>footer.asp</em> from <em>/ajaxedConfig</em>
								''is loaded unless <em>headerFooter</em> property is set. default = FALSE
	public headerFooter			''[array] The procedures which should be executed for the header and the footer instead of loading the
								''<em>header.asp</em> and <em>footer.asp</em> from the config. Useful if you need more headers and footers within one website
								''e.g. site and backend. This example presumes a <em>pageHeader</em> and a <em>pageFooter</em> sub in your code (usually in some custom global config file):
								''<code><% page.headerFooter = array("pageHeader", "pageFooter") % ></code>
 	
	private property get callbackAction
		callbackAction = left(RF(callbackFlagName), 255)
	end property
	
	public property get callbackFlagName ''[string] gets the name of the callback flag. For advanced use only!
		callbackFlagName = "PageAjaxed"
	end property
	
	public property get callbackType ''[int] gets the type of the callback (1 = common, 2 = page part). advanced use only!
		callbackType = p_callbackType
	end property
	
	public property let callbackType(val) ''[int] sets the type of the callback. Advanced use only!
		p_callbackType = val
		if val = 2 then
			'if its a page part then we clear the response and
			'indicate that it is a page part with a "pagePart:" in the beginning of the response
			response.clear()
			writeln("pagePart:")
		end if
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
		set loadedSources = server.createObject("scripting.dictionary")
		loadingText = lib.init(AJAXED_LOADINGTEXT, "loading...")
		DBConnection = lib.init(AJAXED_DBCONNECTION, false)
		sessionCodePage = lib.init(AJAXED_SESSION_CODEPAGE, false)
		debug = false
		plain = false
		ajaxHeaderDrawn = false
		onlyDev = false
		defaultStructure = false
		callbackType = 1
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Draws the page. Must be the first call on a page
	'' @DESCRIPTION:	The lifecycle is the following after executing draw().
	''					- 1. setting the HTTP Headers (response)
	''					- 2. <em>init()</em> (will be only called if exists in your page)
	''					- 3. <em>main()</em> or <em>callback(action)</em> - action is trimmed to 255 chars (for security reasons),
	''					<em>ajaxed.callback(theAction, func, params, onComplete, url)</em> is the Javascript function
	''					which can be used within your HTML to call serverside functions and page parts. The server side execution can either
	''					return value(s) which are accessible as JSON after execution or it returns a page part (some HTML) to update the current page.
	''					The params are described as followed:
	''					- <strong>theAction</strong>: is the action which will be passed as parameter to the server-side callback function e.g. 'do'. if there is a sub starting with <em>pagePart_</em> and the action name (e.g. for action 'do' it would be <em>pagePart_do</em>) then its treated as a page part and returns the subs content.
	''					- <strong>func</strong>: the client-side javascript function which will be called after server-side processing finished SUCCESSFULLY. e.g. done. It accepts one parameter which holds either the returned JSON (on common callback) or the html of the page part. In case of a page part callback its possible to use the ID of the container which should be updated. e.g. <em>ajaxed.callback('one', 'content')</em> => this would update the element with ID <em>content</em> with the response of server side <em>pagePart_one()</em> sub
	''					- <strong>params</strong> (optional): javascript hash with POST parameters you want to pass to the callback function. They are accessible with e.g. lib.RF on within the callback function. e.g. <em>{a:1,b:2}</em>. if there is a form called <em>frm</em> with your page then all its values are passed as POST values to the callback (if you havent specified any params manually). 
	''					- <strong>onComplete</strong> (optional): client-side function which should be invoked when the server-side processing has been completed. This is called even if there was an error (same as with native XMLHttpRequest).
	''					- <strong>url</strong> (optional): you can specify the url of the page where the callback sub is located. Normally the page itself is called but you could also call callbacks of other pages. Note: use this also if you use callbacks on pages which are recognized as default in a folder. So when a user can call them without specifying the file itself. e.g. /demo/ (bug in iis5: http://support.microsoft.com/kb/216493)
	''					Calling a page part named <em>profile</em> and inserting its content into a HTML element with the ID <em>userProfile</em>:
	''					<code>ajaxed.callback('profile', 'userProfile')</code>
	''					Calling a page part named "info" and executing a javascript function "updateInfo" afterwards (which gets the content from the page part as an argument)
	''					We also send a parameter called "id" with the value 1 to the page part:
	''					<code>ajaxed.callback('info', updateInfo, {id: 1})</code>
	''					Calling a server side procedure called <em>getSales</em> which gets the current year and current month as parameters.
	''					The javascript function <em>gotSales</em> is executed after the server side processing of <em>getSales</em> has been completed:
	''					<code>ajaxed.callback('getSales', gotSales, {year: <%= year(date()) % >, month: <%= year(date()) % >})</code>
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
				callbackType = 2
				doLog("Serving page part 'pagePart_" & callbackAction & "()'")
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
				if callbackType = 1 then
					if status = 0 then write(" null ")
					if status = 1 then
						writeln(vbcrlf & "} }")
					else
						writeln(vbcrlf & "}")
					end if
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
			<%
			else
				if not callHeaderFooterFunc(0) then
					%><!--#include virtual="/ajaxedConfig/header.asp"--><%
				end if
			end if
			if not ajaxHeaderDrawn then lib.throwError("AjaxedPage.ajaxedHeader(params) must be called within the /ajaxedConfig/header.asp or the custom header (you can also use lib.page to call the ajaxedHeader() method).")
		end if
	end sub
	
	'******************************************************************************************************************
	'* drawFooter
	'******************************************************************************************************************
	private sub drawFooter()
		if not plain then
			if defaultStructure then %>
				<!--#include file="footer.asp"-->
			<%
			else
				if not callHeaderFooterFunc(1) then
					%><!--#include virtual="/ajaxedConfig/footer.asp"--><%
				end if
			end if
		end if
	end sub
	
	'******************************************************************************************************************
	'* callHeaderFooterFunc 
	'******************************************************************************************************************
	private function callHeaderFooterFunc(idx)
		callHeaderFooterFunc = false
		set func = nothing
		if isArray(headerFooter) then
			if ubound(headerFooter) >= idx then set func = lib.getFunction(headerFooter(idx))
		end if
		if not func is nothing then
			callHeaderFooterFunc = true
			func
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	draws all necessary headers for ajaxed (js, css, etc.)
	'' @DESCRIPTION:	call this method within the HEAD-tag of your header.asp
	'' @PARAM:			params [array]: not used yet, provide empty array array()
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
	'' @SDESCRIPTION:	returns a value on <em>callback()</em>
	'' @DESCRIPTION:	Those values are accessible on the javascript callback function defined in <em>ajaxed.callback()</em>.
	''					See also <em>returnValue()</em> if you consider returning more than one value.
	'' @PARAM:			val [variant]: check <em>JSON.toJSON()</em> for more details
	'******************************************************************************************************************
	public sub return(val)
		if callbackType > 1 then exit sub
		if status < 0 then throwError("AjaxedPage.return() can only be called on a callback.")
		if status < 2 then
			status = 1
			response.clear()
			writeln("{ ""root"": ")
		end if
		jason.toJSON empty, val, true
		status = 3
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns a name value pair on <em>callback()</em>.
	'' @DESCRIPTION:	this method can be called more than once because the value will be named and therefore more 
	''					values can be returned.
	'' @PARAM:			name [string]: name of the value
	'' @PARAM:			val [variant]: refer to <em>JSON.toJSON()</em> method for details
	'******************************************************************************************************************
	public sub returnValue(name, val)
		if callbackType > 1 then exit sub
		if status < 0 then throwError("AjaxedPage.returnValue() can only be called on a callback.")
		if status > 2 then throwError("AjaxedPage.returnValue() cannot be called after return() has been called.")
		if name = "" then throwError("AjaxedPage.returnValue() requires a name.")
		if status > 0 then writeln(",")
		jason.toJSON name, val, true
		status = 1
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns wheather the page was already sent to server or not.
	'' @DESCRIPTION:	It should be the same as the <em>isPostback()</em> from asp.net.
	'' @RETURN:			[bool] posted back (true) or not (false)
	'******************************************************************************************************************
	public function isPostback()
		isPostback = (request.form.count > 0)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @DESCRIPTION:	just a wrapper to <em>str.write()</em>
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function write(value)
		str.write(value)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a line to the output
	'' @DESCRIPTION:	just a wrapper to <em>str.writeln()</em>
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function writeln(value)
		str.write(value)
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	executes a given javascript. input may be a string or an array. each field = a line
	'' @PARAM:			JSCode [string]. [array]: your javascript-code. e.g. <em>window.location.reload()</em>
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
	'' @DESCRIPTION:	just an equivalent for request.form. Note: if you expecting the value to be an array
	''					(e.g. more fields with the same name) then use <em>RFA()</em> method.
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
	'' @SDESCRIPTION:	gets the value of a given form field and treats it as an array (useful if you have more form fields with the same name).
	'' @DESCRIPTION:	this is needed when you give more form fields the same name then its being posted as an array.
	''					Example of two fields with the same name being posted:
	''					<code>
	''					<%
	''					if page.isPostback() then
	''					.	str.write(page.RFA("test")(0)) 'writes 1
	''					.	str.write(page.RFA("test")(1)) 'writes 2,3
	''					end if
	''					% >
	''					<input type="text" name="test" value="1">
	''					<input type="text" name="test" value="2,3">
	''					</code>
	'' @PARAM:			name [string]: name of the formfield you want to get
	'' @RETURN:			[array] array of string values
	'******************************************************************************************************************
	public function RFA(name)
		set val = request.form(name)
		arr = array()
		redim preserve arr(val.count - 1)
		for i = 0 to uBound(arr)
			arr(i) = val(i + 1)
		next
		RFA = arr
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
	'' @DESCRIPTION:	good usage on submit buttons or checkboxes. Example of how to check a checkbox if it has been
	''					checked after posting the form:
	''					<code>
	''					<input type="checkbox" name="cb" value="1" <%= lib.iif(page.RFHas("cb", "checked", "")) % >>
	''					</code>
	'' @PARAM:			name [string]: name of the value you want to get
	'' @RETURN:			[bool] false if there is not value returned. true if yes
	'******************************************************************************************************************
	public function RFHas(name)
		RFHas = (trim(RF(name)) <> "")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for <em>request.querystring</em>. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it EMPTY to get the whole querystring
	'' @RETURN:			[string] value from the request querystring collection
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
	'' @DESCRIPTION:	it will load the file only if it has not been already loaded in the page before.
	''					so you can load the file (more times) and dont have to worry that it might be loaded more than once.
	''					Differentiation between the files is the filename (case-sensitive!).
	''					Does not work on <em>callback()</em>
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
	'' @DESCRIPTION:	Does not work on <em>callback()</em>
	'' @PARAM:			url [string]: url of your stylesheet
	'' @PARAM:			media [string]: what media is this stylesheet for. screen, etc.
	''					Leave it EMPTY if its for every media
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
	'' @PARAM:			format [string]: the format you want the location be returned: <em>PHYSICAL</em> (c:\web\f.asp), <em>VIRTUAL</em> (/web/f.asp),
	''					<em>FULL</em> (http://web/f.asp) or <em>FILE</em> (f.asp). Full takes the protocol into consideration (https or http)
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
	'' @SDESCRIPTION:	indicates if the state of the page is a callback. advanced usage only!
	'' @RETURN:			[bool] true if its a callback otherwise false
	'******************************************************************************************************************
	public function isCallback()
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
	'' @SDESCRIPTION:	OBSOLETE! use <em>lib.throwError</em> instead
	'**********************************************************************************************************
	public sub throwError(args)
		lib.logger.warn("AjaxedPage.throwError() is obsolete. lib.throwError() should be used instead.")
		lib.throwError(args)
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use <em>lib.error()</em> instead.
	'******************************************************************************************************************
	public sub [error](msg)
		lib.logger.warn("AjaxedPage.error() is obsolete. lib.error() should be used instead.")
		lib.error(msg)
	end sub
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	OBSOLETE! use <em>lib.iif</em> instead
	'******************************************************************************************************************
	public function iif(i, j, k)
		lib.logger.warn("AjaxedPage.iif() is obsolete. lib.iif() should be used instead.")
    	iif = lib.iif(i, j, k)
	end function

end class
%>