<!--#include file="../ajaxed.asp"-->
<%
'**************************************************************************************************************
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TestFixture
'' @CREATOR:		michal
'' @CREATEDON:		2008-04-03 19:21
'' @CDESCRIPTION:	Represents a test fixture which can consist of one ore more tests.
''					- Tests must be subs which are named test_1, test_2, ...
''					- call the different assert methods within your tests
''					- if you need to debug your failures then turn on the 'debug' property
''					- create a setup sub if you need a procedure which will be called before every test
''					- run the fixture with the run-method
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class TestFixture

	'private members
	private currentTest, assertsInCurrentTest, testsMade, testsFailed, assertsMade, assertsFailed
	
	'public members
	public debug			''[bool] turn this on to debug your tests. error handlin is turned off then. default = false
	public lineBreak		''[string] the line break which should be used between each message. default = <br>
	public requestTimeout	''[int] timout for page requests. e.g. when using asserResponse() default = 3
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		if not lib.dev then lib.error("Wrong environment.")
		currentTest = empty
		testsMade = 0
		testsFailed = 0
		assertsMade = 0
		assertsFailed = 0
		assertsInCurrentTest = 0
		debug = false
		lineBreak = "<br>"
		requestTimeout = 3
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	runs the tests of the test fixture
	'' @DESCRIPTION:	executes all subs starting with 'test_' followed by a number. Numbers must start with 1
	''					- if there is a setup sub, then its called before every test
	'**********************************************************************************************************
	public sub run()
		set su = lib.getFunction("setup")
		do while true
			'try to get the test function
			currentTest = "test_" & testsMade + 1
			set test = lib.getFunction(currentTest)
			if test is nothing then exit do
			
			'call setup always before
			if not su is nothing then su
			
			'before each test we reset the number of asserts which have been progressed
			assertsInCurrentTest = 0
			if not debug then on error resume next
				test
				failed = err <> 0
				msg = err.description
			if debug then on error goto 0
			testsMade = testsMade + 1
			if failed then testFailed(msg)
		loop
		testfile = request.serverVariables("SCRIPT_NAME")
		if testsMade = 0 then
			println(testfile & ": no test(s) found.")
		else
			println(str.format("{4}: {0} tests progressed ({1} errors). {2} assertions made ({3} failed).", _
				array(testsMade, testsFailed, assertsMade, assertsFailed, testfile)))
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	checks if a given url contains a wanted response (defined with a regex pattern)
	'' @DESCRIPTION:	- will fail if the url cannot be reached
	''					- the regular expression will be executed case INsensitive
	''					- if debug is on then the whole response will be shown if the assert fails
	''					- assert will fail if the page does not return SUCCESS status code (200)
	'' @PARAM:			url [string], [array]: url to request.
	''					- if a string then it will be requested with GET
	''					- if an array then array(method, url)
	''					- only full (http://...) or virtual (starting with /) urls are allowed
	'' @PARAM:			params [array]: parameters for the request. even fields are names, odd fields the values
	''					- if POST request then send as POST values. 
	''					- if GET request then send via querystring. 
	''					- if querystring values needed on the POST request then use within the URL directly
	''					- provide 'empty' if no parameters are needed
	'' @PARAM:			pattern [string]: regex pattern which will be checked against after url has been fetched
	'**********************************************************************************************************
	public sub assertResponse(byVal url, byVal params, pattern, msg)
		assertStarted()
		if not isArray(url) then url = array("get", url)
		if uBound(url) <> 1 then lib.throwError("TestFixture.assertResponse() url parameter has wrong length")
		uri = url(1)
		'we want to catch all errors when getting the request, because we dont want
		'an error. we rather want the assert to fail
		on error resume next
			set req = getRequest(url(0), uri, params, requestTimeout)
			if err <> 0 then eDesc = err.description
		on error goto 0
		if eDesc <> "" then
			assertFailed uri, eDesc, msg
			exit sub
		end if
		resp = req.responseText
		if req.status <> 200 then
			assertFailed uri & " match '" & pattern & "'", "Status-code: " & req.status, msg
		elseif not str.matching(resp, pattern, true) then
			assertFailed uri & " match '" & pattern & "'", lib.iif(debug, resp, str.shorten(resp, 200, "...")), msg
		end if
		set xmlhttp = nothing
	end sub
	
	'**********************************************************************************************************
	'* request 
	'**********************************************************************************************************
	private function getRequest(byVal method, byVal url, byVal params, byVal timeout)
		if str.startsWith(url, "/") then
			protocol = lib.iif(lcase(request.serverVariables("HTTPS")) = "off", "http://", "https://")
			url = protocol & request.serverVariables("SERVER_NAME") & url
		end if
		if isArray(params) then
			if (uBound(params) + 1) mod 2 <> 0 then lib.throwError("getRequest() params must have an even length")
			for i = 0 to uBound(params) step 2
				pQS = pQS & params(i) & "=" & server.URLEncode(params(i + 1)) & "&"
			next
		end if
		'4.0 version cannot be used due to the following problem on WIN2003 server
		'http://support.microsoft.com/default.aspx?scid=kb;en-us;820882#6
		set getRequest = server.createObject("Msxml2.ServerXMLHTTP.3.0")
		timeout = timeout * 1000
		with getRequest
			'resolve, connect, send, receive
			.setTimeouts timeout, timeout, timeout, timeout
			if lcase(method) = "get" then
				if not str.endsWith(url, "?") then url = url & "?"
				.open "GET", url & pQS, false
				.send()
			elseif lcase(method) = "post" then
				.open "POST", url, false
				.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				.setRequestHeader "Encoding", "UTF-8"
				.send(pQS)
			else
				lib.throwError("Not supported method '" & uCase(method) & "' for " & url)
			end if
		end with
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a string to exist in a given file.
	'' @DESCRIPTION:	- will fail if the file does not exist at all as well
	''					- case sensitive
	'**********************************************************************************************************
	public sub assertInFile(virtualFilePath, stringToFind, msg)
		assertStarted()
		if not lib.fso.fileExists(server.mapPath(virtualFilePath)) then
			assertFailed stringToFind, virtualFilePath & " file not found", msg
			exit sub
		end if
		with server.createObject("ADODB.Stream")
			.charset = "utf-8"
			.open()
			.loadFromFile(server.mapPath(virtualFilePath))
			if instr(.readText(-1), stringToFind) = 0 then assertFailed stringToFind, virtualFilePath, msg
			.close()
		end with
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a string NOT to exist in a given file.
	'' @DESCRIPTION:	- will succeed if the file does not exists
	''					- case sensitive
	'**********************************************************************************************************
	public sub assertNotInFile(virtualFilePath, stringToFind, msg)
		assertStarted()
		if not lib.fso.fileExists(server.mapPath(virtualFilePath)) then exit sub
		with server.createObject("ADODB.Stream")
			.charset = "utf-8"
			.open()
			.loadFromFile(server.mapPath(virtualFilePath))
			if instr(.readText(-1), stringToFind) > 0 then assertFailed virtualFilePath, stringToFind, msg
			.close()
		end with
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be true
	'**********************************************************************************************************
	public sub assert(truth, msg)
		assertStarted()
		if not truth then assertFailed true, truth, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be false
	'**********************************************************************************************************
	public sub assertNot(expected, msg)
		assertStarted()
		if expected then assertFailed false, expected, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects two values to be equal
	'' @DESCRIPTION:	- arrays: all its values are compared against each other and must be equal and on the same position
	'**********************************************************************************************************
	public sub assertEqual(expected, actual, msg)
		assertStarted()
		if not isEqual(expected, actual) then assertFailed expected, actual, msg
	end sub
	
	'**********************************************************************************************************
	'* isEqual 
	'**********************************************************************************************************
	private function isEqual(valA, valB)
		isEqual = false
		if isObject(valA) or isObject(valB) then lib.throwError("test for equality does not support equality of objects.")
		if isArray(valA) xor isArray(valB) then exit function
		if isArray(valA) and isArray(valB) then
			if uBound(valA) <> uBound(valB) then exit function
			for i = lBound(valA) to uBound(valA)
				if not isEqual(valA(i), valB(i)) then exit function
			next
		else
			if valA <> valB then exit function
		end if
		isEqual = true
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be contained within a given data structure
	'' @PARAM:			data [array]: data must be an array. if data is no array then the assert will fail.
	'**********************************************************************************************************
	public sub assertHas(data, expected, msg)
		assertStarted()
		if not lib.contains(data, expected) then assertFailed expected, data, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value NOT to be contained within a given data structure
	'' @PARAM:			data [array]: data must be an array. if data is not an array then the assert will succeed (because its not in it)
	'**********************************************************************************************************
	public sub assertHasNot(data, expected, msg)
		assertStarted()
		if lib.contains(data, expected) then assertFailed expected, data, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects two values not to be equal
	'**********************************************************************************************************
	public sub assertNotEqual(expected, actual, msg)
		assertStarted()
		if isEqual(expected, actual) then assertFailed expected, actual, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be of a given type (class)
	'**********************************************************************************************************
	public sub assertInstanceOf(expectedClassName, value, msg)
		assertStarted()
		if lCase(typename(value)) <> lCase(expectedClassName) then assertFailed expectedClassName, typename(value), msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be of a given type (class)
	'**********************************************************************************************************
	public sub assertMatch(pattern, value, msg)
		assertStarted()
		if not str.matching(value, pattern, true) then assertFailed pattern, value, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be nothing (object)
	'**********************************************************************************************************
	public sub assertNothing(value, msg)
		assertStarted()
		if isObject(value) then
			if not value is nothing then
				assertFailed "[nothing]", typename(value), msg
			end if
		else
			assertFailed "[nothing]", value, msg
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects two value to be equal with a given delta
	'' @DESCRIPTION:	2 and 3 with delta of 1 would be equal
	'**********************************************************************************************************
	public sub assertInDelta(expected, actual, delta, msg)
		assertStarted()
		if actual + delta <> expected and actual - delta <> expected then assertFailed expected, actual, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	allows to manually fail an assert
	'' @PARAM:			msg [string]: msg which describes the failure
	'**********************************************************************************************************
	public sub fail(msg)
		assertsFailed = assertsFailed + 1
		println(str.format("Assert [{0}] in [{1}] {2}", _
			array(assertsInCurrentTest, currentTest, str.HTMLEncode(msg))))
	end sub
	
	'**********************************************************************************************************
	'* assertFailed 
	'**********************************************************************************************************
	private sub assertFailed(expected, actual, msg)
		assertsFailed = assertsFailed + 1
		println(str.format("Assert [{0}] in [{1}] expected '{2}' but was '{3}' ({4})", _
			array(assertsInCurrentTest, currentTest, str.HTMLEncode(toString(expected)), str.HTMLEncode(toString(actual)), str.HTMLEncode(msg))))
	end sub
	
	'**********************************************************************************************************
	'* toString 
	'**********************************************************************************************************
	private function toString(val)
		on error resume next
		toString = val & ""
		failed = err <> 0
		on error goto 0
		if failed then
			'if its an array we want to show the length of it
			if varType(val) >= 8192 then
				toString = "[Array(" & uBound(val) & ")]"
			else
				toString = "[" & typename(val) & "]"
			end if
		end if
	end function
	
	'**********************************************************************************************************
	'* testFailed 
	'**********************************************************************************************************
	private sub testFailed(errorMsg)
		testsFailed = testsFailed + 1
		println("TEST [" & currentTest & "] FAILED: " & errorMsg)
	end sub
	
	'**********************************************************************************************************
	'* assertStarted 
	'**********************************************************************************************************
	private sub assertStarted()
		assertsInCurrentTest = assertsInCurrentTest + 1
		assertsMade = assertsMade + 1
	end sub
	
	'**********************************************************************************************************
	'* println 
	'**********************************************************************************************************
	private sub println(msg)
		str.writeln(msg & lineBreak)
	end sub

end class
%>
