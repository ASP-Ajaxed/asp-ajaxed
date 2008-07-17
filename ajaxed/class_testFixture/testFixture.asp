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
''					- Tests must be subs which are named <em>test_1()</em>, <em>test_2()</em>, etc. 
''					- Call the different assert methods within your tests
''					- if you need to debug your failures then turn on the <em>debug</em> property
''					- create a <em>setup()</em> sub if you need a procedure which will be called before every test
''					- run the fixture with the <em>run()</em> method
''					Example of a simple test (put this in an own file):
''					<code>
''					<!--#include virtual="/ajaxed/class_TestFixture/testFixture.asp"-->
''					<%
''					set tf = new TestFixture
''					tf.run()
''					
''					sub test_1()
''					.	tf.assert 1 = 1, "1 is not equal 2"
''					.	'Lets test if our home page works
''					.	tf.assertResponse "/default.asp", empty, "<h1>Welcome</h1>", "Welcome page seems not to work"
''					end sub
''					% >
''					</code>
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class TestFixture

	'private members
	private currentTest, assertsInCurrentTest, testsMade, testsFailed, assertsMade, assertsFailed
	
	'public members
	public debug			''[bool] Turn this on to debug your tests. error handling is turned off then. default = FALSE
	public lineBreak		''[string] The line break which should be used between each message. default = <em>&lt;br&gt;</em>
	public requestTimeout	''[int] Timout for page requests. e.g. when using <em>asserResponse()</em> default = <em>3</em>
	
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
	'' @SDESCRIPTION: 	Runs all defined tests of the test fixture
	'' @DESCRIPTION:	Executes all subs starting with <em>test_</em> followed by a number. Numbers must start with <em>1</em>.
	''					The execution will stop if no test with the next number is found. Thus is you define <em>test_1()</em>, <em>test_2()</em> and <em>test_4()</em> tests then the 4th test wont run because there is no 3rd test.
	''					- if there is a <em>setup()</em> sub then its called before every test
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
	'' @SDESCRIPTION: 	Checks if a given url contains a given response (defined with a regex pattern)
	'' @DESCRIPTION:	- will fail if the <em>url</em> cannot be reached
	''					- the regular expression will run <strong>not</strong> case sensitive
	''					- if <em>debug</em> is on then the whole response will be shown if the assert fails
	''					- assert will fail if the page does not return SUCCESS status code (<em>200</em>)
	''					<code>
	''					<%
	''					'checks if the default.asp contains a <h1> tag
	''					assertResponse "/default.asp", empty, "<h1>.*</h1>", "Default.asp seem not to work"
	''					
	''					'checks if the login.asp contains a <div> tag with a css class "error" when being posted
	''					assertResponse array("POST", "/login.asp"), empty, "<div class=""error"">", "Using login without any credentials should return an error"
	''					% >
	''					</code>
	'' @PARAM:			url [string], [array]: url to request.
	''					- if a STRING then it will be requested with <em>GET</em> method
	''					- if an ARRAY then first field is the desired method and the 2nd field the actual url e.g. <em>array(method, url)</em>
	''					- only full (<em>http://...</em>) or virtual (starting with <em>/</em>) urls are allowed
	'' @PARAM:			params [array]: parameters for the request. Even fields of the array hold the names and the odd fields hold the corresponding values.
	''					- if <em>POST</em> request then send as <em>POST</em> values (if querystring values needed then add them direclty to the URL).
	''					- if <em>GET</em> request then send via querystring. 
	''					- provide EMPTY if no parameters are needed
	'' @PARAM:			pattern [string]: regex pattern which will be checked against after <em>url</em> has been fetched
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
	'' @DESCRIPTION:	Will fail if the file does not exist at all. <strong>Case sensitive</strong>.
	'' @PARAM:			virtualFilePath [string]: virtual path to the file e.g. /test.txt
	'' @PARAM:			stringToFind [string]: The string which you expect to be in the file
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
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
	'' @SDESCRIPTION: 	Expects a string NOT to exist in a given file.
	'' @DESCRIPTION:	Will succeed if the file does not exists. <strong>Case sensitive</strong>
	'' @PARAM:			virtualFilePath [string]: virtual path to the file e.g. /test.txt
	'' @PARAM:			stringToFind [string]: The string which you expect NOT to be in the file
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
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
	'' @SDESCRIPTION: 	Expects a value to be TRUE
	'' @PARAM:			truth [bool]: the boolean expression which should be TRUE in order to pass the test
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assert(truth, msg)
		assertStarted()
		if not truth then assertFailed true, truth, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects a value to be FALSE
	'' @PARAM:			expected [bool]: the boolean expression which should be FALSE in order to pass the test
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertNot(expected, msg)
		assertStarted()
		if expected then assertFailed false, expected, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects two values to be equal
	'' @DESCRIPTION:	You can also compare more values by providing arrays as parameters. All values are compared against each other and must be equal and on the same position.
	''					<code>
	''					<%
	''					'this will fail
	''					assertResponse 1, 2, "values are not equal"
	''					'array equality (will pass)
	''					assertResponse array(1, 2), array(1, 2), "arrays are not equal"
	''					'array equality (both will fail)
	''					assertResponse array(1, 2), array(1, 3), "arrays are not equal"
	''					assertResponse array(1, 2), array(1), "arrays are not equal"
	''					% >
	''					</code>
	'' @PARAM:			expected [variant], [array]: The value(s) you expect. 
	'' @PARAM:			actual [variant], [array]: The actal value(s).
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
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
	'' @SDESCRIPTION: 	Expects a value to be contained within a given data structure
	'' @PARAM:			data [array]: Must be an ARRAY. if it is not an ARRAY then the assert will fail.
	'' @PARAM:			expected [variant]: The value you expects to exist within the <em>data</em>
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertHas(data, expected, msg)
		assertStarted()
		set d = (new DataContainer)(data)
		if d is nothing then assertFailed expected, data, msg
		if not d.contains(expected) then assertFailed expected, data, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects a value NOT to be contained within a given data structure
	'' @PARAM:			data [array]: Must be an ARRAY. If data is not an ARRAY then the assert will succeed (because its not in it)
	'' @PARAM:			expected [variant]: The value you expect NOT be in within the <em>data</em>
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertHasNot(data, expected, msg)
		assertStarted()
		set d = (new DataContainer)(data)
		if d is nothing then exit sub
		if d.contains(expected) then assertFailed expected, data, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects two values not to be equal
	'' @DESCRIPTION:	You can also compare more values by providing arrays as parameters. All values are compared against each other and must be equal and on the same position.
	'' @PARAM:			expected [variant], [array]: The value(s) you expect. 
	'' @PARAM:			actual [variant], [array]: The actal value(s).
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertNotEqual(expected, actual, msg)
		assertStarted()
		if isEqual(expected, actual) then assertFailed expected, actual, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects a value to be of a given type (class)
	'' @PARAM:			expectedClassName [string]: expected class name e.g. <em>User</em>
	'' @PARAM:			value [variant]: The value which type will be checked.
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertInstanceOf(expectedClassName, value, msg)
		assertStarted()
		if lCase(typename(value)) <> lCase(expectedClassName) then assertFailed expectedClassName, typename(value), msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Expects a value to match a given regular expression pattern
	'' @DESCRIPTION:	It uses <em>str.matching()</em> internally.
	'' @PARAM:			pattern [string]: The regular expression pattern which will be used for matching.
	'' @PARAM:			value [string]: The value which wil be checked against the <em>pattern</em>
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertMatch(pattern, value, msg)
		assertStarted()
		if not str.matching(value, pattern, true) then assertFailed pattern, value, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	expects a value to be nothing (<em>object</em>)
	'' @PARAM:			value [variant]: Value which will be checked
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
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
	'' @SDESCRIPTION: 	Expects two values to be equal with a given delta
	'' @DESCRIPTION:	<code>
	''					<%
	''					'these will pass
	''					assertInDelta 10, 11, 1, "Something is wrong"
	''					assertInDelta 5.4, 5.4, 0.1, "Something is wrong"
	''					assertInDelta 5.4, 5.3, 0.1, "Something is wrong"
	''					'those will fail
	''					assertInDelta 4.4, 4.5, 0.1, "Something is wrong"
	''					assertInDelta 33, 12, 10, "Something is wrong"
	''					% >
	''					</code>
	'' @PARAM:			expected [int], [float]: The expected value
	'' @PARAM:			actual [int], [float]: The actual value
	'' @PARAM:			delta [int], [float]: Delta which represents the tolerance for the comparison of <em>expected</em> and <em>actual</em>.
	'' @PARAM:			msg [string]: A message which will be shown if the assert fails
	'**********************************************************************************************************
	public sub assertInDelta(expected, actual, delta, msg)
		assertStarted()
		if actual + delta <> expected and actual - delta <> expected then assertFailed expected, actual, msg
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Allows to manually fail an assert
	'' @PARAM:			msg [string]: message which describes the failure
	'**********************************************************************************************************
	public sub fail(msg)
		assertsFailed = assertsFailed + 1
		println(str.format("Assert [{0}] in [{1}] {2}", _
			array(assertsInCurrentTest, currentTest, str(msg))))
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	prints an information message to the output.
	'' @DESCRIPTION:	useful e.g. on testing email. after the test you could inform the user that a specific amount
	''					of emails should be now in his inbox.
	'' @PARAM:			msg [string]: message which contains the information
	'**********************************************************************************************************
	public sub info(msg)
		if str.matching(msg, "\[|\]", true) then lib.throwError("TestFixture.info() does not accept '[' or ']' within the msg argument.")
		println("INFO [" & str(msg) & "]")
	end sub
	
	'**********************************************************************************************************
	'* assertFailed 
	'**********************************************************************************************************
	private sub assertFailed(expected, actual, msg)
		assertsFailed = assertsFailed + 1
		println(str.format("Assert [{0}] in [{1}] expected '{2}' but was '{3}' ({4})", _
			array(assertsInCurrentTest, currentTest, str(toString(expected)), str(toString(actual)), str(msg))))
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
