<%
'**************************************************************************************************************
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		StringOperations
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at, Michael Rebec
'' @CREATEDON:		11.12.2003
'' @STATICNAME:		str
'' @CDESCRIPTION:	Collection of various useful string operations. An instance of this class
''					called <em>str</em> is created when loading the page. Thus all methods can easily be
''					accessed using <em>str.methodName</em>
'' @VERSION:		1.1

'**************************************************************************************************************
class StringOperations

	'******************************************************************************************
	'' @SDESCRIPTION:	Takes a string and tries to make it look more humanreadable. Its for creating pretty output
	'' @DESCRIPTION:	- <em>_</em> (underscores) are replaced by a space character (also more underscores e.g. <em>___</em> are replaced by only one space)
	''					- <em>_id</em> and <em>_fk</em> in the back are stripped as well as <em>id_</em> and <em>fk_</em> in the beginning
	''					- everything is lower cased unless the first letter
	'' @PARAM:			val [string]: the value which should be humanized
	'' @RETURN:			[string] a pretty string. e.g. <em>created_on</em> => <em>Created on</em> or <em>fk_category</em> => <em>Category</em>
	'******************************************************************************************
	public function humanize(byVal val)
		humanize = lCase(trim(rReplace(rReplace(val & "", "^fk_|^id_|_fk$|_id$", "", true), "_+", " ", true)))
		humanize = uCase(left(humanize, 1)) & right(humanize, abs(len(humanize) - 1))
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	performs a replace with a regular expression pattern
	'' @DESCRIPTION:	e.g. surrounding all numbers of a string with brackets: <code><% str.rReplace("i am 20 and he is 10", "(\d)", "($1)", true) % ></code>
	'' @PARAM:			val [string]: the value you want to replace the matches
	'' @PARAM:			pattern [string]: regular expression pattern
	'' @PARAM:			replaceWith [string]: a string which is used for the replacement of the matches
	''					- <em>$1, $2</em>, .. can be used as placeholders for grouped expressions of the regex pattern
	'' @PARAM:			ignoreCase [bool]: ignore the case on comparison
	'' @RETURN:			[string] returns the new string with replacements made. if no replacements made then the same string is returned as given on input
	'******************************************************************************************
	public function rReplace(val, pattern, replaceWith, ignoreCase)
		set re = new regexp
		re.ignorecase = ignoreCase
		re.global = true
		re.pattern = pattern & ""
		rReplace = re.replace(val & "", replaceWith)
		set re = nothing
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	checks if a given string is matching a given regular expression pattern
	'' @PARAM:			val [string]: the value which needs to be checked against the pattern
	'' @PARAM:			pattern [string]: regular expression pattern
	'' @PARAM:			ignoreCase [bool]: ignore the case on comparison
	'' @RETURN:			[bool] TRUE if val matches the pattern otherwise false
	'******************************************************************************************
	public function matching(val, pattern, ignoreCase)
		set re = new regexp
		re.ignorecase = ignoreCase
		re.global = true
		re.pattern = pattern & ""
		matching = re.test(val & "")
		set re = nothing
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	returns a full URL with a given file and given parameters
	'' @DESCRIPTION:	it encodes automatically the values of the parameters.
	'' @PARAM:			path [string]: the path to the file. e.g. <em>/file.asp</em>, <em>f.asp</em>, <em>http://domain.com/f.asp</em>
	'' @PARAM:			params [string], [array]: the parameters for the url (querystring). if its an ARRAY
	''					then every even field is the name and every odd field is the value. if its a STRING
	''					then its treated as jusst one parameter value for the URL. e.g. <em>file.asp?oneValue</em>
	'' @PARAM:			anker [string]: a jump label. will be appended with <em>#</em> to the end of the URL. EMPTY if no given
	'' @RETURN:			[string] an URL build with the parameters and fully URL-encoded. example: <em>/file.asp?x=10</em>
	'******************************************************************************************
	public function URL(path, params, anchor)
		URL = path
		if isArray(params) then
			for i = 0 to uBound(params) step 2
				URL = URL & lib.iif(i = 0, "?", "&")
				URL = URL & server.URLEncode(params(i)) & "="
				if (i + 1) <= uBound(params) then URL = URL & server.URLEncode(params(i + 1))
			next
		else
			URL = URL & "?" & server.URLEncode(params)
		end if
		if not isEmpty(anchor) then URL = URL & "#" & server.URLEncode(anchor)
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	an extension of the common <em>replace()</em> function. basically the same
	''					but you can replace more parts in one go.
	'' @DESCRIPTION:	- <em>find (array)</em> -> <em>replaceWith (string)</em>: every match of 'find' will be replaced with the string
	''					- <em>find (array[n])</em> -> <em>replaceWith (array[n])</em>: must be same size! every field will be replaced by the field of replaceWith with the same index.
	''					- <em>find (array[n])</em> -> <em>replaceWith (array[m])</em>: n > m! for the missing fields the last one will be taken for replacement. e.g. <code><% str.change("somestring", array(1, 2), array(3)) % ></code> means that the 1 will be replaced with 3 and the 2 will also be replaced with 3 because there is no equivalent for the 2
	'' @PARAM:			source [string]: the source string in which we search
	'' @PARAM:			find [string], [array]: what you are looking for. 
	'' @PARAM:			replaceWith [string], [array]: the value you want to replace the matches with.
	'' @RETURN:			[string] the replaced source string
	'******************************************************************************************
	public function change(byVal source, find, replaceWith)
		change = source
		if isArray(find) then
			for i = 0 to uBound(find)
				if isArray(replaceWith) then
					if uBound(replaceWith) < i then
						toReplace = replaceWith(uBound(replaceWith))
					else
						toReplace = replaceWith(i)
					end if
				else
					toReplace = replaceWith
				end if
				change = replace(change, find(i), toReplace)
			next
		elseif not isArray(find) then
			change = replace(change, find, replaceWith)
		else
			lib.error("(str.replace) One of the arguments is of the wrong type.")
		end if
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	checks if a given string can be found in a given ARRAY
	'' @DESCRIPTION:	all values in the ARRAY are treated as strings when comparing
	'' @PARAM:			aString [string]: the string which should be checked against
	'' @PARAM:			anArray [array]: the array with values where the function will walk through
	'' @PARAM:			caseSensitive [bool]: should the search be case sensitive?
	'' @RETURN:			[int] index of the first found field within the ARRAY. <em>-1</em> if not found
	'******************************************************************************************
	public function existsIn(byVal aString, byVal anArray, caseSensitive)
		aString = lib.iif(caseSensitive, aString & "", lCase(aString & ""))
		for existsIn = 0 to uBound(anArray)
			current = lib.iif(caseSensitive, anArray(existsIn) & "", lCase(anArray(existsIn) & ""))
			if aString = current then exit function
		next
		existsIn = -1
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	tries to parse a given value into the datatype of the alternative. If it cannot be parsed
	''					then the alternative is passed through
	'' @DESCRIPTION:	It ALWAYS returns the type of the alternative. Recommended usage when getting values from querystring or form:
	''					<code>
	''					'gets the ID value from querystring and ensures its an integer.
	''					'if it cannot be parsed then always 0 is returned
	''					id = str.parse(page.QS("id"), 0)
	''					'the same with a form value
	''					id = str.parse(page.RF("id"), 0)
	''					'parsing to a float number
	''					nr = str.parse("212.22", 0.0)
	''					</code>
	'' @PARAM:			value [string]: value which should be parsed.
	''					Float values are only recognized with the comma which is common for the current locale (check <em>local.comma</em>.
	''					<code>
	''					<%
	''					'e.g. US
	''					str.parse("12.3", 0.0) ' => 12.3
	''					str.parse("12,3", 0.0) ' => 123
	''					'e.g. Germany
	''					str.parse("12.3", 0.0) ' => 123
	''					str.parse("12,3", 0.0) ' => 12.3
	''					% >
	''					</code>
	'' @PARAM:			alternative [variant]: alternative value if converting is not possible
	''					- if a FLOAT value is needed then use a comma. e.g. <em>0.0</em>
	'' @RETURN:			[variant] the string parsed into the alternative type or the alternative itself
	'******************************************************************************************************************
	public function parse(value, alternative)
		val = trim(cstr(value) & "")
		parse = alternative
		if val = "" then exit function
		on error resume next
		select case varType(parse)
			case 2, 3 'integer, long
				parse = cLng(val)
			case 4, 5 'single, double
				parse = cdbl(val)
			case 7 'date
				parse = cDate(val)
			case 11 'bool
				parse = cBool(val)
			case 8 'string
				parse = value & ""
			case else
				on error goto 0
				lib.throwError("Type not supported for string parsing.")
		end select
		on error goto 0
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use <em>parse()</em> instead
	'******************************************************************************************************************
	public function toFloat(value, alternative)
		toFload = parse(value, alternative)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use <em>parse()</em> instead
	'******************************************************************************************************************
	public function toInt(value, alternative)
		toInt = parse(value, alternative)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	right-aligns a given value by padding left a given character to a totalsize
	'' @DESCRIPTION:	example: input: <em>22</em> -> output: <em>00022</em> (padded to total length of 5 with the paddingchar 0)
	'' @PARAM:			value [string]: the value which should be aligned right
	'' @PARAM:			totalLength [string]: whats the total Length of the result string
	'' @PARAM:			paddingChar [string]: the char which is taken for padding
	'' @RETURN:			[string] right aligned string.
	'******************************************************************************************************************
	public function padLeft(value, totalLength, paddingChar)
		padLeft = right(str.clone(paddingChar, totalLength) & value, totalLength)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	left-aligns a given value by padding right a given character to a totalsize
	'' @DESCRIPTION:	example: input: <em>22</em> -> output: <em>22000</em> (padded to total length of 5 with the paddingchar 0)
	'' @PARAM:			value [string]: the value which should be aligned left
	'' @PARAM:			totalLength [string]: whats the total Length of the result string
	'' @PARAM:			paddingChar [string]: the char which is taken for padding
	'' @RETURN:			[string] left aligned string.
	'******************************************************************************************************************
	public function padRight(value, totalLength, paddingChar)
		padRight = left(value & str.clone(paddingChar, totalLength), totalLength)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	defuses the HTML of given string. so html code wont be recognized as HTML code by browser
	'' @PARAM:			value [string]: the value which should be defused
	'' @RETURN:			[string] defused value
	'******************************************************************************************************************
	public function defuseHTML(value)
		'does not work properly ... maybe regex.
		defuseHTML = replace(value & "", "<", "<code>&lt;</code>")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	makes a given string safe for the use within sql statements
	'' @DESCRIPTION:	e.g. if its necessary to pass through an user input directly into a sql-query
	''					example on a basic login scenario: <code><% sql = "SELECT * FROM user WHERE login = " & str.SQLSafe(username) % ></code>
	'' @PARAM:			value [string]: the value which should be made "safe"
	'' @RETURN:			[string] safe value. e.g. <em>'</em> are escaped with <em>''</em>, etc.
	'******************************************************************************************************************
	public function SQLSafe(value)
		SQLSafe = replace(value & "", "--", "")
		SQLSafe = replace(SQLSafe & "", "'", "''")
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Makes a string javascript persistent. Changes special characters, etc.
	'' @DESCRIPTION:	using this function it has to be possible to pass any string which should be 
	''					executed in a javascript later. example: you want to execute the following
	''					<code>
	''					<button onclick="this.value='" & str.JSEncode(usrInput) & "'"></button>
	''					</code>
	'' @PARAM:			val [string]: the value which needs to be encoded
	'' @RETURN:			[string] encoded string which can be used within javascript strings.
	'******************************************************************************************************************
	public function JSEncode(byVal val)
		val = val & ""
		tmp = replace(tmp, chr(92), "\\")
		tmp = replace(val, chr(39), "\'")
		tmp = replace(tmp, chr(34), "&quot;")
		tmp = replace(tmp, chr(13), "<br>")
		tmp = replace(tmp, chr(10), " ")
		JSEncode = tmp
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	generates a hidden input field and returns the HTML for it
	'' @DESCRITPION:	its recommended to use this function. e.g. value is HTMLEncoded.
	'' @PARAM:			name [string]: the name of the value
	'' @PARAM:			value [string]: the value it should hold
	'******************************************************************************************************************
	public function getHiddenInput(name, value)
		getHiddenInput = "<input type=""Hidden"" name=""" & name & """ value=""" & HTMLEncode(value) & """>"
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a value to the output and stops the response.
	'' @DESCRIPTION:	The response will be stopped after writing the value. good for debuging.
	'' @PARAM:			value [string]: the value you want to write
	'******************************************************************************************************************
	public function writeEnd(value)
		write(value)
		me.end()
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	encodes a value so that any special chars will be transformed into html entities.
	''					e.g. <em>"</em> will converted to <em>&quot;</em>
	'' @DESCRIPTION:	should be used in all views to prevent XSS. for that reason its also the default function so its
	''					easier to use. example of usage in your views (as its the default function it can be used directly on the instance): <code>str("some value <>")</code>
	'' @PARAM:			value [string]: the value you want to encode
	'******************************************************************************************************************
	public default function HTMLEncode(value)
		HTMLEncode = server.HTMLEncode(value & "")
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	checks if a given string is a syntactically correct email address
	'' @PARAM:			val [string]: the value to check
	'' @RETURN:			[bool] true if string seems to be an email
	'******************************************************************************************
	public function isValidEmail(val)
		isValidEmail = matching(val, "^[A-Z0-9._%-]+@[A-Z0-9._%-]+\.[A-Z]{2,4}$", true)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns null if the input is empty. 
	'' @PARAM:			value [variant]: the value you are dealing with
	'' @RETURN:			[variant] NULL if the value is EMPTY, otherwise the value itself
	'******************************************************************************************************************
	public function nullIfEmpty(value)
		nullIfEmpty = value
		if trim(value & "") = "" then nullIfEmpty = null
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	removes all non-printable chars from the end and the beginning of the string
	'' @DESCRIPTION:	spaces, returns, line-feeds, tabs, etc. will be removed
	'' @PARAM:			inputStr [string]: string to be trimmed
	'' @RETURN:			[string] trimmed string
	'******************************************************************************************************************
	public function trimComplete(inputStr)
		stringLen = len(inputStr)
		
		if stringLen > 0 then
			'trim-left
			for i = 1 to stringLen
				if asc(mid(inputStr, i, 1)) > 32 then exit for
			next
			
			trimComplete = mid(inputStr, i)
			stringLen = len(trimComplete)
			
			'trim-right
			for i = stringLen to 1 Step - 1
				if asc(mid(trimComplete, i, 1)) > 32 then exit for
			next
			trimComplete = left(trimComplete, i)
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	removes all Tags from a given string
	'' @DESCRIPTION:	Tags are defined as string-parts surrounded by a <em><</em> and a <em>></em>. example: <em>&lt;sample&gt;</em>
	'' @PARAM:			inputStr [string]: string where the Tags should be removed
	'' @RETURN:			[string] the input-String without any Tags
	'******************************************************************************************************************
	public function stripTags(inputStr)
		set regEX = new RegExp
		regEX.pattern = "<[^>]*>"
		regEX.global = true
		stripTags = regEX.replace(inputStr & "", "")
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	divides a string into several string-paritions of a given length
	'' @DESCRIPTION:	string <em>HAXN</em> will result (<em>partitionlength = 2</em>) in the following array:
	''					<em>array("HA", "XN")</em>
	'' @PARAM:			inputStr [string]: string which should be divided
	'' @PARAM:			partitionLength [string]: the length of each partion
	'' @RETURN:			[array] array with all partitions
	'******************************************************************************************************************
	public function divide(inputStr, byVal partitionLength)
		if partitionLength < 1 then exit function
		i = 0
		tmpArray = array()
		while (i * partitionLength) < len(inputStr)
			redim preserve tmpArray(i)
			tmpArray(i) = mid(inputStr, (i * partitionLength) + 1, partitionLength)
			i = i + 1
		wend
		divide = tmpArray
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	stops to response
	'******************************************************************************************************************
	public sub [end]()
		response.end
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Shortens a string and adds a custom string at the end if string is longer than a given value.
	'' @DESCRIPTION:	Useful if you want to display excerpts:
	''					<code><%= str.shorten("some value", 10, "...") % ></code>
	'' @PARAM:			str [string]: string which should be checked against cutting
	'' @PARAM:			maxChars [string]: whats the maximum allowed length of chars
	'' @PARAM:			overflowString [int]: what string should be added at the end of the string if it has been cutted
	'' @RETURN:			[string] cutted string
	'******************************************************************************************************************
	public function shorten(byVal str, maxChars, overflowString)
		str = str & ""
		if len(str) > maxChars then str = left(str, maxChars) & overflowString
		shorten = str
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	splits a string and returns a specified field of the array
	'' @DESCRIPTION:	it uses the split function but immediately returns you the field you want.
	'' @PARAM:			stringToSplit [string]: the string you want to split
	'' @PARAM:			delimiter [string]: whats the delimiter for splitting
	'' @PARAM:			returnIndex [int]: what index should be returned after spliting?. <em>-1</em> = get the last index
	'' @RETURN:			[string] string content for the needed field of the ARRAY
	'******************************************************************************************************************
	public function splitValue(stringToSplit, delimiter, returnIndex)
		tmpArray = split(stringToSplit, delimiter)
		if returnIndex = -1 then returnIndex = uBound(tmpArray)
		splitValue = tmpArray(returnIndex)
		tmpArray = null
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	adds a slash <em>/</em> at the end of the string if there isnt one.
	'' @DESCRIPTION:	Good use for urls or paths if you want to be sure that there is a slash at the end
	''					of an url or a path. 
	'' @PARAM:			pathString [string]: The string (url, path) you want to check
	'' @RETURN:			[string] the new url. with a slash at the end.
	'******************************************************************************************************************
	public function ensureSlash(pathString)
		ensureSlash = pathString
		if not me.endsWith(pathString, "/") then ensureSlash = pathString & "/"
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function write(value)
		response.write(value)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a line to the output
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function writeln(value)
		response.write(value & vbcrlf)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Converts a STRING to a CHAR ARRAY
	'' @PARAM:			str [string]: the string you want to convert
	'' @RETURN:			[array] a "char" array
	'**************************************************************************************************************
	public function toCharArray(byVal str)
		redim charArray(len(str) - 1)
		for i = 0 to uBound(charArray)
			charArray(i) = mid(str, i + 1, 1)
		next
		toCharArray = charArray
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! Converts an ARRAY to a STRING. Use native <em>join()</em> function instead
	'' @DESCRIPTION:	<code><% str.arrayToString(array(1, 2, 3), ",") '=> 1,2,3 % ></code>
	'' @PARAM:			arr [Array]: the array you want to concat
	'' @PARAM:			seperator [string]: seperator between the array fields e.g. <em>,</em>
	'' @RETURN:			[string] concated array
	'**************************************************************************************************************
	public function arrayToString(byVal arr, seperator)
		arrayToString = join(arr, seperator)
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	Converts a part of a multidimensional ARRAY to a STRING
	'' @PARAM:			arr [Array]: the array you want to concat
	'' @PARAM:			seperator [string]: seperator between the array-fields
	'' @PARAM:			dimension [int]: the dimension index in the array e.g. you have an
	''					array of the size <em>(5, 2)</em> -> with <em>dimension</em> <em>2</em> you get the array fields
	''					<em>(0, 2), (1, 2)</em>, ..., <em>(4, 2)</em> as a string.
	'' @RETURN:			[string] concated array
	'******************************************************************************************
	public function multiArrayToString(byVal arr, seperator, dimension)
		for i = 0 to uBound(arr)
			if i > 0 then
				strObj = strObj & seperator
			end if
			strObj = strObj & arr(i, dimension)
		next
		multiArrayToString = strObj
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if the string begins with a given string. Case sensitive.
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			chars [string]: the compare char/string
	'' @RETURN:			[bool] true if source string starts with the <em>chars</em> string
	'**************************************************************************************************************
	public function startsWith(byVal str, chars)
		startsWith = (left(str, len(chars)) = chars)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if the string ends with a given string. Case sensitive.
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			chars [string]: the compare char/string
	'' @RETURN:			[bool] true if the source string ends with the <em>chars</em> string
	'**************************************************************************************************************
	public function endsWith(byVal str, chars)
		endsWith = (right(str, len(chars)) = chars)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Concats a string <em>n</em> times
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			n [int]: the number of concats
	'' @RETURN:			[string] the concated string
	'**************************************************************************************************************
	public function clone(byVal str, n)
		for i = 1 to n : clone = clone & str : next
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Removes a given amount of characters from the begining of a string
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			n [int]: the number of characters you want to remove
	'' @RETURN:			[string] the trimmed string
	'**************************************************************************************************************
	public function trimStart(byVal str, n)
		value = mid(str, n + 1)
		trimStart = value
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Removes a given amount of characters from the end of a string
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			n [int]: the number of characters you want to remove
	'' @RETURN:			[string] the trimmed string
	'**************************************************************************************************************
	public function trimEnd(byVal str, n)
		value = left(str, len(str) - n)
		trimEnd = value
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Removes characters from a given string with all occurencies of a matching string
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			stringToTrim [string]: the string which will be trimmed out of the string
	'' @RETURN:			[string] the trimmed string
	'**************************************************************************************************************
	public function trimString(byVal str, stringToTrim)
		trimString = replace(str, stringToTrim, "", 1, -1)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if given a char is an alphabetic character (<em>A-Z</em> or <em>a-z</em>) or not.
	'' @PARAM:			character [char]: The char to check
	'' @RETURN:			[bool] TRUE the char is alphabetic otherwise FALSE.
	'**************************************************************************************************************
	public function isAlphabetic(byVal character)
		isAlphabetic = false
		if character = "" then exit function
		asciiValue = cint(asc(character))
		isAlphabetic = ((65 <= asciiValue and asciiValue <= 90) or (97 <= asciiValue and asciiValue <= 122))
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Swaps the Case of a String. lowercase chars to uppercase and vice a versa
	'' @PARAM:			str [string]: the source string
	'' @RETURN:			[string] the swapped string
	'**************************************************************************************************************
	public function swapCase(str)
		for i = 1 to len(str)
			current = mid(str, i, 1)
			if isAlphabetic(current) then
				high = asc(ucase(current))
				low = asc(lcase(current))
				sum = high + low
				return = return & chr(sum - asc(current))
			else
				return = return & current
			end if
		next
		swapCase = return
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Converts every first letter of every word into uppercase.
	'' @DESCRIPTION:	Good for Names or cities. Example: <em>jack johnson</em> will result in <em>Jack Johnson</em>
	'' @PARAM:			str [string]: the source string
	'' @RETURN:			[string] changed string
	'**************************************************************************************************************
	public function capitalize(inputStr)
		words = split(inputStr, " ")
		for i = 0 to ubound(words)
			if not i = 0 then tmp = " "
			if len(words(i)) > 0 then tmp = tmp & ucase(left(words(i), 1)) & right(words(i), len(words(i)) - 1)
			words(i) = tmp
		next
		capitalize = arrayToString(words, empty)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Switches the case of the letters alternatly to upper/lowercase - just for fun :)
	'' @DESCRIPTION:	So e.g the word <em>testing</em> will result in <em>TeStInG</em>
	'' @PARAM:			str [string]: the source string
	'' @RETURN:			[string] changed string
	'**************************************************************************************************************
	function leetIt(byVal str)
		for i = 1 to len(str)
			if i mod 2 = 0 then
				leetIt = leetIt & lCase(Mid(str, i, 1))
			else
				leetIt = leetIt & uCase(Mid(str, i, 1))
			end if
		next
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Replaces all <em>{n}</em> placeholders in a given string through the n-index field of a given array.
	'' @DESCRIPTION:	Its just like the <em>string.format</em> method in .NET. So if you have a string <em>my Name is {0}</em> as
	''					your input then the <em>{0}</em> will be replaced by the first field of your array. And so on and so forth.
	''					- Placeholders can be used multiple times within the source string
	''					- In case you get an error about unsupported character sequence, then your string contains a special character sequence which is used internally for escaping (the sequence consists of non printable characters). Thus it cannot be used as input.
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			values [array], [string]: your values which should replace the placeholders. Use a STRING if you have only one placeholder to replace
	'' @RETURN:			[string] changed string
	'**************************************************************************************************************
	public function format(byVal str, byVal values)
		escapeChar = asc(24) & asc(27)
		if instr(str, escapeChar) > 0 then lib.throwError("input string of StringOperations.format() contains unsupported character sequence.")
		if not isArray(values) then values = array(values)
		for i = 0 to ubound(values)
			val = cStr(values(i))
			if instr(val, escapeChar) > 0 then lib.throwError("at least one value given for StringOperations.format() contains unsupported character sequence.")
			val = replace(val, "{", escapeChar)
			str = replace(str, "{" & i & "}", val)
		next
		format = replace(str, escapeChar, "{")
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Writes a formatted string. (combination of <em>write()</em> and <em>format()</em>)
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			values [array], [string]: your values which should replace the placeholders. Use a STRING if you have only one placeholder to replace. Check <em>format()</em> function for more details
	'**************************************************************************************************************
	public sub writef(byVal str, byVal values)
		write(format(str, values))
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Encodes a string into ASCII Characters
	'' @DESCRIPTION:	This function takes a string (for example an email-address) and converts it to
	''					standardized ASCII Character codes, thus blocking bots/spiders from reading your
	''					email address, while yet allowing visitors to continue to read and use your proper
	''					address via mailto links, etc.
	'' @PARAM:			str [string]: the string you want to convert
	'' @RETURN:			[string] the ASCII Encoded String
	'' @CREDIT:			Teco from Planet Source Code
	'**************************************************************************************************************
	public function ASCIICode(byval str)
		ASCIICode = ""
		for i = 1 to len(str)
			ASCIICode = ASCIICode & "&#" & asc(mid(str, i, 1)) & ";"
		next
    end function

end class
%>
