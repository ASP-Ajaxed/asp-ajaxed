<%
'**************************************************************************************************************
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		StringOperations
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at, Michael Rebec
'' @CREATEDON:		11.12.2003
'' @STATICNAME:		str
'' @CDESCRIPTION:	Collection of various useful string operations. An instance of this class
''					called "str" is created when loading the page. Thus all methods can easily be
''					accessed using str.methodName.
'' @VERSION:		1.1

'**************************************************************************************************************
class StringOperations

	'******************************************************************************************
	'' @SDESCRIPTION:	returns a full URL with a given file and given parameters
	'' @DESCRIPTION:	it encodes automatically the values of the parameters.
	'' @PARAM:			path [string]: the path to the file. e.g. /file.asp, f.asp, http://domain.com/f.asp
	'' @PARAM:			params [string], [array]: the parameters for the url (querystring). if its an array
	''					then every even field is the name and every odd field is the value. if its a string
	''					then its treated as jusst one parameter value for the URL. e.g. file.asp?oneValue
	'' @PARAM:			anker [string]: a jump label. will be appended with # to the end of the URL. empty if no given
	'' @RETURN:			[string] an URL build with the parameters and fully URL-encoded. example: /file.asp?x=10
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
	'' @SDESCRIPTION:	an extension of the common replace() function. basically the same
	''					but you can replace more parts in one go.
	'' @DESCRIPTION:	find (array) -> replaceWith (string): every match of 'find' will be replaced with the string
	''					find (array[n]) -> replaceWith (array[n]): must be same size! every field will be replaced by the field
	''					of replaceWith with the same index.
	''					find (array[n]) -> replaceWith (array[m]): n > m! for the missing fields the last one will be taken
	''					for replacement. e.g. find  = array(1, 2), replaceWith = array(3) means that the 1 will be replaced with
	''					3 and the 2 will also be replaced with 3 because there is no equivalent for the 2
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
	'' @SDESCRIPTION:	checks if a given string can be found in a given array
	'' @DESCRIPTION:	all values in the array are treated as strings when comparing
	'' @PARAM:			aString [string]: the string which should be checked against
	'' @PARAM:			anArray [array]: the array with values where the function will walk through
	'' @PARAM:			caseSensitive [bool]: should the search be case sensitive?
	'' @RETURN:			[int] index of the first found field within the array. -1 if not found
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
	'' @DESCRIPTION:	it ALWAYS returns the type of the alternative
	'' @PARAM:			value [string]: value which should be parsed
	'' @PARAM:			alternative [variant]: alternative value if converting is not possible
	''					- if a float value is needed then use a comma. e.g. 0.0
	'' @RETURN:			[variant] the string parsed into the alternative type or the alternative itself
	'******************************************************************************************************************
	public function parse(value, alternative)
		val = trim(value & "")
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
	'' @SDESCRIPTION:	OBSOLETE! use parse() instead
	'******************************************************************************************************************
	public function toFloat(value, alternative)
		toFload = parse(value, alternative)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	OBSOLETE! use parse() instead
	'******************************************************************************************************************
	public function toInt(value, alternative)
		toInt = parse(value, alternative)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	right-aligns a given value by padding left a given character to a totalsize
	'' @DESCRIPTION:	example: input: 22 -> output: 00022 (padded to total length of 5 with the paddingchar 0)
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
	'' @DESCRIPTION:	example: input: 22 -> output: 22000 (padded to total length of 5 with the paddingchar 0)
	'' @PARAM:			value [string]: the value which should be aligned left
	'' @PARAM:			totalLength [string]: whats the total Length of the result string
	'' @PARAM:			paddingChar [string]: the char which is taken for padding
	'' @RETURN:			[string] left aligned string.
	'******************************************************************************************************************
	public function padRight(value, totalLength, paddingChar)
		padRight = left(value & str.clone(paddingChar, totalLength), totalLength)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	defuses the HTML of  given string. so html code wont be recognized as HTML code by browser
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
	'' @PARAM:			value [string]: the value which should be made "safe"
	'' @RETURN:			[string] safe value. e.g. ' are escaped with '', etc.
	'******************************************************************************************************************
	public function SQLSafe(value)
		SQLSafe = replace(value & "", "--", "")
		SQLSafe = replace(SQLSafe & "", "'", "''")
	end function
	
	'******************************************************************************************************************
	'' @DESCRIPTION: 	Makes a string javascript persistent. Changes special characters, etc.
	'' @DESCRIPTION:	using this function it has to be possible to pass any string which should be 
	''					executed in a javascript later. example: you want to execute the following
	''					onclick="obj.value='" & usrInput & "'. so in this case usrInput needs to be validated that no
	''					javascript error happens when he/she enters e.g. a '.
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
	''					e.g. " will be &quot;
	'' @PARAM:			value [string]: the value you want to encode
	'******************************************************************************************************************
	public function HTMLEncode(value)
		HTMLEncode = server.HTMLEncode(value & "")
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	checks if a given string is a syntactically valid email
	'' @RETURN:			[bool] true if it is valid
	'***********************************************************************************************************
	public function isValidEmail(value)
		isValidEmail = false
		set regEx = new regExp
		
		regEx.pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$" 
		regEx.ignoreCase = true
		retVal = regEx.test(value)
		
		if retVal then isValidEmail = true
		
		set regEx = nothing
		
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns null if the input is empty. 
	'' @PARAM:			value [variant]: the value you are dealing with
	'' @RETURN:			[variant] null if the value is empty, otherwise the value itself
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
	'' @DESCRIPTION:	Tags are defined as string-parts surrounded by a < and a >. example: <sample>
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
	'' @DESCRIPTION:	string "HAXN" will result (partitionlength=2) in the following array:
	''					a(0) = "HA"; a(1) = "XN"
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
	'' @SDESCRIPTION:	shortens a string and adds a custom string at the end if string is longer than a given value.
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
	'' @PARAM:			returnIndex [int]: what index should be returned after spliting?. -1 = get the last index
	'' @RETURN:			[string] string content for the wanted field of the array
	'******************************************************************************************************************
	public function splitValue(stringToSplit, delimiter, returnIndex)
		tmpArray = split(stringToSplit, delimiter)
		if returnIndex = -1 then returnIndex = uBound(tmpArray)
		splitValue = tmpArray(returnIndex)
		tmpArray = null
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	adds a slash "/" at the end of the string if there isnt one.
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
	'' @SDESCRIPTION:	Converts a string to a "char" array
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
	'' @SDESCRIPTION:	Converts an array to a string
	'' @PARAM:			arr [Array]: the array you want to concat
	'' @PARAM:			seperator [string]: seperator between the array-fields
	'' @RETURN:			[string] concated array
	'**************************************************************************************************************
	public function arrayToString(byVal arr, seperator)
		for i = 0 to uBound(arr)
			if i > 0 then
				strObj = strObj & seperator
			end if
			strObj = strObj & arr(i)
		next
		arrayToString = strObj
	end function
	
	'******************************************************************************************
	'' @SDESCRIPTION:	Converts a part of a multidimensional array to a string
	'' @PARAM:			arr [Array]: the array you want to concat
	'' @PARAM:			seperator [string]: seperator between the array-fields
	'' @PARAM:			dimension [int]: the dimension index in the array -> e.g. you have an
	''					array of the size (5, 2) -> with dimension 2 you get the array fields
	''					(0, 2), (1, 2), ..., (4, 2) as the string object, savvy ?
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
	'' @SDESCRIPTION:	Checks if the string begins with ...
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			chars [string]: the compare char/string
	'' @RETURN:			[bool] true if the two strings are equal
	'**************************************************************************************************************
	public function startsWith(byVal str, chars)
		startsWith = (left(str, len(chars)) = chars)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if the string ends with ...
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			chars [string]: the compare char/string
	'' @RETURN:			[bool] true if the two strings are equal
	'**************************************************************************************************************
	public function endsWith(byVal str, chars)
		endsWith = (right(str, len(chars)) = chars)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Concats a string "n" times
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			n [int]: the number of concats
	'' @RETURN:			[string] the concated string
	'**************************************************************************************************************
	public function clone(byVal str, n)
		for i = 1 to n : clone = clone & str : next
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Removes characters from the begin of a string
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			n [int]: the number of characters you want to remove
	'' @RETURN:			[string] the trimmed string
	'**************************************************************************************************************
	public function trimStart(byVal str, n)
		value = mid(str, n + 1)
		trimStart = value
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Removes characters from the end of a string
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
	'' @PARAM:			stringToTrim [string]: the "find" string which will be trimmed out of the string
	'' @RETURN:			[string] the trimmed string
	'**************************************************************************************************************
	public function trimString(byVal str, stringToTrim)
		trimString = replace(str, stringToTrim, "", 1, -1)
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if a char is an alphabetic character or not. A-Z or a-z
	'' @PARAM:			character [char]: The char to check
	'' @RETURN:			[bool] Wether the char is alphabetic or not.
	'**************************************************************************************************************
	public function isAlphabetic(byVal character)
		asciiValue = cint(asc(character))
		isAlphabetic = ((65 <= asciiValue and asciiValue <= 90) or (97 <= asciiValue and asciiValue <= 122))
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Swaps the Case of a String. Lower to Upper and vice a versa
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
	'' @SDESCRIPTION:	Makes every first letter of every word to upper-case.
	'' @DESCRIPTION:	Good for Names or cities. Example: "axel schweis" will result in "Axel Schweis"
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
	'' @SDESCRIPTION:	Switches the case of the letters alternatly to U/Lcase - just for fun :)
	'' @DESCRIPTION:	So "testing" will result in "TeStInG"
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
	'' @SDESCRIPTION:	Replaces all {n} values in a given string through the n-index field of an array.
	'' @DESCRIPTION:	Its just like the string.format method in .NET. so if you provide "my Name is {0}" as
	''					your input then the {0} will be replaced by the first field of your array. and so on.
	'' @PARAM:			str [string]: the source string
	'' @PARAM:			arr [array]: the array with your values
	'' @RETURN:			[string] changed string
	'**************************************************************************************************************
	public function format(byVal str, arr)
		for i = 0 to ubound(arr)
			str = replace(str, "{" & i & "}", cstr(arr(i)))
		next
		format = str
	end function
	
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
