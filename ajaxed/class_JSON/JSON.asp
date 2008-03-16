<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt										
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		JSON
'' @CREATOR:		Michal Gabrukiewicz (gabru at grafix.at), Michael Rebec
'' @CONTRIBUTORS:	- Cliff Pruitt (opensource at crayoncowboy.com)
''					- Sylvain Lafontaine
'' @CREATEDON:		2007-04-26 12:46
'' @CDESCRIPTION:	Comes up with functionality for JSON (http://json.org) to use within ASP.
'' 					Correct escaping of characters, generating JSON Grammer out of ASP datatypes and structures
'' @REQUIRES:		-
'' @OPTIONEXPLICIT:	yes
'' @VERSION:		1.4.1

'**************************************************************************************************************
class JSON

	'private members
	private output, innerCall
	
	'public members
	public toResponse		''[bool] should generated results be directly written to the response? default = false
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		newGeneration()
		toResponse = false
	end sub
	
	'******************************************************************************************
	'' @SDESCRIPTION:	STATIC! takes a given string and makes it JSON valid
	'' @DESCRIPTION:	all characters which needs to be escaped are beeing replaced by their
	''					unicode representation according to the 
	''					RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
	'' @PARAM:			val [string]: value which should be escaped
	'' @RETURN:			[string] JSON valid string
	'******************************************************************************************
	public function escape(val)
		dim cDoubleQuote, cRevSolidus, cSolidus
		cDoubleQuote = &h22
		cRevSolidus = &h5C
		cSolidus = &h2F
		
		dim i, currentDigit
		for i = 1 to (len(val))
			currentDigit = mid(val, i, 1)
			if asc(currentDigit) > &h00 and asc(currentDigit) < &h1F then
				currentDigit = escapequence(currentDigit)
			elseif asc(currentDigit) >= &hC280 and asc(currentDigit) <= &hC2BF then
				currentDigit = "\u00" + right(padLeft(hex(asc(currentDigit) - &hC200), 2, 0), 2)
			elseif asc(currentDigit) >= &hC380 and asc(currentDigit) <= &hC3BF then
				currentDigit = "\u00" + right(padLeft(hex(asc(currentDigit) - &hC2C0), 2, 0), 2)
			else
				select case asc(currentDigit)
					case cDoubleQuote: currentDigit = escapequence(currentDigit)
					case cRevSolidus: currentDigit = escapequence(currentDigit)
					case cSolidus: currentDigit = escapequence(currentDigit)
				end select
			end if
			escape = escape & currentDigit
		next
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	generates a representation of a name value pair in JSON grammer
	'' @DESCRIPTION:	the generation is done fully recursive so the value can be a complex datatype as well. e.g.
	''					toJSON("n", array(array(), 2, true), false) or toJSON("n", array(RS, dict, false), false), etc.
	'' @PARAM:			name [string]: name of the value (accessible with javascript afterwards). leave empty to get just the value
	'' @PARAM:			val [variant], [int], [float], [array], [object], [dictionary], [recordset]: value which needs
	''					to be generated. Conversation of the data types (ASP datatype -> Javascript datatype):
	''					NOTHING, NULL -> null, ARRAY -> array, BOOL -> bool, OBJECT -> name of the type, 
	''					MULTIDIMENSIONAL ARRAY -> generates a 1 dimensional array (flat) with all values of the multidim array
	''					DICTIONARY -> valuepairs. each key is accessible as property afterwards
	''					RECORDSET -> array where each row of the recordset represents a field in the array. 
	''					fields have properties named after the column names of the recordset (LOWERCASED!)
	''					e.g. generate(RS) can be used afterwards r[0].ID
	''					INT, FLOAT -> number
	''					OBJECT with reflect() method -> returned as object which can be used within JavaScript
	'' @PARAM:			nested [bool]: is the value pair already nested within another? if yes then the {} are left out.
	'' @RETURN:			[string] returns a JSON representation of the given name value pair
	''					(if toResponse is on then the return is written directly to the response and nothing is returned)
	'******************************************************************************************************************
	public function toJSON(name, val, nested)
		if not nested and not isEmpty(name) then write("{")
		if not isEmpty(name) then write("""" & escape(name) & """: ")
		generateValue(val)
		if not nested and not isEmpty(name) then write("}")
		toJSON = output
		
		if innerCall = 0 then newGeneration()
	end function
	
	'******************************************************************************************************************
	'* generate 
	'******************************************************************************************************************
	private function generateValue(val)
		if isNull(val) then
			write("null")
		elseif isArray(val) then
			generateArray(val)
		elseif isObject(val) then
			if val is nothing then
				write("null")
			elseif typename(val) = "Dictionary" then
				generateDictionary(val)
			elseif typename(val) = "Recordset" then
				generateRecordset(val)
			else
				generateObject(val)
			end if
		else
			'bool
			dim varTyp
			varTyp = varType(val)
			if varTyp = 11 then
				if val then write("true") else write("false")
			'int, long, byte
			elseif varTyp = 2 or varTyp = 3 or varTyp = 17 or varTyp = 19 then
				write(cLng(val))
			'single, double, currency
			elseif varTyp = 4 or varTyp = 5 or varTyp = 6 or varTyp = 14 then
				write(replace(cDbl(val), ",", "."))
			else
				write("""" & escape(val & "") & """")
			end if
		end if
		generateValue = output
	end function
	
	'******************************************************************************************************************
	'* generateArray 
	'******************************************************************************************************************
	private sub generateArray(val)
		dim item, i
		write("[")
		i = 0
		'the for each allows us to support also multi dimensional arrays
		for each item in val
			if i > 0 then write(",")
			generateValue(item)
			i = i + 1
		next
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateDictionary 
	'******************************************************************************************************************
	private sub generateDictionary(val)
		dim keys, i
		innerCall = innerCall + 1
		write("{")
		keys = val.keys
		for i = 0 to uBound(keys)
			if i > 0 then write(",")
			toJSON keys(i), val(keys(i)), true
		next
		write("}")
		innerCall = innerCall - 1
	end sub
	
	'******************************************************************************************************************
	'* generateRecordset 
	'******************************************************************************************************************
	private sub generateRecordset(val)
		dim i
		write("[")
		while not val.eof
			innerCall = innerCall + 1
			write("{")
			for i = 0 to val.fields.count - 1
				if i > 0 then write(",")
				toJSON lCase(val.fields(i).name), val.fields(i).value, true
			next
			write("}")
			val.movenext()
			if not val.eof then write(",")
			innerCall = innerCall - 1
		wend
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateObject 
	'******************************************************************************************************************
	private sub generateObject(val)
		dim props
		on error resume next
		set props = val.reflect()
		if err = 0 then
			on error goto 0
			innerCall = innerCall + 1
			toJSON empty, props, true
			innerCall = innerCall - 1
		else
			on error goto 0
			write("""" & escape(typename(val)) & """")
		end if
	end sub
	
	'******************************************************************************************************************
	'* newGeneration 
	'******************************************************************************************************************
	private sub newGeneration()
		output = empty
		innerCall = 0
	end sub
	
	'******************************************************************************************
	'* JsonEscapeSquence 
	'******************************************************************************************
	private function escapequence(digit)
		escapequence = "\u00" + right(padLeft(hex(asc(digit)), 2, 0), 2)
	end function
	
	'******************************************************************************************
	'* padLeft 
	'******************************************************************************************
	private function padLeft(value, totalLength, paddingChar)
		padLeft = right(clone(paddingChar, totalLength) & value, totalLength)
	end function
	
	'******************************************************************************************
	'* clone 
	'******************************************************************************************
	public function clone(byVal str, n)
		dim i
		for i = 1 to n : clone = clone & str : next
	end function
	
	'******************************************************************************************
	'* write 
	'******************************************************************************************
	private sub write(val)
		if toResponse then
			response.write(val)
		else
			output = output & val
		end if
	end sub

end class
%>