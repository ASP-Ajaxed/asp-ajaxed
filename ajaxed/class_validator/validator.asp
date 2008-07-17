<%
'**************************************************************************************************************

'' @CLASSTITLE:		Validator
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2005-02-06 15:25
'' @CDESCRIPTION:	Represents a general validation container which can be used for the validation of business objects 
''					or any other kind of validation.
''					It stores invalid fields (e.g. property of a class) with an associated error message
''					(why is the field invalid). The underlying storage is a dictionary.
''					- It implements a <em>reflect()</em> method thus it can be used nicely on callbacks. A callback could return a whole validator ;)
''					Example of simple usage:
''					<code>
''					<%
''					set v = new Validator
''					if lastname = "" then v.add "lastname", "Lastname cannot be empty"
''					if str.parse(age, 0) <= 0 then v.add "age", "Age must be a number and greater than 0"
''					if v then
''					.	save()
''					else
''					.	str.write(v.getErrorSummary("<ul>", "</ul>", "<li>", "</li>"))
''					end if
''					% >
''					</code>
'' @POSTFIX:		val
'' @VERSION:		0.2

'**************************************************************************************************************
class Validator

	'private members
	private dictInvalidData
	
	'public members
	public reflectItemPrefix	''[string] the prefix for each item within the summary which is returned on reflection. default = <em>&lt;li&gt;</em>
	public reflectItemPostfix	''[string] the prefix for each item within the summary which is returned on reflection. default = <em>&lt;/li&gt;</em>
	
	public default property get valid ''[bool] indicates if the validator is valid (contains no invalid fields)
		valid = (dictInvalidData.count <= 0)
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	private sub class_initialize()
		set dictInvalidData = lib.newDict(empty)
		reflectItemPrefix = "<li>"
		reflectItemPostfix = "</li>"
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns a DICTIONARY with all invalid-fields 
	'' @RETURN:			[dictionary] All descriptions and fieldnames of invalid fields
	'**************************************************************************************************************
	public function getInvalidData()
		set getInvalidData = dictInvalidData
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns the description of the error for the requested-field.
	'' @PARAM:			fieldName [string]: the name of your field to get the error-description for.
	'' @RETURN:			[string] the description of the error for the requested-field. EMPTY if there isnt any error
	'**************************************************************************************************************
	public function getDescription(fieldName)
		getDescriptionFor = empty
		if dictInvalidData.exists(uCase(fieldName)) then getDescription = dictInvalidData(uCase(fieldName))
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if a given fiel (or more fields) is invalid
	'' @PARAM:			fieldName [string], [array]: the name of your field to check (case insensitive). can also be an ARRAY with names
	'' @RETURN:			[bool] TRUE if the field is invalid otherwise FALSE
	'**************************************************************************************************************
	public function isInvalid(byVal fieldName)
		isInvalid = false
		fields = fieldName
		if not isArray(fields) then fields = array(fieldName)
		for each f in fields
			if dictInvalidData.exists(uCase(f)) then
				isInvalid = true
				exit for
			end if
		next
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Adds a new invalid field. only if it does not exists yet.
	'' @PARAM:			fieldName [string]: the name of your field. leave EMPTY if you want the field be auto-generated.
	'' @PARAM:			errorDescription [string]: a reason why the field is invalid
	'' @RETURN:			[bool] TRUE if added, FALSE if not added (because already exists)
	'**************************************************************************************************************
	public function add(byVal fieldName, errorDescription)
		if fieldname = empty then fieldName = lib.getUniqueID()
		if not dictInvalidData.exists(uCase(fieldName)) then
			dictInvalidData.add uCase(fieldName), errorDescription
			add = true
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Returns a custom formatted error summary.
	'' @DESCRIPTION:	usefull if you want to show the errors for example in a HTML list. Summary will be just
	''					returned if there is at least one error (so at least one field must be invalid).
	'' @PARAM:			overallPrefix [string]: prefix for the whole summary e.g. <em>&lt;ul&gt;</em>
	'' @PARAM:			overallPostfix [string]: postfix for the whole summary e.g. <em>&lt;/ul&gt;</em>
	'' @PARAM:			itemPrefix [string]: prefix for each item <em>&lt;li&gt;</em>
	'' @PARAM:			itemPostfix [string]: prefix for each item <em>&lt;/li&gt;</em>
	'' @RETURN:			[string] formatted error summary
	'**************************************************************************************************************
	public function getErrorSummary(overallPrefix, overallPostfix, itemPrefix, itemPostfix)
		getErrorSummary = empty
		if not me then
			getErrorSummary = getErrorSummary & overallPrefix
			for each key in dictInvalidData.keys
				getErrorSummary = getErrorSummary & itemPrefix & dictInvalidData(key) & itemPostfix
			next
			getErrorSummary = getErrorSummary & overallPostfix
		end if
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Reflection method which can be used by JSON (e.g. returning the validator as a whole on a callback)
	'' @DESCRIPTION:	as the class has no real properties the status is exposed with:
	''					- <em>data</em>: holds a dictionary with the invalid fields
	''					- <em>valid</em>: indicates if its valid or not
	''					- <em>summary</em>: holds a summary of the invalid data (<em>reflectItemPrefix</em> and <em>reflectItemPostfix</em> can be used to format the items)
	''					Example of returning it on callback (server-side asp):
	''					<code>
	''					<%
	''					sub callback(a)
	''					.	set v = new Validator
	''					.	if str.parse(age, 0) <= 0 then v.add "age", "Age is invalid"
	''					.	page.return v
	''					end sub
	''					% >
	''					</code>
	''					Afterwards the validator can be accessed directly within the javascript callback.
	''					e.g. update a HTML list with the error summary of the validator (javascript):
	''					<code>
	''					function validated(val) {
	''					.	if (val.valid) $('someList').update(val.summary);
	''					}
	''					</code>
	'' @RETURN:			[dictionary] Returns a DICTIONARY with the class properties and its values
	'**************************************************************************************************************
	public function reflect()
		set reflect = lib.newDict(empty)
		with reflect
			.add "valid", valid
			.add "data", getInvalidData()
			.add "summary", getErrorSummary("", "", reflectItemPrefix, reflectItemPostfix)
		end with
	end function

end class
%>