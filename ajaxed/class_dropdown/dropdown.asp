<!--#include file="class_dropdownItem.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		Dropdown
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		2005-02-01 16:31
'' @CDESCRIPTION:	represents a HTML-selectbox in an OO-approach.
''					- naming and functionallity is based on .NET-control <em>DropdownList</em>
''					- easy use with different datasources: RECORDSET, DICTIONARY and ARRAY
''					- there are 3 different rendering types when using it as a multiple dropdown. e.g. you can render a dropdown where each item has a radiobutton so the user can make only one selection. check <em>multipleSelectionType</em> property for more details
''					- to build a new dropdown create a new instance, set the desired properties and use <em>draw()</em> method to render the dropdown on the page. If the output is needed as a string then <em>toString()</em> can be used
''					- in most cases you need a dropdown with data from the database. in that case you can use <em>getNew()</em> method which quickly creates a dropdown for you. as it is the default method you can simply use it with <code><% (new Dropdown)("SELECT * FROM table", "name", "selected").toString() % ></code>
''					- Example of simply creating a dropdown with all months of a year and the current month selected: <code><% (new Dropdown)(lib.range(1, 12, 1), "month", month(date())).toString() % ></code>
'' @POSTFIX:		DD
'' @COMPATIBLE:		Internet Explorer, Mozilla Firefox
'' @VERSION:		1.0

'**************************************************************************************************************

const DD_DATASOURCE_ARRAY			= 1
const DD_DATASOURCE_RECORDSET		= 2
const DD_DATASOURCE_DICTIONARY		= 4
const DD_OUTPUT_DIRECT				= 1
const DD_OUTPUT_STRING				= 2
const DD_SELECTIONTYPE_COMMON		= 1
const DD_SELECTIONTYPE_MULTIPLE		= 2
const DD_SELECTIONTYPE_SINGLE		= 4

class Dropdown

	'private members
	private output					'[string], [stringBuilder] holds the complete output-string if dropdown will be returned as string
	private outputMethod			'[int] 0 = direct output, 1 = output to string
	private datasourceType			'[int] 1 = array, 2 = adodb.recordset, 4 = dictionary
	private currentIteration		'[int] index of the current iteration for the items
	private tmpDatasourceLength		'[int] temporary stored length of datasource. saves time on looping!
	private datasourceItems, datasourceKeys, p_selectedValue, selectedFound, p_uniqueID
	
	private property get newCBName 'gets the name of the checkbox which is used to create a new item
		newCBName = name & "_CB"
	end property
	
	'public members
	public name						''[string] Name of the control
	public datasource				''[array], [recordset], [dictionary], [string] when using recordset be sure 
									''to use <em>Database.getUnlockedRS()</em>
									''- if datasource is a string then it will be recognized as a SQL-query (note: after <em>draw()</em> datasource changes to a recordset)
									''- When using an array as datasource, you should set the <em>valuesDatasource</em> if values are different than the captions
									''- dictionaries values are used as values and the items as the captions for each option
									'' property. If this is not set,the datasource elements will be used as the values and options for the dropdown
	public ID						''[string] ID of the control. is generated automatically by default but can be set also.
	public onItemCreated			''[string] name of function (sub) which should handle onItemCreated. Event will be raised just before printing the item.
	public attributes				''[string] additional attributes which go into the <em>&lt;select></em>-element. e.g. onClick, onChange, etc.
	public style					''[string] css-Styles for the control. added to <em>&lt;select style="..."</em>
	public cssClass					''[string] name of the css-class you want to assign to the control
	public valuesDatasource			''[array] if array is used as datasource, please provide an array of same length for values too. If no array is given, values are indexed
	public dataValueField			''[string], [int] field-name of the datasource that provides the value for each list-item. if the datasource is a recordset then its the first column by default
	public dataTextField			''[string], [int] field-name of the datasource that provides the text-content for each list-item. if the datasource is a recordset then its the second column by default
	public tabIndex					''[int] index of the control. -1 = dont add to the tabindex-collection
	public multiple					''[bool] is it a multiple dropdown? default = false
	public size						''[int] number of displayed rows if its a multiple dropdown. default = 1. If its not a common-multiple-dropdown
									''then the size is used as pixels for the height of the dropdown.
	public disabled					''[bool] indicates whether the control is disabled or not. default = false
	public commonFieldText			''[string] text of the common-field. e.g. <em>"--- please select a value ---"</em>
	public commonFieldValue			''[string] value for the common-field. default = 0
	public multipleSelectionType	''[int] Set this property if you want to change the selectiontype of a multiple dropdown.
									''Useful if you want to have formatted items, because the dropdown isn't rendered as a SELECT-Element. There are 3 variations:
									''- <em>DD_SELECTIONTYPE_COMMON</em> = its the common one. hold down CTRL to select more
									''- <em>DD_SELECTIONTYPE_MULTIPLE</em> = selection comes with checkboxes (user can select more items).
									''- <em>DD_SELECTIONTYPE_SINGLE</em> = selection comes with radiobuttons. so just one selection is allowed.
	public autoDrawItems			''[bool] defines if each item will be drawn autmatically after the <em>onItemCreated</em>-Event
									''has been raised. default = TRUE. disabling usefull if you want to add items during the runtime
	
	public property get uniqueID ''[int] gets a unique ID (scope: the page) of the dropdown
		uniqueID = p_uniqueID
	end property
	
	public property let selectedValue(val) ''[string], [array] what value(s) is selected. array needed if multiple dropdown
		if isArray(val) then
			p_selectedValue = val
		else
			redim p_selectedValue(0)
			p_selectedValue(0) = val & "" 'concat to allow NULLS
		end if
	end property
	
	public property get selectedValue ''[array], [string] returns the selected value(s)
		if uBound(p_selectedValue) = -1 then
			selectedValue = ""
		else
			selectedValue = p_selectedValue(0)
		end if
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		tmpDatasourceLength		= -1
		size					= 1
		currentIteration		= 0
		commonFieldValue		= 0
		p_selectedValue			= array()
		valuesDatasource		= array()
		dataValueField			= 0
		dataTextField			= 1
		multiple				= false
		disabled				= false
		autoDrawItems			= true
		selectedFound			= false
		commonFieldText			= empty
		output	 				= empty
		p_uniqueID				= lib.getUniqueID()
		ID						= "dropdown_" & uniqueID
		outputMethod			= DD_OUTPUT_DIRECT
		datasourceType			= DD_DATASOURCE_ARRAY
		multipleSelectionType	= DD_SELECTIONTYPE_COMMON
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set output = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the dropdown on the page
	'**********************************************************************************************************
	public sub draw()
		outputMethod = DD_OUTPUT_DIRECT
		renderControl()
		str.write(getOutput())
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	STATIC! helper method which quickly creates a dropdown using a SQL-query
	'' @DESCRIPTION:	first column is used as value for the items and second as the caption
	'' @PARAM:			datasrc [dictionary], [recordset], [string], [array]: datasource for the dropdown
	'' @PARAM:			dropdownName [string]: name of the dropdown
	'' @PARAM:			selectedVal [string]: selected value
	'' @RETURN:			[Dropdown] a ready-to-use dropdown
	'**********************************************************************************************************
	public default function getNew(datasrc, dropdownName, selectedVal)
		set getNew = new Dropdown
		with getNew
			if isObject(datasrc) then
				set .datasource = datasrc
			else
				.datasource = datasrc
			end if
			.name = dropdownName
			.selectedValue = selectedVal
		end with
	end function
	
	'**********************************************************************************************************
	'' @DESCRIPTION:	gets you a new dropdown-item which can be drawn after that.
	'' @PARAM:			itemValue [string] value of the item
	'' @PARAM:			itemText [string] text of the item
	'' @RETURN:			[DropdownItem] the created dropdownItem
	'**********************************************************************************************************
	public function getNewItem(itemValue, itemText)
		set getNewItem = getRawItem(getDatasourceLength() + 1, itemValue, itemText)
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns the dropdown as a string
	'' @RETURN:			[string]: returns a string representation of this dropdown
	'**********************************************************************************************************
	public function toString()
		outputMethod = DD_OUTPUT_STRING
		renderControl()
		toString = getOutput()
	end function
	
	'**********************************************************************************************************
	' renderControl 
	'**********************************************************************************************************
	private function renderControl()
		if trim(name) = "" then lib.throwError("Dropdown: name must be given")
		selectedFound = false
		initStringBuilder()
		determineDatasource()
		
		printBeginTag()
		
		if commonFieldText <> empty and not multiple then addItem(getRawItem(-1, commonFieldValue, commonFieldText))
		
		for currentIteration = 0 to getDatasourceLength()
			addItem(getRawItem(currentIteration, getCurrentItemValue(), getCurrentItemText()))
			moveDatasourceCursor()
		next
		
		printEndTag()
	end function
	
	'**************************************************************************************************************
	' printBeginTag 
	'**************************************************************************************************************
	private sub printBeginTag()
		if isCommonDropdown() then
			print("<select" & _
				getAttribute("tabindex", tabindex) & _
				getAttribute("size", size) & _
				getAttribute("name", name) & _
				getAttribute("style", style) & _
				lib.iif(multiple, " multiple ", empty))
		else
			print("<div" & getAttribute("style", style))
		end if
		
		print(getAttribute("id", ID) & _
			getAttribute("class", cssClass) & _
			lib.iif(disabled or dropdownDisabled, " disabled", empty) & _
			" " & attributes & ">" & vbCrLf)
	end sub
	
	'**************************************************************************************************************
	' printEndTag 
	'**************************************************************************************************************
	private sub printEndTag()
		if isCommonDropdown() then
			print("</select>")
		else
			print("</div>")
		end if
	end sub
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	gets an Attribute HTML-string
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'' @PARAM:			attributeName [string], [int] name int of the attribute
	'' @PARAM:			attributeValue [string] value of the attribute
	'' @RETURN:			[string]: returns a nicely formatted attribute
	'**************************************************************************************************************
	function getAttribute(attributeName, attributeValue)
		if attributeValue <> "" then getAttribute = " " & attributeName & "=""" & attributeValue & """"
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	common Dropdown means that it will be rendered as a SELECT-Tag.  For internal use only!
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'' @RETURN:			[bool]: returns true if it is a common dropdown
	'**************************************************************************************************************
	function isCommonDropdown()
		isCommonDropdown = (not multiple or (multiple and multipleSelectionType = DD_SELECTIONTYPE_COMMON))
	end function
	
	'**************************************************************************************************************
	' initStringBuilder 
	'**************************************************************************************************************
	private sub initStringBuilder()
		set output = new StringBuilder
	end sub
	
	'**********************************************************************************************************
	' getDropdownItem 
	'**********************************************************************************************************
	private function getRawItem(itemIndex, itemValue, itemText)
		set getRawItem = new dropdownItem
		with getRawItem
			set .dropdown = me
			.index = itemIndex
			.value = itemValue & ""
			.text = itemText & ""
		end with
	end function
	
	'**********************************************************************************************************
	'* addItem 
	'**********************************************************************************************************
	private sub addItem(item)
		item.selected = isSelected(item.value)
		raiseOnItemCreatedEvent(item)
		if autoDrawItems then item.draw()
	end sub
	
	'**********************************************************************************************************
	' isSelected 
	'**********************************************************************************************************
	private function isSelected(valueToCheck)
		if not selectedFound then
			for i = 0 to uBound(p_selectedValue)
				if cStr(p_selectedValue(i)) = valueToCheck then
					isSelected = true
					selectedFound = (not multiple)
					exit for
				end if
			next
		end if
	end function
	
	'**********************************************************************************************************
	' raiseOnItemCreatedEvent 
	'**********************************************************************************************************
	private sub raiseOnItemCreatedEvent(byRef currentDropdownItem)
		if onItemCreated <> empty then
			set eHandler = getRef(onItemCreated)
			eHandler(currentDropdownItem)
		end if
	end sub
	
	'**********************************************************************************************************
	' determineDatasource 
	'**********************************************************************************************************
	private sub determineDatasource()
		if isArray(datasource) then
			datasourceType = DD_DATASOURCE_ARRAY
		else
			select case lCase(typename(datasource))
				case "recordset"
					datasourceType = DD_DATASOURCE_RECORDSET
				case "dictionary"
					datasourceItems = datasource.items
					datasourceKeys = datasource.keys
					datasourceType = DD_DATASOURCE_DICTIONARY
				case else 'its a string
					set datasource = db.getUnlockedRS(datasource, empty)
					datasourceType = DD_DATASOURCE_RECORDSET
			end select
		end if
	end sub
	
	'**********************************************************************************************************
	' getDatasourceLength 
	'' @RETURN: 	[int] gets the length of the datasource
	'**********************************************************************************************************
	private function getDatasourceLength()
		if tmpDatasourceLength = -1 then
			select case datasourceType
				case DD_DATASOURCE_ARRAY
					tmpDatasourceLength = uBound(datasource)
				case DD_DATASOURCE_RECORDSET
					tmpDatasourceLength = datasource.recordCount - 1
				case DD_DATASOURCE_DICTIONARY
					tmpDatasourceLength = datasource.count - 1
			end select
		end if
		getDatasourceLength = tmpDatasourceLength
	end function
	
	'**********************************************************************************************************
	' geCurrentItemValue 
	'' @RETURN: 	[string] gets the value of the item on current iteration
	'**********************************************************************************************************
	private function getCurrentItemValue()
		select case datasourceType
			case DD_DATASOURCE_ARRAY
				if uBound(valuesDatasource) > -1 then
					getCurrentItemValue = valuesDatasource(currentIteration)
				else
					getCurrentItemValue = datasource(currentIteration)
				end if
			case DD_DATASOURCE_RECORDSET
				getCurrentItemValue = datasource.fields(dataValuefield)
			case DD_DATASOURCE_DICTIONARY
				getCurrentItemValue = datasourceKeys(currentIteration)
		end select
	end function
	
	'**********************************************************************************************************
	' getCurrentItemText 
	'' @RETURN: 	[string] gets the text of the item on current iteration
	'**********************************************************************************************************
	private function getCurrentItemText()
		select case datasourceType
			case DD_DATASOURCE_ARRAY
				getCurrentItemText = datasource(currentIteration)
			case DD_DATASOURCE_RECORDSET
				getCurrentItemText = datasource.fields(dataTextfield)
			case DD_DATASOURCE_DICTIONARY
				getCurrentItemText = datasourceItems(currentIteration)
		end select
	end function
	
	'**********************************************************************************************************
	' moveDatasourceCursor 
	'**********************************************************************************************************
	private sub moveDatasourceCursor()
		select case datasourceType
			case DD_DATASOURCE_RECORDSET
				datasource.movenext()
		end select
	end sub
	
	'**********************************************************************************************************
	' getOutput 
	'**********************************************************************************************************
	private function getOutput()
		getOutput = output.toString()
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	prints out to the output. For internal use only!
	'' @DESCRIPTION:	just for internal use. public because the dropdownItem-class needs it too.
	'' @PARAM:			value [string] the value to print
	'**************************************************************************************************************
	sub print(value)
		if outputMethod = DD_OUTPUT_DIRECT then
			str.write(value)
		else
			output.append(value)
		end if
	end sub

end class
%>