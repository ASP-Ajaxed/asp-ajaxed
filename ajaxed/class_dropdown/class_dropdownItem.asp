<%
'**************************************************************************************************************

'' @CLASSTITLE:		DropdownItem
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		2005-02-02 15:16
'' @CDESCRIPTION:	represents an item of the dropdown
'' @VERSION:		0.2
'' @FRIENDOF:		Dropdown

'**************************************************************************************************************
class DropdownItem

	'public members
	public index				''[int] index of the item in the dropdown
	public value				''[string] value of the item. refers to <em><option value=""></em>
	public text					''[string] text of the item. refers to the displayed-value of the item
	public style				''[string] css-Styles for the item
	public title				''[string] title for the option
	public selected				''[bool] indicates whether the item is selected or not
	public attributes			''[string] additional attributes
	public show					''[bool] show the item or not. default = TRUE
	public dropdown				''[Dropdown] the dropdown it belongs to.
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		index				= -1
		value 				= empty
		text				= empty
		style				= empty
		title				= empty
		selected			= false
		attributes			= empty
		show 				= true
		set dropdown 		= nothing
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	private sub class_terminate()
		set dropdown = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the item using the output of the dropdown 
	'**********************************************************************************************************
	public sub draw()
		if show then
			dropdown.print(vbTab)
			printBeginTag()
			printText()
			printEndTag()
		end if
	end sub
	
	'**************************************************************************************************************
	' printText 
	'**************************************************************************************************************
	private sub printText()
		if dropdown.isCommonDropdown() then
			dropdown.print(str.HTMLEncode(text))
		else
			dropdown.print("<label for=""" & dropdown.name & dropdown.uniqueID & index & """>" & str.HTMLencode(text) & "</label>")
		end if
	end sub
	
	'**************************************************************************************************************
	' printBeginTag 
	'**************************************************************************************************************
	private sub printBeginTag()
		if dropdown.isCommonDropdown() then
			dropdown.print("<option value=""" & value & """" & _
				lib.iif(selected, " selected", empty) & _
				dropdown.getAttribute("style", style) & _
				dropdown.getAttribute("title", title) & _
				lib.iif(attributes <> empty, " " & attributes, empty) & ">")
		else
			dropdown.print("<div" & _
				dropdown.getAttribute("style", style) & _
				dropdown.getAttribute("title", title) & _
				lib.iif(attributes <> empty, " " & attributes, empty) & _
				"><input" & _
				dropdown.getAttribute("type", lib.iif(dropdown.multipleSelectionType = DD_SELECTIONTYPE_MULTIPLE, "checkbox", "radio")) & _
				dropdown.getAttribute("id", dropdown.name & dropdown.uniqueID & index) & _
				dropdown.getAttribute("name", dropdown.name) & _
				dropdown.getAttribute("value", value) & _
				lib.iif(selected, " checked=""checked""", empty) & ">")
		end if
	end sub
	
	'**************************************************************************************************************
	' printEndTag 
	'**************************************************************************************************************
	private function printEndTag()
		if dropdown.isCommonDropdown() then
			dropdown.print("</option>")
		else
			dropdown.print("</div>")
		end if
		dropdown.print(vbCrLf)
	end function

end class
%>