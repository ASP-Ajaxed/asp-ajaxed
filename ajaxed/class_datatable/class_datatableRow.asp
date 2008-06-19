<%
'**************************************************************************************************************

'' @CLASSTITLE:		DatatableRow
'' @CREATOR:		michal
'' @CREATEDON:		2008-06-17 15:56
'' @CDESCRIPTION:	Represents a row in a datatable. Can only be accessed when using
''					the onRowCreated event of the Datatable. When the event is raised its
''					possible to access the current row with the row property of the Datatable.
''					Example of how to mark first 10 records of a datatable as selected using the row's selected property:
''					<code>
''					dt.selection = "multiple"
''					dt.onRowCreated = "onRow"
''					sub onRow(callerDT)
''					.	callerDT.row.selected = callerDT.row.number <= 10
''					end sub
''					</code>
'' @FRIENDOF:		Datatable
'' @VERSION:		0.1

'**************************************************************************************************************
class DatatableRow

	'private members
	private p_number, p_selected
	
	'public members
	public dt			''[Datatable] the datatable which contains the row
	public cssClass		''[string] css class which will be placed within the &lt;tr> tag
	public disabled		''[bool] indicates if the row is disabled or not. disabled does not allow to select the row and it wont be clickable
	
	public property get selected ''[bool] indicates if the row is selected or not. keeps state after postback as well. So when user changes the selection it will remember the selection after postback.
		selected = p_selected
		if lib.page.isPostback() then
			selected = lib.contains(lib.page.RFA(dt.ID), PK)
		end if
	end property
	
	public property let selected(val) ''[bool]
		p_selected = val
	end property
	
	public default property get number ''[int] gets the number of the row within the datatable. on paging the number is continously. Good for numbering your records
		number = p_number
	end property
	
	public property get ID ''[int] gets a unique ID of the row. the &lt;tr> tag contains this ID
		ID = dt.ID & "_row_" & PK
	end property
	
	public property get PK ''[int] gets the primary key value of the rows record
		'TODO: check if the PK is a number and throw a nice understandable error
		PK = cLng(dt.data.fields(dt.pkColumn))
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set dt = nothing
	end sub
	
	'**********************************************************************************************************
	'* draw 
	'**********************************************************************************************************
	sub draw(num, byRef currentCol, byRef cols, byRef output)
		raiseOnRowCreated()
		p_number = num
		css = lib.iif(number mod 2 = 0, "axdDTRowEven", "axdDTRowOdd") & " " & _
			lib.iif(selected, "axdDTRowSelected " & cssClass, cssClass)
		output("<tr id=""" & ID & """ " & dt.attribute("class", css) & ">")
		drawSelectionColumn false, output
		for each currentCol in cols
			currentCol.drawData dt.data.fields(currentCol.name), output
		next
		output("</tr>")
	end sub
	
	'******************************************************************************************
	'* raiseOnRowCreated 
	'******************************************************************************************
	private sub raiseOnRowCreated()
		if isEmpty(dt.onRowCreated) then exit sub
		set onRowCreatedFunc = lib.getFunction(dt.onRowCreated)
		if onRowCreatedFunc is nothing then lib.throwError("Datatable.onRowCreated sub '" & dt.onRowCreated & "' does not exist.")
		onRowCreatedFunc(dt)
	end sub
	
	'******************************************************************************************
	'* drawSelectionColumn 
	'******************************************************************************************
	sub drawSelectionColumn(header, byRef output)
		if dt.selection = "" then exit sub
		if header then
			output("<th class=""axdDTColFirst axdDTColSelection""></th>")
		else
			output("<td class=""axdDTColFirst axdDTColSelection"">")
			if dt.selection = "single" then singl = true
			output("<input")
			output(dt.attribute("type", lib.iif(singl, "radio", "checkbox")))
			output(dt.attribute("name", dt.ID))
			output(dt.attribute("value", PK))
			output(dt.attribute("disabled", lib.iif(disabled, "disabled", empty)))
			output(dt.attribute("checked", lib.iif(selected, "checked", empty)))
			output(dt.attribute("onclick", dt.ID & ".toggleRow('" & ID & "', this.checked, " & lib.iif(singl, "true", "false") & ")"))
			output("/></td>")
		end if
	end sub

end class
%>