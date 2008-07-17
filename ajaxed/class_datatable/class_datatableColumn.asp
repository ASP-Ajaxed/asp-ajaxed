<%
'**************************************************************************************************************

'' @CLASSTITLE:		DatatableColumn
'' @CREATOR:		michal
'' @CREATEDON:		2008-06-06 13:15
'' @CDESCRIPTION:	Represents a column of the Datatable. It is identified by a name (must exist in the datatables SQL query).
''					Use the <em>Datatable.newColumn(name, caption)</em> factory method to add a new column to the datatable.
''					Assigning the column to a variable allows changing further properties.<br>
''					Example of adding a column for a database column named <em>firstname</em> and labeling it <em>Firstname</em>.
''					Furthermore it will use the css class <em>colFirstname</em> (we assume an existing Datatable instance <em>dt</em>):
''					<code>
''					<%
''					set c = dt.newColumn("firstname", "Firstname")
''					c.cssClass = "colFirstname"
''					% >
''					</code>
'' @FRIENDOF:		Datatable
'' @VERSION:		0.1

'**************************************************************************************************************
class DatatableColumn

	'private members
	private cellCreatedFunc, p_value, p_nullValue
	
	'public members
	public name				''[string], [int] name/index of the column. must exist in datatables SQL query.
	public caption			''[string] a caption which will be displayed as the columns header
	public index			''[int] gets the index of the column. starting with 0 (first column)
	public onCellCreated	''[string] name of the <strong>function</strong> which should be executed just before a data cell of this column will be drawn.
							''expects one argument which will hold the datatable executing the function. The function must return the value which should be drawn.
							''The datatable has a col property which holds the current column. using its value property its possible to acces the data value. 
							''Example of how to change the data cell value using other data columns from the current row
							''(e.g. making the text red if the record is deleted):
							''<code>
							''<%
							''function onFirstname(dt)
							''.	color = lib.iif(dt.data("deleted") = 1, "#f00", "#000")
							''.	onFirstname = "<span style=""color:" & color & """>" & dt.col & "</span>"
							''end function
							''% >
							''</code>
	public dt				''[Datatable] the datatable it belongs to
	public cssClass			''[string] gets/sets the css class which should be used for the column. Its being used within the header cells (&lt;th> tags) and the data cells (&lt;td> tags).
							''With proper css selectors its possible to style the header different than the data cells of a column.
							''Example with a css class named "colID" (makes the header bold and data cells colored green):
							''<code>
							''th.colID {
							''.	font-weight:bold;
							''}
							''td.colID {
							''.	color:#0f0;
							''}
							''</code>
	public encodeHTML		''[bool] should the data value encode HTML markup? default = TRUE. Set this to false if you want that HTML is recognized within the data cells
	
	public default property get value ''[string] gets the data value of the column when a cell is created. Its always a string. If its a null its converted into an empty string. Use <em>nullValue</em> property to check if it was NULL
		value = p_value
	end property
	
	public property get nullValue ''[bool] indicates if the value is actually a NULL value in the database. 
		nullValue = p_nullValue
	end property
	
	private property get css
		if index = 0 and dt.selection = "" then css = "axdColFirst "
		if index = dt.columnsCount - 1 then css = "axdColLast "
		if nullValue then css = css & "axdDTNull "
		css = css & cssClass
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		caption = "column"
		index = 0
		set dt = nothing
		encodeHTML = true
		p_nullValue = false
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the column. Internal use!
	'**********************************************************************************************************
	sub draw(byRef output)
		output("<th class=""axdCol_" & lCase(name) & " " & css & """>")
		if dt.sorting then output( _
			"<a href=""javascript:void(0);"" " & _
				"onclick=""" & dt.ID & ".sort('" & name & "');"" " & _
				"ondblclick=""" & dt.ID & ".sort('" & name & "');""" & _
			">")
		output(caption)
		if dt.sorting then output("</a>")
		output("</th>")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the value of data cell of this row. Internal use!
	'**********************************************************************************************************
	sub drawData(val, output)
		p_nullValue = isNull(val)
		if p_nullValue then p_value = "" else p_value = cstr(val) end if
		if encodeHTML and not nullValue then p_value = str(p_value)
		if not isEmpty(onCellCreated) then
			if isEmpty(cellCreatedFunc) then
				set cellCreatedFunc = lib.getFunction(onCellCreated)
				if cellCreatedFunc is nothing then lib.throwError("DatatableColumn.onCellCreated function '" & onCellCreated & "' does not exist for column '" & name & "'")
			end if
			p_value = cellCreatedFunc(dt)
		end if
		output("<td class=""" & css & """>")
		if nullValue and dt.nullSign <> "" then
			output(dt.nullSign)
		else
			output(highlight(value, dt.fullsearchValue))
		end if
		output("</td>")
	end sub
	
	'**********************************************************************************************************
	'* highlight 
	'**********************************************************************************************************
	private function highlight(val, keywords)
		if ubound(keywords) = -1 then highlight = val
		for each k in keywords
			highlight = highlight & _
				str.rReplace(val, "(" & k & ")", _
					"<span class=""axdDTHighlight"">$1</span>", true)
		next
	end function

end class
%>