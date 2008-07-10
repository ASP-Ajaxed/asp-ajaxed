<!--#include file="class_datatableColumn.asp"-->
<!--#include file="class_datatableRow.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		Datatable
'' @CREATOR:		michal
'' @CREATEDON:		2008-06-06 09:26
'' @CDESCRIPTION:	Represents a control which uses data based on a SQL query and renders it as a data table.
''					The whole datatable is represented as a html table and contains rows (records) and columns (columns defined in the SQL query).
''					Columns (DatatableColumn instances) must be added to the table. Columns are auto generated if none are defined by the client.
''					Datatable contains the following main features which makes it a very powerful component:
''					- only a SQL query is needed as the datasource. All columns which are exposed in the query can be used within the table.
''					- sorting, pagination and quick deletion of records
''					- filters can be applied to columns (either predefined or free text) or even to the whole datatable (fullsearch)
''					- TODO: supports grouping of data
''					- export of the data to all kind of different formats
''					- optimized for printing
''					- it is loaded with a default style but styles can be adopted to own needs and custom application design
''					- uses ajaxed StringBuilder so it works also with large sets of data
''					- For better performance and nicer user experience all requests are handled with ajax (using pagePart feature of ajaxed)
''					- all user settings (applied filters, sorting orders, current page, ...) are remembered during the session. This makes bulk editing easier for users.
''					- offers the possibility to change properties during runtime (hide/disable rows, change styles, change data, ...).
''					Full example of a simple Datatable which shows all users of an application (columns are auto detected):
''					<code>
''					<%
''					set dt = new Datatable
''					set page = new AjaxedPage
''					page.draw()
''					
''					sub init()
''					.	dt.sql = "SELECT * FROM user"
''					.	dt.sort = "firstname"
''					end sub
''					
''					sub callback(a)
''					.	dr.draw()
''					end sub
''					
''					sub main()
''					.	dt.draw()
''					end sub
''					% >
''					</code>
''					As you can see in the example its necessary to draw the datatable in the main and the callback sub.
''					Thats needed because all the sorting, paging, etc. is done with AJAX using ajaxed callback feature (pageParts).
''					Setting the properties of the datatable must be done within the init sub (which is executed before callback and main).
''					In most cases its necessary to decide which columns should be used. Therefore columns must be added manually using newColmn() metnod:
''					<code>dt.newColumn("firstname", "Firstname")</code>
'' @REQUIRES:		Dropdown, Pageable
'' @VERSION:		0.1

'**************************************************************************************************************
class Datatable

	'private members
	private output, p_ID, columns, p_callback, sessionStorageName, dataLoaded, p_col, p_selection, p_row
	
	'public members
	public sql				''[string] SQL query which gets all the data. Paging, etc. is done by the control itself so
							''provide only the query which returns all records. Note: no ORDER BY clause (use sort property for this)
	public recsPerPage		''[int] how many records should be displayed per page. default = 100 (0 = no paging)
	public onRowCreated		''[string] name of the <strong>sub</strong> you want to execute when a row is created. it is raised before it will be drawn.
							''eventargument is the current Datatable instance. You can access the current row with the 'row' property. Example of how to disable and mark all rows which contain a deleted record with the class "deleted":
							''<code>
							''dt.onRowCreated = "onRow"
							''sub onRow(callerDT)
							''.	callerDT.row.disabled = callerDT.data("deleted") = 1
							''.	if callerDT.data("deleted") = 1 then callerDT.row.class = "deleted"
							''end sub
							''</code>
	public pkColumn			''[string], [int] name (or index) of the column which holds the primary key (uniquely identifies each record). Must be a numeric column! by default its the first column. Used for e.g. selection of records
	public attributes		''[string] any attributes which will be placed within the table tag which is used for rendering a datatable.
	public cssClass			''[string] name of the css class which should be applied to the table-tag
	public sorting			''[bool] should it be possible to sort the data? default = true
	public sort				''[string] indicates how the data should be sorted (ORDER BY part of the SQL). on sorting this property is changed (thus its possible to access it during runtime).
							''e.g. "lastname" or "created_on DESC", etc.
	public rememberState	''[bool] indicates if the state (page position, soring, filters, etc.) should be remembered for the next request. default = true
	public css				''[string] virtual path of the CSS file which should be used for formatting. set 'empty' if no specific file should be used.
							''by default the ajaxed default styles are used. By creating your own style files its possible to skin the datatable and share the styles with others.
							''Refer to the datatable.css to see how the styling works (also check the source of a generated datatable to see the classes being used)
	public cssPrint			''[string] virtual path of the CSS file which should be used for formatting when printing. set 'empty' if no specific file should be used.
	public data				''[recordset] gets the recordset which has been generated by the SQL query. Can be useful to access it when working with events like e.g. onRowCreated
	public customControls	''[string] name of the sub which draws custom controls for the datatable. e.g. if you want to add additional button, etc. with custom functions. takes the current datatable as eventargument.
							''Example of adding an own print button:
							''<code>
							''<%
							''dt.customControls = "dtCustomControls"
							''sub dtCustomControls(callerDT) % >
							''.	Please click the button to print:
							''.	<button onclick="window.print()">print</button>
							''<% end sub % >
							''</code>
	public fullsearch		''[bool] enable/disable the fullsearch feature. Lets the user search all columns with keyword(s). 
							''If more words are entered (seperated with space) then all must be found within the record. Matches are highlighted
							''with the data. Fullsearch is enabled by default.
	public nullSign			''[string] a sign which will be shown if a data value is NULL. default = empty. (The data cell gets a axdDTNull css class if it contains a null)
	
	public property let selection(val) ''[string] sets if it should be possible to select records (checkbox/radio button is automatically placed in front of each record). use "single" to allow selection of one record (radiobutton) or "multiple" for the selection of more records (checkboxes). default = empty. Name of the checkbox/radio is the ID of the datatable (needed if you do POST it with a form) and value is the value of the pkColumn.
		if val <> "" and not str.matching(val, "^single|multiple$", true) then lib.throwError("Datatable.selection only allows 'single', 'multiple' or 'empty'")
		p_selection = lCase(val)
	end property
	
	public property get selection ''[string] gets the records selection type.
		selection = p_selection
	end property
	
	public property get ID ''[string] gets the datatables ID (id attribute of the table-tag). unique within the calling page
		ID = p_ID
	end property
	
	public property get col ''[DatatableColumn] gets the column which is being currently drawn. (needed when using <strong>onCellCreated</strong> property of the DatatableColumn)
		set col = p_col
	end property
	
	public property get row ''[DatabaseRow] gets the row which is being currenly drawn. Useful when using onRowCreated
		set row = p_row
	end property
	
	private property get Lsql 'gets the sql lowercased
		Lsql = lcase(sql)
	end property
	
	private property get callback 'indicates if the request is a callback for this instance
		if isEmpty(p_callback) then p_callback = _
			lib.page.isCallback() and _
			str.startsWith(lib.page.RF(lib.page.callbackFlagName), "axd_dt_") and _
			lib.page.RF("axd_dt_id") = ID
		callback = p_callback
	end property
	
	public property get columnsCount ''[int] gets the amount of columns
		columnsCount = uBound(columns) + 1
	end property
	
	public property get fullsearchValue ''[array] gets the value which has been used within the fullsearch. It is an array because it will return each keyword if there are more (separated with space)
		fullsearchValue = trim(getState("fullsearch", lib.iif(callback(), RF("fullsearch"), empty)))
		setState "fullsearch", fullsearchValue
		fullsearchValue = split(fullsearchValue, " ")
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		'TODO: lib.require "Pageable", "Datatable"
		p_ID = "axdDT_" & lib.getUniqueID()
		sessionStorageName = "ajaxed_database"
		dataLoaded = false
		recsPerPage = 100
		auto = true
		sorting = true
		pkColumn = 0
		columns = array()
		set output = new StringBuilder
		rememberState = true
		css = path("datatable.css")
		cssPrint = path("datatablePrint.css")
		set p_col = nothing
		set p_row = new DatatableRow
		set p_row.dt = me
		fullsearch = true
		''set lang = lib.newDict(empty)
		''server.execute(path("de.asp"))
		'll(lang)
		'lib.logger.debug lang.count
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the control. Must be called within <strong>main()</strong> and <strong>callback()</strong>
	'' @DESCRIPTION:	draws the datatable with all its previosuly defined columns/filters.
	''					If no columns are defined then the control automatically grabs all columns which are defined within the sql.
	''					It must be called within the main() sub and the callback() sub of the AjaxedPage. All initializations
	''					(setting properties) must be done in the init().
	'**********************************************************************************************************
	public sub draw()
		if lib.page.isCallback() then
			if not callback then exit sub
			lib.page.callbackType = 2
		end if
		init()
		loadData()
		if not data.eof then
			if str.parse(data(pkColumn), -1) <= -1 then lib.throwError("Datatable.pkColumn must be a column which contains a positive number. When using autogenerated columns be sure that the first column is the primary key column.")
		end if
		if not callback then
			if not isEmpty(css) then lib.page.loadCSSFile css, empty
			if not isEmpty(cssPrint) then lib.page.loadCSSFile cssPrint, "print"
			lib.page.loadJSFile(path("datatable.js"))
			str.write("<table " & _
				attribute("id", ID) & _
				attribute("class", "axd_dt " & cssClass) & " " & _
				attributes & ">")
		end if
		drawHeader()
		drawData()
		if not callback then
			output("</table>")
			output("<script>var " & ID & " = new AxdDT('" & ID & "', '" & sort & "');</script>")
		end if
		str.write(output.toString())
	end sub
	
	'**********************************************************************************************************
	'* init 
	'**********************************************************************************************************
	private sub init()
		sql = trim(sql)
		if trim(sql) = "" then lib.throwError("Datatable.sql cannot be empty.")
		if str.matching(sql, "order by.*$", true) then lib.throwError("Database.sql cannot contain ORDER BY clause. Use sort property.")
		if trim(pkColumn) = "" then lib.throwError("Datatable.pkColumn must be set.")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	indicates if the datatable is sorted by a given column name
	'' @PARAM:			name [string]: name of column you want to check if its sorted
	'' @RETURN:			[int] empty = not sorted by this column otherwise ASC or DESC
	'**********************************************************************************************************
	public function sortedBy(name)
		if not dataLoaded then lib.throwError("Datatable.sortedBy is only accessible after data has been loaded.")
		sortedBy = empty
		sorts = split(lCase(sort), ",")
		for each s in sorts
			dir = split(trim(s), " ")
			if dir(0) = lCase(name) then
				if uBound(dir) > 0 then sortedBy = uCase(dir(1))
				exit function
			end if
		next
	end function
	
	'******************************************************************************************
	'* path - gets path to a file located within this control 
	'******************************************************************************************
	private function path(filename)
		path = lib.path("class_datatable/" & filename)
	end function
	
	'******************************************************************************************
	'* loadData 
	'******************************************************************************************
	private sub loadData()
		'first manipulate the sql as necessary
		if sorting then
			sort = getState("sort", sort)
			if callback and RF("sort") <> "" then sort = RF("sort")
			setState "sort", sort
		end if
		'now execute the final sql. it is actually a subselect
		'so all columns are passed through and can be accessed by filters, etc.
		sqlFinal = "SELECT * FROM (" & sql & ") datatable" & lib.iif(sort <> "", " ORDER BY " & str.SqlSafe(sort), "")
		set data = db.getUnlockedRS(sqlFinal, empty)
		
		autoGenerateColumns()
		
		if data.eof then dataLoaded = true : exit sub
		
		'now perform the fullsearch on the returned recordset
		keywords = fullsearchValue
		if uBound(keywords) = -1 then dataLoaded = true : exit sub
		
		for each c in columns
			typ = data.fields(c.name).type
			if lib.contains(db.stringFieldTypes, typ) then
				if not isEmpty(fTemplate) then fTemplate = fTemplate & " OR "
				fTemplate = fTemplate & c.name & " LIKE '*{0}*' "
			end if
		next
		for i = 0 to ubound(keywords)
			'TODO: bug with more keywords. seems not to know AND
			'TODO: use wildcards *
			'TODO: negation of the keyword with !
			if i > 0 then fltr = fltr & " AND "
			fltr = fltr & "(" & str.format(fTemplate, str.sqlSafe(keywords(i))) & ")"
		next
		data.filter = fltr
		
		dataLoaded = true
	end sub
	
	'******************************************************************************************
	'* RF 
	'******************************************************************************************
	private function RF(name)
		RF = lib.page.RFT("axd_dt_" & name)
	end function
	
	'******************************************************************************************
	'* getState - gets a stored state value of the datatable from the session storage
	'* if it does not exist then the alternative values is returned
	'******************************************************************************************
	private function getState(name, alternative)
		getState = alternative
		if not rememberState then exit function
		if isEmpty(session(sessionStorageName)) then set session(sessionStorageName) = lib.newDict(empty)
		if session(sessionStorageName).exists(name) then getState = session(sessionStorageName)(name)
	end function
	
	'******************************************************************************************
	'* setState 
	'******************************************************************************************
	private function setState(name, value)
		setState = value
		if not rememberState then exit function
		sn = sessionStorageName & "_" & lib.page.getLocation("virtual", false)
		'TODO: make state for each datatable per page
		if isEmpty(session(sn)) then set session(sn) = lib.newDict(empty)
		if session(sn).exists(name) then
			session(sn)(name) = value
		else
			session(sn).add name, value
		end if
	end function
	
	'******************************************************************************************
	'* autoGenerateColumns 
	'******************************************************************************************
	private sub autoGenerateColumns()
		if uBound(columns) > -1 then exit sub
		for each f in data.fields
			set c = newColumn(f.name, str.humanize(f.name))
			c.index = uBound(columns)
		next
	end sub
	
	'******************************************************************************************
	'* drawHeader 
	'******************************************************************************************
	private sub drawHeader()
		if callback then exit sub
		
		'NOTE: the beginning of the table is being sent directly to the response
		'so its possible to use customControls without returning the HTML as a string.
		'afterwards stringbuilder is used. check drawHeader() for more details
		str.write("<thead>")
		str.write("<tr class=""axdDTControlsRow"">")
		str.write("<td colspan=" & columnsCount + lib.iif(selection <> "", 1, 0) & ">")
		str.write("<div class=""axdDTCustomControls"">")
		if not isEmpty(customControls) then lib.exec customControls, me
		str.write("</div>")
		str.write("<div class=""axdDTControls"">")
		if fullsearch then
			str.write("<input type=""text"" ")
			str.write(attribute("value", str(str.arrayToString(fullsearchValue, " "))))
			'TODO: works fine with enter, but if there is form outside then it gets submitted on enter
			str.write(attribute("onchange", ID & ".search(this.value);"))
			str.write(")>")
		end if
		str.write("</div>")
		str.write("</td>")
		str.write("</tr>")
		
		output("<tr>")
		row.drawSelectionColumn true, output
		for each c in columns
			output(c.draw(output))
		next
		output("</tr>")
		output("</thead>")
	end sub
	
	'******************************************************************************************
	'* drawData 
	'******************************************************************************************
	private sub drawData()
		if not callback then output("<tbody id=""" & ID & "_body"">")
		num = 1
		while not data.eof
			set p_row = new DatatableRow
			set p_row.dt = me
			p_row.draw num, p_col, columns, output
			data.movenext()
			num = num + 1
		wend
		if not callback then output("</tbody>")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a new column with name and caption and returns it
	'' @PARAM:			name [string], [int]: name/index of the column (should exist within the SQL)
	'' @PARAM:			caption [string]: caption for the column header
	'' @RETURN:			[DatatableColumn] returns an already added column (properties can be changed afterwards).
	'**********************************************************************************************************
	public function newColumn(name, caption)
		set newColumn = new DatatableColumn
		with newColumn
			.name = name
			.caption = caption
			set .dt = me
			.index = uBound(columns) + 1
			redim preserve columns(.index)
			set columns(.index) = newColumn
		end with
	end function
	
	'**********************************************************************************************************
	'* attribute 
	'**********************************************************************************************************
	function attribute(name, val)
		if val <> "" then attribute = " " & name & "=""" & val & """"
	end function

end class
%>