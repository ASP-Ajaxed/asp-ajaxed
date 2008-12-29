<!--#include file="class_datatableColumn.asp"-->
<!--#include file="class_datatableRow.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		Datatable
'' @CREATOR:		michal
'' @CREATEDON:		2008-06-06 09:26
'' @CDESCRIPTION:	Represents a control which uses data based on a SQL query and renders it as a data table.
''					The whole datatable is represented as a html table and contains rows (records) and columns (columns defined in the SQL query).
''					Columns (<em>DatatableColumn</em> instances) must be added to the table. Columns are auto generated if none are defined by the client.
''					Datatable contains the following main features which makes it a very powerful component:
''					- only a SQL query is needed as the datasource. All columns which are exposed in the query can be used within the table.
''					- supports sorting & pagination
''					- full text search can be applied
''					- TODO: quick deletion of records
''					- TODO: export of the data to all kind of different formats
''					- TODO: all user settings (sorting orders, current page, ...) are remembered during the session. This makes bulk editing easier for users.
''					- optimized for printing
''					- it is loaded with a default style but styles can be adopted to own needs and custom application design
''					- uses ajaxed StringBuilder so it works also with large sets of data
''					- For better performance and nicer user experience all requests are handled with ajax (using pagePart feature of ajaxed)
''					- offers the possibility to change properties during runtime (hide/disable rows, change styles, change data, ...).
''					Full example of a fsimple Datatable which shows all users of an application (columns are auto detected):
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
''					Setting the properties of the datatable must be done within the pages <em>init()</em> sub (which is executed before <em>callback</em> and <em>main</em>).
''					In most cases its necessary to decide which columns should be used. Therefore columns must be added manually using <em>newColmn()</em> metnod:
''					<code><% dt.newColumn("firstname", "Firstname")% ></code>
'' @REQUIRES:		Dropdown, Pageable
'' @VERSION:		0.1

'**************************************************************************************************************
class Datatable

	'private members
	private output, p_ID, columns, p_callback, sessionStorageName, dataLoaded, p_col, p_selection, p_row, totalRecords
	private recordLinkPlaceholders
	
	'public members
	public sql				''[string] SQL query which gets all the data. Paging, etc. is done by the control itself so
							''provide only the query which returns all records. Note: no <em>ORDER BY</em> clause (use <em>sort</em> property for this)
	public recsPerPage		''[int] how many records should be displayed per page. default = 100 (0 = no paging)
	public onRowCreated		''[string] name of the <strong>sub</strong> you want to execute when a row is created. it is raised before it will be drawn.
							''eventargument is the current Datatable instance. You can access the current row with the <em>row</em> property. Example of how to disable and mark all rows which contain a deleted record with the class "deleted":
							''<code>
							''<%
							''dt.onRowCreated = "onRow"
							''sub onRow(callerDT)
							''.	callerDT.row.disabled = callerDT.data("deleted") = 1
							''.	if callerDT.data("deleted") = 1 then callerDT.row.class = "deleted"
							''end sub
							''% >
							''</code>
	public pkColumn			''[string], [int] name (or index) of the column which holds the primary key (uniquely identifies each record). Must be a numeric column! by default its the first column. Used for e.g. selection of records
	public attributes		''[string] any attributes which will be placed within the table tag which is used for rendering a datatable.
	public cssClass			''[string] name of the css class which should be applied to the table-tag
	public sorting			''[bool] should it be possible to sort the data? default = TRUE
	public sort				''[string] indicates how the data should be sorted (ORDER BY part of the SQL). on sorting this property is changed (thus its possible to access it during runtime and get the actual sort).
							''e.g. <em>"lastname"</em> or <em>"created_on DESC"</em>, etc.
	public name				''[string] Required if you want to preserve the datatables state (paging, sorting, ..) during the users session.
							''Provide a unique name for your datatable instance. (it must be unique within your calling page. e.g. if <em>mypage.asp</em> contains 2 datatable instances then each of them must have a unique name).
							''Why would you want to preserve the state? Check this: Your datatable contains a lot of data, the user customizes its view (sorting, paging, ...), leaves the page and then comes back. Now he
							''has to customize the datatable again. But if you preserve the state then the settings are still remembered when he comes back.
							''Leave EMPTY if you don't want to preserve the state. default = EMPTY.
	public css				''[string] virtual path of the CSS file which should be used for formatting. set EMPTY if no specific file should be used.
							''by default the ajaxed default styles are used. By creating your own style files its possible to skin the datatable and share the styles with others.
							''Refer to the <em>datatable.css</em> to see how the styling works (also check the source of a generated datatable to see the classes being used)
	public data				''[recordset] gets the recordset which has been generated by the SQL query. Can be useful to access it when working with events like e.g. <em>onRowCreated</em>
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
	public nullSign			''[string] a sign which will be shown if a data value is NULL. default = EMPTY. (The data cell gets a <em>axdDTNull</em> css class if it contains a null)
	public recordLink		''[string] if given then each cells content will be surrounded by an a href tag whose href attribute will contain this link. Perfect if you want to make whole rows clickable quickly.
							''<em>{field_name}</em> represents a placeholder for a data column value within your link where <em>field_name</em> must be replaced with the actual field name which has been selected.
							''Example of a link: <em>user.asp?i={id}</em> (which means that the ID columns value will be passed into the user.asp file as a parameter called i).
							''The link can be deactivated for specified columns by setting the <em>enableLink</em> property to FALSE.
							''Note: Only letters, numbers and underscores are allowed for the placeholders.
	
	public property get recordLinkF ''[string] gets the link (if any specified with <em>recordLink</em>) for the current record (placeholders replaced). only available during runtime.
		if not dataLoaded then lib.throwError("Datatable.recordLinkF is only accessible after draw() has been called.")
		'if the placeholders havent been loaded yet then do it
		'we cache them into a dictionary
		if isEmpty(recordLinkPlaceholders) then
			set recordLinkPlaceholders = ["D"](empty)
			with new Regexp
				.pattern = "{([a-z0-9_]+)}"
				.global = true
				.ignoreCase = true
				for each match in .execute(recordLink)
					ph = match.submatches(0)
					if not recordLinkPlaceholders.exists(ph) then recordLinkPlaceholders.add ph, empty
				next
			end with
		end if
		'now replace the placeholders with the data values
		recordLinkF = recordLink
		for each ph in recordLinkPlaceholders.keys
			recordLinkF = replace(recordLinkF, "{" & ph & "}", data(ph))
		next
	end property
	
	public property let selection(val) ''[string] sets if it should be possible to select records (checkbox/radio button is automatically placed in front of each record). use <em>"single"</em> to allow selection of one record (radiobutton) or <em>"multiple"</em> for the selection of more records (checkboxes). default = EMPTY. Name of the checkbox/radio is the ID of the datatable (needed if you do POST it with a form) and value is the value of the pkColumn.
		if val <> "" and not str.matching(val, "^single|multiple$", true) then lib.throwError("Datatable.selection only allows 'single', 'multiple' or 'empty'")
		p_selection = lCase(val)
	end property
	
	public property get selection ''[string] gets the records selection type.
		selection = p_selection
	end property
	
	public property get ID ''[string] gets the datatables ID (id attribute of the table-tag). unique within the calling page
		ID = p_ID
	end property
	
	public property get col ''[DatatableColumn] gets the column which is being currently drawn. (needed when using <strong>onCellCreated</strong> property of the <em>DatatableColumn</em>)
		set col = p_col
	end property
	
	public property get row ''[DatabaseRow] gets the row which is being currenly drawn. Useful when using <em>onRowCreated</em>
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
	
	public property get fullsearchValue ''[array] gets the value which has been used within the fullsearch. It is an ARRAY because it will return each keyword if there are more (separated with space)
		fullsearchValue = trim(getState("fullsearch", lib.iif(callback(), RF("fullsearch"), empty)))
		setState "fullsearch", fullsearchValue
		fullsearchValue = split(fullsearchValue, " ")
	end property
	
	private property get tableColumnsCount ''[int] gets the number of columns which define a data row
		tableColumnsCount = columnsCount + lib.iif(selection <> "", 1, 0)
	end property
	
	private property get currentPage ''[int] gets the page number which is currently active. 
		currentPage = str.parse(RF("page"), -1)
		if currentPage < 0 then currentPage = 1
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		p_ID = "axdDT_" & lib.getUniqueID()
		sessionStorageName = "ajaxed_datatable"
		dataLoaded = false
		recsPerPage = 100
		auto = true
		sorting = true
		pkColumn = 0
		columns = array()
		set output = new StringBuilder
		css = path("datatable.css")
		set p_col = nothing
		set p_row = new DatatableRow
		set p_row.dt = me
		fullsearch = true
		totalRecords = 0
		'set dictionary = lib.newDict(empty)
		'server.execute(path("de.asp"))
		'lib.logger.debug dictionary.count
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the control. Must be called within <strong>main()</strong> and <strong>callback()</strong>
	'' @DESCRIPTION:	draws the datatable with all its previosuly defined columns/filters.
	''					If no columns are defined then the control automatically grabs all columns which are defined within the sql.
	''					It must be called within the <em>main()</em> sub and the <em>callback()</em> sub of the AjaxedPage. All initializations
	''					(setting properties) must be done in the <em>init()</em>.
	'**********************************************************************************************************
	public sub draw()
		if lib.page.isCallback() then
			if not callback then exit sub
			lib.page.callbackType = 2
		end if
		init()
		loadData()
		if not data.eof then
			if str.parse(cstr(data(pkColumn)), -1) <= -1 then lib.throwError("Datatable.pkColumn must be a column which contains a positive number. When using autogenerated columns be sure that the first column is the primary key column.")
		end if
		if not callback then
			if not isEmpty(css) then lib.page.loadCSSFile css, empty
			lib.page.loadJSFile(path("datatable.js"))
			str.write("<table " & _
				attribute("id", ID) & _
				attribute("class", "axd_dt " & cssClass) & " " & _
				attributes & ">")
		end if
		drawHeader()
		drawData()
		drawFooter()
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
	'' @PARAM:			columnName [string]: name of column you want to check if its sorted
	'' @RETURN:			[int] EMPTY means its not sorted by this column otherwise <em>ASC</em> or <em>DESC</em>
	'**********************************************************************************************************
	public function sortedBy(columnName)
		if not dataLoaded then lib.throwError("Datatable.sortedBy is only accessible after data has been loaded.")
		sortedBy = empty
		sorts = split(lCase(sort), ",")
		for each s in sorts
			dir = split(trim(s), " ")
			if dir(0) = lCase(columnName) then
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
		sqlFinal = "SELECT * FROM (" & sql & ") datatable" & lib.iif(sort <> "", " ORDER BY " & db.SqlSafe(sort), "")
		set data = db.getUnlockedRS(sqlFinal, empty)
		totalRecords = data.recordCount
		autoGenerateColumns()
		
		if data.eof then dataLoaded = true : exit sub
		
		'now perform the fullsearch on the returned recordset
		keywords = fullsearchValue
		if uBound(keywords) = -1 then dataLoaded = true : exit sub
		
		for each c in columns
			typ = data.fields(c.name).type
			if lib.contains(db.stringFieldTypes, typ) then
				if not isEmpty(fTemplate) then fTemplate = fTemplate & " OR "
				fTemplate = fTemplate & c.name & " LIKE '%{0}%' "
			end if
		next
		for i = 0 to ubound(keywords)
			'TODO: bug with more keywords.
			'	should be (c1 LIKE x OR c2 LIKE x) AND (c1 LIKE x OR c2 LIKE x)
			'	- this would select only records which contain both keywords
			'	- but this combination does not work within ADO!
			'TODO: negation of the keyword with !
			kw = keywords(i)
			if i > 0 then fltr = fltr & " AND "
			fltr = fltr & "(" & str.format(fTemplate, db.sqlSafe(kw)) & ")"
		next
		lib.logger.debug fltr
		data.filter = fltr
		
		dataLoaded = true
	end sub
	
	'******************************************************************************************
	'* RF 
	'******************************************************************************************
	private function RF(fieldname)
		RF = lib.page.RFT("axd_dt_" & fieldname)
	end function
	
	'******************************************************************************************
	'* getState - gets a stored state value of the datatable from the session storage
	'* if it does not exist then the alternative values is returned
	'******************************************************************************************
	private function getState(sname, alternative)
		getState = alternative
		if name = "" then exit function
		if isEmpty(session(sessionStorageName)) then set session(sessionStorageName) = lib.newDict(empty)
		scope = lib.page.getLocation("virtual", false)
		if not session(sessionStorageName).exists(scope) then exit function
		if session(sessionStorageName)(scope).exists(sname) then getState = session(sessionStorageName)(scope)(sname)
	end function
	
	'******************************************************************************************
	'* setState 
	'******************************************************************************************
	private function setState(sname, value)
		setState = value
		if name = "" then exit function
		scope = lib.page.getLocation("virtual", false)
		if isEmpty(session(sessionStorageName)) then set session(sessionStorageName) = ["D"](array(scope, ["D"](empty)))
		if not session(sessionStorageName).exists(scope) then session(sessionStorageName).add scope, ["D"](empty)
		if not session(sessionStorageName)(scope).exists(sname) then
			session(sessionStorageName)(scope).add sname, value
		else
			session(sessionStorageName)(scope)(sname) = value
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
		with str
			.write("<thead>")
			.write("<tr class=""axdDTControlsRow"">")
			.write("<td colspan=" & tableColumnsCount & ">")
			.write("<div class=""axdDTCustomControls"">")
			if not isEmpty(customControls) then lib.exec customControls, me
			.write("</div>")
			.write("<div class=""axdDTControls"">")
			if fullsearch then
				.write("<input type=""text"" ")
				.write(attribute("value", str(.arrayToString(fullsearchValue, " "))))
				'TODO: works fine with enter, but if there is form outside then it gets submitted on enter
				'.write(attribute("onkeypress", "if(event.keyCode == 13) " & ID & ".search(this.value);"))
				.write(attribute("onchange", ID & ".search(this.value)"))
				.write(">")
			end if
			.write("</div>")
			.write("</td>")
			.write("</tr>")
		end with
		output("<tr>")
		row.drawSelectionColumn true, output
		for each c in columns
			output(c.draw(output))
		next
		output("</tr>")
		output("</thead>")
	end sub
	
	'******************************************************************************************
	'* drawFooter 
	'******************************************************************************************
	private sub drawFooter()
		if not callback then output("<tfoot>")
		output("<tr><td colspan=""")
		output(tableColumnsCount)
		output(""">")
		set dc = (new DataContainer)(data)
		
		'determine the bounds of displayed records
		firstRecord = (lib.iif(currentPage = 0, 1, currentPage) - 1) * recsPerPage
		lastRecord = firstRecord + recsPerPage
		firstRecord =  firstRecord + 1
		if lastRecord = 0 then
			lastRecord = dc.count
		elseif lastRecord > dc.count then
			lastRecord = dc.count
		end if
		
		output("<div class=""recordsIndicator"">")
		if lastRecord - firstRecord + 1 = dc.count then
			output(str.format("Displaying {0} records ", dc.count))
		else
			output(str.format("Displaying {0}-{1} of {2} records ", array(firstRecord, lastRecord, dc.count)))
		end if
		output(str.format("({0} total)", totalRecords))
		output(".")
		output("</div>")
		if recsPerPage > 0 then
			output("<div class=""pagingBar"">")
			pages = dc.paginate(recsPerPage, currentPage, 10)
			if ubound(pages) > 0 then
				if currentPage > 1 then output("<span class=""pPrev""><a href=""javascript:" & ID & ".goTo(" & currentPage - 1 & ")"">&lt; prev " & recsPerPage & "</a></span>")
				for each p in pages
					if p = "..." then
						output("<span class=""pMore"">...</span>")
					elseif p = currentPage then
						output("<span class=""pCurrent"">" & p & "</span>")
					else
						output("<span class=""pPage""><a href=""javascript:" & ID & ".goTo(" & p & ")"">" & p & "</a></span>")
					end if
					lastPage = p
				next
				if lastPage = "..." then lastPage = 0
				if currentPage > 0 and (currentPage < lastPage or lastPage = 0) then
					output("<span class=""pNext""><a href=""javascript:" & ID & ".goTo(" & currentPage + 1 & ")"">next " & recsPerPage & " &gt;</a></span>")
				end if
				output("<span class=""pAll" & lib.iif(currentPage = 0, " pCurrent", "") & """>")
				if currentPage = 0 then
					output("all")
				else
					output("<a href=""javascript:" & ID & ".goTo(0)"">all</a>")
				end if
				output("</span>")
			end if
			output("</div>")
		end if
		output("</td></tr>")
		if not callback then output("</tfoot>")
	end sub
	
	'******************************************************************************************
	'* drawData 
	'******************************************************************************************
	private sub drawData()
		if not callback then output("<tbody id=""" & ID & "_body"">")
		num = 1
		'if paging is enabled then prepare the recordset for paging
		if not data.eof and recsPerPage > 0 and currentPage > 0 then
			data.cacheSize = recsPerPage
			data.pageSize = recsPerPage
			data.absolutePage = currentPage
		end if
		while not data.eof and (num <= recsPerPage or recsPerPage = 0 or currentPage = 0)
			set p_row = new DatatableRow
			set p_row.dt = me
			p_row.draw num + (lib.iif(currentPage = 0, 1, currentPage) - 1) * recsPerPage, p_col, columns, output
			data.movenext()
			num = num + 1
		wend
		if not callback then output("</tbody>")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	adds a new column with name and caption and returns it
	'' @PARAM:			cname [string], [int]: name/index of the column (should exist within the <em>sql</em>)
	'' @PARAM:			caption [string], [array]: caption for the column header. If ARRAY then the first value is the caption and the second is the help text.
	'' @RETURN:			[DatatableColumn] returns an already added column (properties can be changed afterwards).
	'**********************************************************************************************************
	public function newColumn(cname, caption)
		set newColumn = new DatatableColumn
		with newColumn
			.name = cname
			if isArray(caption) then
				if uBound(caption) <> 1 then lib.throwError("Datatable.newColumn() caption if used as array must contain 2 elements")
				.caption = caption(0)
				.help = caption(1)
			else
				.caption = caption
			end if
			set .dt = me
			.index = uBound(columns) + 1
			redim preserve columns(.index)
			set columns(.index) = newColumn
		end with
	end function
	
	'**********************************************************************************************************
	'* attribute 
	'**********************************************************************************************************
	function attribute(aname, val)
		if val <> "" then attribute = " " & aname & "=""" & val & """"
	end function

end class
%>