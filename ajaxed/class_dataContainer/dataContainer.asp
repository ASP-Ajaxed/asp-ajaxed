<%
'**************************************************************************************************************

'' @CLASSTITLE:		DataContainer
'' @CREATOR:		michal
'' @CREATEDON:		2008-07-12 06:28
'' @CDESCRIPTION:	Represents generic container which holds data. The container offers different
''					functions to manipulate the underlying data. The data is bound by reference and therefore all changes will affect the original source.
''					Ajaxed loads this class automatically so its available everywhere. Example of usage with different datasources:
''					<code>
''					<%
''					'container with an array as datasource
''					set dc = (new DataContainer)(array(1, 2, 3))
''					
''					'container with a dictionary as datasource
''					set dc = (new DataContainer)(lib.newDict(1, 2, 3))
''					
''					'container for a recordset
''					set dc = (new DataContainer)(lib.getRS("SELECT * FROM table"))
''					% >
''					</code>
''					After the intantiation its possible to use the different methods on the created DataContainer. Example:
''					<code>
''					<%
''					set dc = (new DataContainer)(array(1, 2, 3))
''					'check if a given value exists in the container
''					dc.contains(2)
''					'its even possible in one line
''					((new DataContainer)(array(1, 2, 3))).contains(2)
''					% >
''					</code>
'' @COMPATIBLE:		Array, ADODB.Recordset, Scripting.Dictionary
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class DataContainer

	'private members
	private instantiated, p_datasource
	
	'public members
	public data ''[array], [recordset], [dictionary] gets the underlying data (which was set on initialization). <strong>Never set this property manually!</strong>
	
	public property get first ''[variant] Gets the first data item. If data is a DICTIONARY then the first <em>key</em> is returned.
		init()
		if datasource = "array" then
			if uBound(data) > -1 then first = data(0)
		elseif datasource = "dictionary" then
			keys = data.keys
			first = keys(0)
		elseif datasource = "recordset" then
			first = getRecordsetBound("first")
		else
			notSupported("first")
		end if
	end property
	
	public property get last ''[variant] gets the last data item. The last <em>key</em> is returned if data is a DICTIONARY. The last first column is returned if its a RECORDSET. Returns EMPTY if <em>data</em> has no items.
		init()
		last = empty
		if datasource = "array" then
			if uBound(data) > -1 then last = data(uBound(data))
		elseif datasource = "dictionary" then
			keys = data.keys
			last = keys(uBound(keys))
		elseif datasource = "recordset" then
			last = getRecordsetBound("last")
		else
			notSupported("last")
		end if
	end property
	
	public property get count ''[int] Gets the number of items (<em>0</em> means no items).
		init()
		if datasource = "array" then
			count = uBound(data) + 1
		elseif datasource = "dictionary" then
			count = data.count
		elseif datasource = "recordset" then
			count = data.recordCount
		else
			notSupported("count")
		end if
	end property
	
	public property get datasource ''[string] Gets the type of the underlying data. ARRAY, DICTIONARY, RECORDSET
		datasource = p_datasource
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		instantiated = false
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	STATIC! Creates a new <em>DataContainer</em> instance for a given datasource
	'' @PARAM:			datasrc [array], [recordset], [dictionary]: datasource used for the container
	''					<strong>Note:</strong> Only keyset, dynamic or static RECORDSET cursor types are supported. Use <em>db.getUnlockedRS()</em> for it.
	'' @RETURN:			[DataContainer] NOTHING if the datasource is of a not supported type
	'**********************************************************************************************************
	public default function newWith(byRef datasrc)
		set newWith = nothing
		p_datasource = lCase(typename(datasrc))
		if isArray(datasrc) then
			p_datasource = "array"
			data = datasrc
		elseif p_datasource = "dictionary" then
			set data = datasrc
		elseif p_datasource = "recordset" then
			if datasrc.cursorType <= 0 then exit function
			set data = datasrc
		else
			exit function
		end if
		instantiated = true
		set newWith = me
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Removes all duplicates from the data. Changes the underlying <em>data</em>. Only for ARRAY yet.
	'' @DESCRIPTION:	- If not case sensitive then duplicate values will be represented by the first match
	''					- Order of data items stays the same
	'' @PARAM:			caseSensitive [bool]: should the uniquness be case sensitive.
	'' @RETURN:			[array] also returns the new generated array
	'***********************************************************************************************************
	public function unique(caseSensitive)
		init()
		if datasource <> "array" then notSupported("unique()")
		newData = array()
		if uBound(data) > -1 then newData = array(data(lBound(data)))
		for each oldD in data
			found = false
			for each newD in newData
				if (newD = oldD and caseSensitive) or (not caseSensitive and lcase(newD) = lcase(oldD)) then
					found = true
					exit for
				end if
			next
			if not found then
				redim preserve newData(uBound(newData) + 1)
				newData(uBound(newData)) = oldD
			end if
		next
		data = newData
		unique = data
	end function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	Paginates the data. Useful for e.g. paging bars
	'' @DESCRIPTION:	Pagination is performed in a way so the current page stays always in the middle.
	''					Example: Lets say we have 100 records and want to page them with 10 records on each page. The current selected page is 5
	''					and we only want to display 3 pages at one time (note: we take the current page from querystring as most pages implement it like this):
	''					<code>
	''					<%
	''					'first we get the current page from querystring and parse it so its always a number
	''					currentPage = str.parse(page.QS("page"), 0)
	''					pages = ((new DataContainer)(data)).paginate(10, currentPage, 3)
	''					'results in an array which contains the following values
	''					'... 4 5 6 ...
	''					% >
	''					</code>
	''					Now we could create a paging bar out of this array now (there are more solutions but this is the most common).
	''					Note: The "..." indicates that there are more pages. In this example there are more before 4th page and more after 6th page.
	''					<code>
	''					<%
	''					link = "<a href=""default.asp?page={0}"">{1}<a/>"
	''					str.write(str.format(link, array(1, "<<")))
	''					str.write(str.format(link, array(currentPage - 1, "< prev")))
	''					for each p in pages
	''					.	p = str.parse(p, 0)
	''					.	display = lib.iif(currentPage = p, "<strong>" & p & "</strong>", p)
	''					.	str.write(lib.iif(p > 0), str.format(link, array(p, display)), "...")
	''					next
	''					str.write(str.format(link, array(currentPage + 1, "next >")))
	''					str.write(str.format(link, array(empty, ">>")))
	''					% >
	''					</code>
	''					This code results in a paging which allows us navigating our data by going directly to a page,
	''					jumping to first/last page and moving to next/previous page. Moreover it also displays "..." if there are more pages
	''					and puts the currentpage in a strong-tag.
	'' @PARAM:			recsPerPage [int]: how many records are being displayed per page. Provide <em>0</em> if all.
	'' @PARAM:			currentPage [int]: holds the current active page.
	''					- only numbers greater 0 are recognized (otherwise 1 is used).
	''					- use EMPTY to set the current page to the last page
	''					- if the number is higher than the actual last page then its adjusted to the last page
	'' @PARAM:			numberOfPages [int]: indicates how many pages should be returned in total (incl. current page)
	''					number should be odd so the currentpage is always centered. If even then it will be rounded up to the next odd number.
	'' @RETURN:			[array] returns an array which contains the calculated page numbers in ascending order.
	''					- The amount of pages is always equal or or less then <em>numberOfPages</em>
	''					- It may contain a '...' marker on FIRST position to indicate there are more pages BEFORE those returned
	''					- It may contain a '...' marker on LAST position to indicate there are more pages AFTER those returned.
	''					- The array is never empty.
	''					- It never contains a 0 value
	''					Possible returns:
	''					<code>
	''					There are only 3 pages
	''					[1, 2, 3]
	''					There are pages before the 3rd page and pages after the 5th
	''					["...", 3, 4, 5, "..."]
	''					There are only page before or only after
	''					["...", 3, 4, 5]
	''					[1, 2, 3, "..."]
	''					only one page
	''					[1]
	''					</code>
	'***********************************************************************************************************
	public function paginate(recsPerPage, byVal currentPage, numberOfPages)
		init()
		paginate = array(1)
		pageCount = 0
	    if numberOfPages >= 2 then numberOfPages = int(numberOfPages / 2)
		if recsPerPage > 0 and count > 0 then pageCount = ceil(count, recsPerPage)
		if pageCount <= 1 then exit function
		
		if isEmpty(currentPage) or currentPage > pageCount then currentPage = pageCount
		if currentPage <= 0 then currentPage = 1
		
		firstPage = 1
		if currentPage > numberOfPages then
			cnt = 0
			if pageCount - currentPage <= numberOfPages then cnt = numberOfPages - abs(pageCount - currentPage)
			tmp = currentPage - numberOfPages - cnt
			if tmp >= 1 then firstPage = tmp
		end if
		
		lastPage = pageCount
		if currentPage + numberOfPages  <= pageCount then
			cnt = 0
			if currentPage <= numberOfPages then cnt = abs(currentPage - numberOfPages) + 1
			tmp = currentPage + numberOfPages + cnt
			if tmp < pageCount then lastPage = tmp
		end if
		
		arr = lib.iif(firstPage > 1, array("..."), array())
		for i = firstPage to lastPage
			redim preserve arr(uBound(arr) + 1)
			arr(uBound(arr)) = i
		next
		if lastPage < pageCount then
			redim preserve arr(uBound(arr) + 1)
			arr(uBound(arr)) = "..."
		end if
		paginate = arr
	end function
	
	'**********************************************************************************************************
	'* ceil 
	'**********************************************************************************************************
	private function ceil(dividend, divider)
	    if (dividend mod divider) = 0 Then
	   	 	ceil = dividend / divider
	    else
		    ceil = Int(dividend / divider) + 1
	    end if
    end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Checks if it contains a given value
	'' @DESCRIPTION:	DICTIONARY uses the key for comparison.
	'' @PARAM:			val [variant], [array]: The value which should be contained within the container. If an ARRAY is provided
	''					then all its values must be contained within the <em>DataContainer</em> in order to return TRUE.
	'' @RETURN:			[bool] TRUE if it contains the value otherwise FALSE
	'**********************************************************************************************************
	public function contains(val)
		init()
		contains = true
		if isArray(val) then
			for each v in val
				contains = contains(v)
				if not contains then exit function
			next
			exit function
		end if
		
		if datasource = "array" then
			for each d in data
				if d & "" = val & "" then exit function
			next
		elseif datasource = "dictionary" then
			for each k in data.keys
				if k & "" = val & "" then exit function
			next
		else
			lib.throwError("DataContainer.contains() does not support Recordset yet.")
		end if
		contains = false
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Returns a string representation of the data
	'' @DESCRIPTION:	<code>
	''					<%
	''					set d = (new DataContainer)(array(1, 2, 3))
	''					'=> 1 - 2 - 3
	''					str.write(d.toString(" - "))
	''					set d = (new DataContainer)(lib.newDict(empty))
	''					d.data.add 1, "foo"
	''					d.data.add 2, "some"
	''					'=> 1 foo, 2 some
	''					str.write(d.toString(", "))
	''					'=> 1 -> foo, 2 -> some
	''					str.write(d.toString(array(", ", " -> ")))
	''					% >
	''					</code>
	'' @PARAM:			delimiter [string], [array]: Use a STRING when the <em>data</em> is an ARRAY. In case of
	''					RECORDSET or DICTIONARY you may use an ARRAY to add an delimiter for columns (in case of a RECORDSET) or
	''					<em>key</em> and <em>item</em> (in case of a DICTIONARY). If a STRING is used on DICTIONARY or RECORDSET then
	''					the second delimiter is a whitespace by default.
	'**********************************************************************************************************
	public function toString(byVal delimiter)
		init()
		colDel = " "
		if datasource = "array" then
			if isArray(delimiter) then delimiter = delimiter(0)
			toString = str.arrayToString(data, delimiter)
		elseif datasource = "dictionary" then
			if isArray(delimiter) then
				if uBound(delimiter) < 1 then lib.throwError("DataContainer.toString() delimiter argument must contain at least 2 fields when its an array and used for Dictionary data.")
				lineDel = delimiter(0)
				colDel = delimiter(1)
			else
				lineDel = delimiter
			end if
			keys = data.keys
			for i = 0 to uBound(keys)
				if i > 0 then toString = toString & lineDel
				toString = toString & keys(i) & colDel & data(keys(i))
			next
		else
			notSupported("toString()")
		end if
	end function
	
	'**********************************************************************************************************
	'* init 
	'**********************************************************************************************************
	private sub init()
		if not instantiated then lib.throwError("Instances of DataContainer must be created using newWith() method.")
	end sub
	
	'**********************************************************************************************************
	'* getRecordsetBound 
	'**********************************************************************************************************
	private function getRecordsetBound(bound)
		getRecordsetBound = empty
		if data.recordCount = 0 then exit function
		pos = data.absolutePosition
		if pos = -1 then lib.throwError("Cannot determine recordsets " & bound & " bound because its positions is unknown.")
		if bound = "first" then data.moveFirst() else data.moveLast() end if
		getRecordsetBound = data.fields(0)
		'restore the original position of recordset
		if pos = -3 then 'EOF
			data.move 1, count
		elseif pos = -2 then 'BOF
			data.move -1, 1
		else 'any other position
			data.absolutePosition = pos
		end if
	end function
	
	'**********************************************************************************************************
	'* notSupported 
	'**********************************************************************************************************
	private sub notSupported(member)
		lib.throwError("DataContainer." & member & " is not supported by '" & datasource & "'")
	end sub

end class
%>