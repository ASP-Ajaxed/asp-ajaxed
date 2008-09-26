<!--#include file="class_RSSItem.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		RSS
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-09-08 15:07
'' @CDESCRIPTION:	Reads RSS feeds and gives you the possibility of formating it with xsl or using
''					the items programmatically as objects. Additionally it lets you generate your own feeds.
''					- READING: achieved with <em>load()</em> or <em>draw()</em>. Refer to the methods for the supported versions
''					- WRITING: achiedved with <em>generate()</em>. Refer to the method for the supported versions
''					If caching should be enabled use the <em>setCache()</em> method. By default caching is off.
''					- on <em>draw()</em> caching stores the complete transformed output in the cache
''					- on <em>load()</em> caching stores the xml in the cache and parses it out from the cache.
''					Simple example of how to read an RSS feed:
''					<code>
''					<%
''					set r = new RSS
''					r.url = "http://somehost.com/somefeed.xml"
''					r.load()
''					if r.failed then lib.error("could not read feed. could be down or wrong format")
''					for each it in r.items
''					.	str.write(r.title & "<br>")
''					next
''					% >
''					</code>
'' @REQUIRES:		Cache
'' @POSTFIX:		rss
'' @VERSION:		1.0

'**************************************************************************************************************
class RSS

	private xml, generator, xmlDomVersion
	
	'public members
	public url				''[string] url of the RSS feed (necessary for <em>load()</em> and <em>draw()</em>)
	public failed			''[bool] indicates if any of the methods <em>load()</em> or <em>draw()</em>
							''e.g. server not responding, timeouts, etc.
	public title			''[string] title of the feed
	public link				''[string] link which is associated with the feed
	public description		''[string] description of the feed
	public language			''[string] the language of the feed. e.g. <em>en</em>
	public publishedDate	''[date] date when the feed has been published. you local timezone!
	public items			''[dictionary] collection of the items (<em>RSSItem</em>). after <em>load()</em> this gets filled
							''- for writing a feed this collection needs to be filled with items which sould be published. Use <em>addItem()</em>
	public theCache			''[Cache] holds the cache which is used for caching the feed. by default it is nothing and
							''should only be set by <em>setCache()</em>. afterwards this property can be used.
	public debug			''[bool] turns debuging on/off. if on then useful information will be shown which can be used
							''if the feed never loads. e.g. the response of the server, etc. default = false
	public timeout			''[int] timeout in seconds for the requests. default = 2
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set items = server.createObject("scripting.dictionary")
		set theCache = nothing
		failed = false
		debug = false
		timeout = 2
		generator = "asp ajaxed RSS Component v1.0"
		xmlDomVersion = "Microsoft.XMLDOM" '"Msxml2.DOMDocument.4.0"
	end sub
	
	'**********************************************************************************************************
	'* destructor 
	'**********************************************************************************************************
	public sub class_terminate()
		set xml = nothing
		set theCache = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	draws the RSS feed with a given XSL. if cached then it will be taken from the cache
	'' @DESCRIPTION:	The whole HTML is taken from the cache and not just the data. so tranformation
	''					is performed just once.
	''					- supported formats: all because the XSL needs to be provided by the client manually
	''					- check failed property after execution
	'' @PARAM:			xslPath [string]: virtual path to the xsl. e.g. <em>/xsl/foo.xsl</em>
	'**********************************************************************************************************
	public sub draw(xslPath)
		initialize("reading")
		'check if we can get it from cache
		if not theCache is nothing then
			cacheID = url & "|DRAW"
			cachedItem = theCache.getItem(cacheID)
			if cachedItem <> "" then
				str.write(cachedItem)
				exit sub
			end if
		end if
		
		'when this is reached then whether caching is off or nothing found in the cache
		'then we get it again
		loadXML(getXMLHTTPResponse())
		if failed then exit sub
		set xsl = server.createObject(xmlDomVersion)
		xsl.async = false
		xsl.load(server.mappath(xslPath))
		contentTranformed = xml.transformNode(xsl)
		set xsl = nothing
		
		'if caching is on then we need to cache it.
		if not theCache is nothing then theCache.store cacheID, contentTranformed
		
		'last but not least draw the content
		str.write(contentTranformed)
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	loads the feed. so the items, title, link, etc. will be set
	'' @DESCRIPTION:	use this if you want to access e.g. the items of the feed and dont want to use
	''					the <em>draw()</em> method using the XSL. Note: using this method wont use the caching!
	''					- supported formats: RSS1.0 (RDF), RSS2.0, RSS0.92, RSS0.94 and ATOM
	''					- check failed property after execution
	'**********************************************************************************************************
	public sub load()
		initialize("reading")
		
		'first check if we can get it from cache
		if theCache is nothing then
			loadXML(getXMLHTTPResponse())
		else
			cacheID = url & "|LOAD"
			cachedItem = theCache.getItem(cacheID)
			if cachedItem <> "" then
				'the cache for this method stores the data as xml and needs to be parsed again
				loadXML(cachedItem)
			else
				'load and save to cache
				xmlResponse = getXMLHTTPResponse()
				loadXML(xmlResponse)
				theCache.store cacheID, xmlResponse
			end if
		end if
		
		if failed then exit sub
		
		select case getRSSType()
			case "RSS1.0"
				readRSS("1.0")
			case "RSS2.0"
				readRSS("2.0")
			case "ATOM"
				readATOM()
		end select
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	generates a feed with the given items and the metadata (title, etc.). Returns the
	''					generated XMLDOM and provides the possibility to save it into a file immediately
	'' @DESCRIPTION:	Feed is generated with UTF-8 (Note: your file must be saved as UTF-8 in order to support this).
	''					there are two ways to generate the feed:<br><br>
	''					<strong>1.</strong> Using a file which will be stored on your server (by providing a target param)
	''					(normally used if you have an action in your application which invokes the feed
	''					generation. e.g. new post in a blog system)<br><br>
	''					<strong>2.</strong> Getting the XML and using it as desired. e.g. sending to the response (no target param)
	''					(normally used if the feed is generated on each feed request - so the feed is an ASP file
	''					itself and is generated on each request.)<br><br>
	''					Example (2nd approach) of how to create a feed on your site (the example creates the feed everytime a feed reader accesses the page)
	''					which gets the news from a <em>news</em> database table:
	''					<code>
	''					<%
	''					set r = new RSS
	''					r.title = "My new ajaxed feed"
	''					r.description = "important stuff"
	''					r.language = "en"
	''					r.link = "http://mysite.com"
	''					
	''					set RS = db.getRS("SELECT * FROM news ORDER BY published_on DESC LIMIT 10", empty)
	''					while not RS.eof
	''					.	if isEmpty(pubDate) then pubDate = cDate(RS("published_on"))
	''					.	set it = new RSSItem
	''					.	it.author = RS("author")
	''					.	it.description = RS("excerpt")
	''					.	it.link = "http://mysite.com/news/?" & RS("id")
	''					.	it.guid = it.link
	''					.	it.title = RS("title")
	''					.	it.publishedDate = cDate(RS("published_on"))
	''					.	r.addItem(it)
	''					.	RS.movenext()
	''					wend
	''					if isEmpty(pubDate) then pubDate = now()
	''					r.publishedDate = pubDate
	''					set xml = r.generate("RSS2.0", empty)
	''					xml.save(response)
	''					% >
	''					</code>
	'' @PARAM:			format [string]: Currently only the option <em>RSS2.0</em> is possible
	'' @PARAM:			target [string]: path to the file you want to save the feed to. e.g. <em>/feeds/feed.xml</em>
	''					leave EMPTY if you want to get the XML returned
	'' @RETURN:			[microsoft.xmldom] the resulted xmldom
	'**********************************************************************************************************
	public function generate(format, target)
		initialize("writing")
		if isEmpty(title) or isEmpty(link) or isEmpty(description) then
			lib.throwError("Title, Link and description are required.")
		end if
		
		select case uCase(format)
			case "RSS2.0"
				set RSSNode = generateRSS20()
			case else
				lib.throwError("Unknown (or not supported) feed format.")
		end select
		
		xml.appendChild(RSSNode)
		set PINF = xml.createProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
		xml.insertBefore PINF, xml.childNodes(0)
		if not isEmpty(target) then xml.save(server.mapPath(target))
		if lib.logger.logsOnLevel(1) then lib.logger.debug xml.xml
		set generate = xml
	end function
	
	'**********************************************************************************************************
	'* generateRSS20 
	'**********************************************************************************************************
	private function generateRSS20()
		set RSSNode = getNewNode("rss", empty)
		RSSNode.setAttribute "version", "2.0"
		set cNode = getNewNode("channel", empty)
		with cNode
			.appendChild(getNewNode("title", title))
			.appendChild(getNewNode("link", link))
			.appendChild(getNewNode("description", description))
			if not isEmpty(language) then cNode.appendChild(getNewNode("language", language))
			.appendChild(getNewNode("generator", generator))
			if not isEmpty(publishedDate) then .appendChild(getNewNode("pubDate", formatDate(publishedDate)))
			'.appendChild(getNewNode("lastBuildDate", formatDate(now())))
		end with
		
		for each i in items.items
			with i
				set iNode = getNewNode("item", empty)
				if isEmpty(.title) and isEmpty(.description) then lib.throwError("Title or description of an RSS item must be available.")
				if not isEmpty(.title) then iNode.appendChild(getNewNode("title", .title))
				if not isEmpty(.link) then iNode.appendChild(getNewNode("link", .link))
				if not isEmpty(.author) then iNode.appendChild(getNewNode("author", .author))
				if not isEmpty(.description) then iNode.appendChild(getNewNode("description", array(.description)))
				if not isEmpty(.category) then iNode.appendChild(getNewNode("category", array(.category)))
				if not isEmpty(.GUID) then iNode.appendChild(getNewNode("guid", .GUID))
				if not isEmpty(.publishedDate) then iNode.appendChild(getNewNode("pubDate", formatDate(.publishedDate)))
				'add the item to the channel node
				cNode.appendChild(iNode)
			end with
		next
		RSSNode.appendChild(cNode)
		set generateRSS20 = RSSNode
	end function
	
	'**********************************************************************************************************
	'* getNewNode - if value is array then value is treated as CDATA
	'**********************************************************************************************************
	private function getNewNode(name, value)
		set getNewNode = xml.createElement(name)
		if isArray(value) then
			getNewNode.appendChild(xml.createCDATASection(value(0)))
		else
			getNewNode.appendChild(xml.createTextNode(value))
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	adds an item to the items collection
	'' @PARAM:			rItem [RSSItem]: the item you want to add
	'**********************************************************************************************************
	public sub addItem(rItem)
		items.add items.count + 1, rItem
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	reflection of the properties and its values
	'' @RETURN:			[dictionary] <em>key</em> = property-name, <em>value</em> = property value
	'**********************************************************************************************************
	public function reflect()
		set reflect = server.createObject("scripting.dictionary")
		with reflect
			.add "url", url
			.add "failed", failed
			.add "title", title
			.add "link", link
			.add "description", description
			.add "language", language
			.add "publishedDate", publishedDate
			arr = array()
			itemsArr = items.items
			redim arr(uBound(itemsArr))
			for i = 0 to uBound(arr)
				set arr(i) = itemsArr(i).reflect()
			next
			.add "items", arr
			.add "debug", debug
			.add "timeout", timeout
		end with
	end function
	
	'**********************************************************************************************************
	'* readRSS 
	'**********************************************************************************************************
	private function readRSS(version)
		set root = xml.documentElement
		title = getNodeText(root, "channel/title")
		link = getNodeText(root, "channel/link")
		description = getNodeText(root, "channel/description")
		language = getNodeText(root, "channel/language")
		publishedDate = parseDate(getNodeText(root, array("channel/pubDate", "channel/dc:date")))
		
		if version = "1.0" then
			set itemNodes = root.selectNodes("item")
		else
			set itemNodes = root.selectNodes("channel/item")
		end if
		for each n in itemNodes
			set item = new RSSItem
			with item
				.title = getNodeText(n, "title")
				.link = getNodeText(n, "link")
				.author = getNodeText(n, array("author", "dc:creator"))
				.description = getNodeText(n, "description")
				.category = getNodeText(n, array("category", "dc:subject"))
				.GUID = getNodeText(n, "guid")
				.publishedDate = parseDate(getNodeText(n, array("pubDate", "dc:date")))
			end with
			addItem(item)
		next
	end function
	
	'**********************************************************************************************************
	'* readATOM 
	'**********************************************************************************************************
	private function readATOM()
		set root = xml.documentElement
		title = getNodeText(root, "title")
		description = getNodeText(root, array("subtitle", "tagline"))
		link = root.selectSingleNode("link").getAttribute("href")
		publishedDate = parseDate(getNodeText(root, array("updated", "modified")))
		set itemNodes = root.selectNodes("entry")
		i = 0
		for each n in itemNodes
			set item = new RSSItem
			with item
				.title = getNodeText(n, "title")
				.link = getAttribute(n, "link", "href")
				.author = getNodeText(n, "author/name")
				.category = getAttribute(n, "category", "term")
				'atom has summary and/or content. we try first summary and then content if summary is not here
				.description = getNodeText(n, "summary")
				if .description = "" then .description = getNodeText(n, "content")
				.publishedDate = parseDate(getNodeText(n, array("published", "issued", "modified")))
			end with
			items.add i, item
			i = i + 1
		next
	end function
	
	'**********************************************************************************************************
	'* gets an attribute of a node only if the node exists
	'**********************************************************************************************************
	private function getAttribute(sourceNode, xpathQuery, attribute)
		getAttribute = ""
		set aNode = sourceNode.selectSingleNode(xpathQuery)
		if not aNode is nothing then getAttribute = aNode.getAttribute(attribute)
		set aNode = nothing
	end function
	
	'**********************************************************************************************************
	'* selectSingleNode - the same as from the xmldom, with exception that it returns an empty string if
	'* node cannot be found
	'* - if xpathQuery is an ARRAY then every node will be checked for existence (if previous has no value)
	'**********************************************************************************************************
	private function getNodeText(sourceNode, xpathQuery)
		getNodeText = ""
		if isArray(xpathQuery) then
			for each q in xpathQuery
				getNodeText = getNodeText(sourceNode, q)
				if getNodeText <> "" then exit for
			next
		else
			set aNode = sourceNode.selectSingleNode(xpathQuery)
			if not aNode is nothing then getNodeText = aNode.text
			set aNode = nothing
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	sets the caching of the RSS. the Cache class needs to be loaded for this
	'' @DESCRIPTION:	Specify how long the feed should be cashed till it will be requested again. For more
	''					information about the caching look up the <em>Cache</em> class
	'' @PARAM:			interval [string]: <em>m</em> = months, <em>y</em> = year, <em>d</em> = day, etc. everything accepted which is supported by <em>dateadd</em> function
	'' @PARAM:			value [int]: value for the interval. e.g. 1 (month)
	'**********************************************************************************************************
	public sub setCache(interval, value)
		if theCache is nothing then
			lib.require "Cache", "RSS"
			set theCache = new Cache
			theCache.name = "RSS"
		end if
		theCache.interval = interval
		theCache.intervalValue = value
	end sub
	
	'**********************************************************************************************************
	'* loadXML - tries to load an xmldom with a given xml-string. <em>failed</em> is set to TRUE if parse failed
	'**********************************************************************************************************
	private sub loadXML(xmlString)
		if failed then exit sub
		if not xml.loadxml(xmlString) then
			failed = true
			exit sub
		end if
		if debug then dWrite("ResponseText:" & server.HTMLEncode(xmlString) & "<br>")
		if xml.parseError.errorCode <> 0 then
			failed = true
			if debug then
				dWrite("Parseerror Reason: " & xml.parseError.reason & "<br>")
				dWrite("Parseerror Position: line " & xml.parseError.line & " (position " & xml.parseError.linePos & ")<br>")
				dWrite("Parseerror Source: " & server.HTMLEncode(xml.parseError.srcText) & "<br>")
			end if
		end if
	end sub
	
	'**********************************************************************************************************
	'* getXMLHTTPResponse - returns the response text of an XML file (based on the url). sets failed to true if failed
	'**********************************************************************************************************
	private function getXMLHTTPResponse()
		if failed then exit function
		tout = timeout * 1000
		on error resume next
		set xmlhttp = server.createObject("Msxml2.ServerXMLHTTP.3.0")
		with xmlhttp
			.setTimeouts tout, tout, tout, tout 'resolve, connect, send, receive
			.open "GET", url, false
			.send()
			getXMLHTTPResponse = .responseText
		end with
		set xmlhhtp = nothing
		if err <> 0 then failed = true
		on error goto 0
	end function
	
	'**********************************************************************************************************
	'* initialize - for load(), draw() and generate()
	'* mode = reading, writing
	'**********************************************************************************************************
	private sub initialize(mode)
		if mode = "reading" then
			if url = "" then lib.throwError("No URL given. When reading a feed it is necessary to set the URL property")
		end if
		set xml = server.createObject(xmlDomVersion)
		namespaces = "xmlns:dc='http://purl.org/dc/elements/1.1/'"
		xml.setProperty "SelectionNamespaces", namespaces
		xml.async = false
		failed = false
	end sub
	
	'******************************************************************************************
	'* formatDate - (wrapper for JScript function)
	'******************************************************************************************
	private function formatDate(aDate)
		d = cDate(aDate)
		formatDate = toUTC(year(d), month(d) - 1, day(d), hour(d), minute(d), second(d))
	end function
	
	'******************************************************************************************
	'* parseDate - parses a date string into a VBScript date
	'******************************************************************************************
	private function parseDate(dateStr)
		parseDate = dateStr
		if instr(trim(dateStr), " ") > 0 then
			'UTC Date: Thu, 28 Jun 2007 18:16:37 +0000
			if len(dateStr) < 15 then exit function
			parts = split(fromUTC(dateStr), ":")
			parseDate = newDate(parts(0), parts(1) + 1, parts(2), parts(3), parts(4), parts(5))
		else
			'format: 2007-05-03T16:57:21.875+02:00 or 2007-06-28T18:16:37Z
			if len(dateStr) < 19 then exit function
			stripped = split(dateStr, ".")
			stripped(0) = replace(stripped(0), "Z", "")
			parts = split(stripped(0), "T")
			dat = split(parts(0), "-")
			tim = split(parts(1), ":")
			parseDate = newDate(dat(0), dat(1), dat(2), tim(0), tim(1), tim(2))
		end if
	end function
	
	'**********************************************************************************************************
	'* newDate 
	'**********************************************************************************************************
	private function newDate(yyyy, mm, dd, h, m, s)
		datStr = dateSerial(yyyy, mm, dd) & " " & h & ":" & m & ":" & s
		newDate = cDate(datStr)
	end function
	
	'**********************************************************************************************************
	'* getRSSType 
	'**********************************************************************************************************
	private function getRSSType()
		getRSSType = empty
		select case uCase(xml.documentElement.baseName)
			case "RDF"
				getRSSType = "RSS1.0"
			case "RSS"
				ver = xml.documentElement.getAttribute("version")
				if ver = "2.0" or ver = "0.94" or ver = "0.93" or ver = "0.92" or ver = "0.91" then getRSSType = "RSS2.0"
			case "FEED"
				getRSSType = "ATOM"
		end select
		if getRSSType = empty then
			lib.throwError("Unrecognized format. Only RSS 1.0, RSS 2.0 and Atom supported.")
		else
			if debug then dWrite("Recognized RSS version: " & getRSSType & "<br>")
		end if
	end function
	
	'**********************************************************************************************************
	'* dWrite 
	'**********************************************************************************************************
	private function dWrite(msg)
		str.write(msg)
	end function

end class
%>
<script language="JScript" runat="server">
	//those functions are here because i could not found any equivalent for the
	//UTC conversion within VBScript.. at least not with that less code ;)
	function toUTC(y, mo, d, h, m, s) {
		return (new Date(y, mo, d, h, m, s).toUTCString());
	}
	function fromUTC(dateString) {
		d = new Date(dateString);
		return (d.getFullYear() + ':' + d.getMonth() + ':' + d.getDate() + ':' + d.getHours() + ':' + d.getMinutes() + ':' + d.getSeconds());
	}
</script>