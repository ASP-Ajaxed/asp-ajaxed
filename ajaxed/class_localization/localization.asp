<%
'**************************************************************************************************************

'' @CLASSTITLE:		Localization
'' @CREATOR:		michal
'' @CREATEDON:		2008-07-16 11:18
'' @CDESCRIPTION:	Contains all stuff which has to do with Localization.
''					"Localization is the configuration that allows a program to be adaptable to local national-language features."
''					Also stuff about the client can be found in this class e.g. clients IP address (often needed to localize the user)
'' @REQUIRES:		-
'' @VERSION:		0.1
'' @STATICNAME:		local

'**************************************************************************************************************
class Localization

	public property get comma ''[char] Gets the char which represents the comma when using floating numbers. Returns either "," or "."
		comma = left(right(formatNumber(1.1, 1), 2), 1)
	end property
	
	public property get IP ''[string] gets the clients IP address (can also be the clients ISP IP).
		IP = request.serverVariables("REMOTE_ADDR")
	end property
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Checks if the client supports cookies
	'' @DESCRIPTION:	Basically a cookies is written to the client and tried to retrieve it. If cannot retrieve it then no cookies support is assumed.
	'' @RETURN:			[bool] TRUE if supports cookies otherwise FALSE
	'**************************************************************************************************************
	function supportsCookies()
		response.cookies("ajaxedTestCookie") = "1"
		response.cookies("ajaxedTestCookie").expires = dateAdd("yyyy", 1, now())
		supportsCookies = request.cookies("ajaxedTestCookie") = "1"
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Locates the country of a given clients IP address.
	'' @DESCRIPTION:	Currently Information is gathered from the free service at http://www.hostip.info and the clients IP is
	''					taken from the <em>serverVariables</em>. Note: As the location is the ISPs location its not sure that the
	''					client is located in the same location.
	''					- The location is cached within the users session. Therefore only the first call will access the service (unless its a failure).
	''					- Hint: For debuging you can check the logfile to see all the communication details with the service.
	''					Example of a correct call:
	''					<code>
	''					<%
	''					countryCode = local.locateClient(2, empty)
	''					if isEmpty(countryCode) then str.writeEnd("Service unavailable")
	''					if countryCode = "XX" then str.writeEnd("Location unknown")
	''					'if location was found then the country is available for sure
	''					str.writef("You are located in '{0}', arent you?", countryCode)
	''					% >
	''					</code>
	'' @PARAM:			timeout [int]: Timeout for the request in seconds. <em>0</em> mean as long as possible (not recommended!)
	'' @PARAM:			clientIP [string]: The IP you want to check. Provide EMPTY to check the clients IP (<em>Localization.IP</em>). Only if EMPTY is given then the query will be temporary cached in the session.
	'' @RETURN:			[string]
	''					- Returns the 2 letters upper-cased country code (if could determine) as given in ISO 3166-1.
	''					- Returns <em>XX</em> if the location is unknown (service is available but could not determine location). Also if the IP is private.
	''					- Returns EMPTY on a failure (xml parsing errors & network errors e.g. timeout, service down, etc.)
	'**************************************************************************************************************
	public function locateClient(timeout, clientIP)
		locateClient = empty
		'if its cached then take it and run...
		if isEmpty(clientIP) and session("ajaxed_locateClient") <> "" then
			locateClient = session("ajaxed_locateClient")
			lib.logger.debug "Got cached Localization.locateClient() from session."
			exit function
		end if
		set ixmlhttp = lib.requestURL("get", "http://api.hostip.info/", array("ip", lib.iif(isEmpty(clientIP), IP, clientIP)), empty)
		if ixmlhttp is nothing then exit function
		if ixmlhttp.status <> 200 then
			lib.logger.debug "Localization.geocodeClient() response was NOT successful: " & ixmlhttp.status
			exit function
		end if
		set xml = server.createObject("microsoft.xmldom")
		xml.loadXML(ixmlhttp.responseText)
		if xml.parseError.errorCode <> 0 then
			lib.logger.debug "Localization.geocodeClient() could not parse XML: " & xml.parseError.reason
			exit function
		end if
		locateClient = "XX"
		'now lets parse the XML. its specific to the provider (currently api.hostip.info)
		set n = xml.getElementsByTagName("countryAbbrev")
		if n.length > 0 then
			locateClient = uCase(n(0).text)
			'if the country abbreviation is XX then its an unknown address
			if len(locateClient) <> 2 then locateClient = "XX"
		end if
		set xml = nothing
		'cache the value if its for the clients IP
		if isEmpty(clientIP) then session("ajaxed_locateClient") = locateClient
	end function

end class
%>