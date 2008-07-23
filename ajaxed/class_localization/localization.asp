<%
'**************************************************************************************************************

'' @CLASSTITLE:		Localization
'' @CREATOR:		michal
'' @CREATEDON:		2008-07-16 11:18
'' @CDESCRIPTION:	Contains all stuff which has to do with Localization.
''					"Localization is the configuration that allows a program to be adaptable to local national-language features."
'' @REQUIRES:		-
'' @VERSION:		0.1
'' @STATICNAME:		local

'**************************************************************************************************************
class Localization

	public property get comma ''[char] Gets the char which represents the comma when using floating numbers. Returns either "," or "."
		comma = left(right(formatNumber(1.1, 1), 2), 1)
	end property
	
	public property get IP ''[string] gets the clients IP address (can also be the clients ISP IP).
		IP = request.serverVariables("REMOTE_HOST")
	end property
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Gets geo information about the clients ISP according to its IP address like e.g. country, coordinates, etc.
	'' @DESCRIPTION:	Information is gathered from the free service at http://www.hostip.info and the clients IP is
	''					taken from the <em>serverVariables</em>. Note: As the location is the ISPs location its not sure that the
	''					client is located in the same location.
	''					- The location is cached within the users session. Therefore only the first call will access the service (unless its a failure).
	''					- Hint: For debuging you can check the logfile to see all the communication details with the service.
	''					Example of a correct call:
	''					<code>
	''					<%
	''					set gInfo = local.getClientInfo(2)
	''					if gInfo is nothing then str.writeEnd("Service unavailable")
	''					if gInfo.count = 0 thn str.writeEnd("Location unknown")
	''					'if location was found then the country is available for sure
	''					str.writef("You are located in '{0}', arent you?", gInfo("country"))
	''					% >
	''					</code>
	'' @PARAM:			timeout [int]: Timeout for the request in seconds. <em>0</em> mean as long as possible (not recommended!)
	'' @RETURN:			[Dictionary] Returns a DICTIONARY with the following lowercased keys:
	''					- <em>country</em>: 2-letter uppercased countrycode
	''					- <em>lat</em>: Latitude of the location (EMPTY if unknown)
	''					- <em>lng</em>: Longitude of the location (EMPTY if unknown)
	''					All the different failures are handled as followed:
	''					- NOTHING is returned on a failure (xml parsing errors & network errors e.g. timeout, service down, etc.)
	''					- DICTIONARY with no keys (<em>count = 0</em>) is returned if the location is unknown (service is available but could not determine location).
	'**************************************************************************************************************
	public function geocodeClient(timeout)
		set geocodeClient = nothing
		set ixmlhttp = lib.requestURL("get", "http://api.hostip.info/", _
			array("ip", IP), empty)
		if ixmlhttp is nothing then exit function
		if ixmlhttp.status <> 200 then
			lib.logger.debug "Localization.geocodeClient() response was not a success: " & ixmlhttp.status
			exit function
		end if
		set xml = server.createObject("microsoft.xmldom")
		xml.loadXML(ixmlhttp.responseText)
		if xml.parseError.errorCode <> 0 then
			lib.logger.debug "Localization.geocodeClient() could not parse XML: " & xml.parseError.reason
			exit function
		end if
		set geocodeClient = ["D"](empty)
		set n = xml.getElementsByTagName("countryAbbrev")
		if n.length > 0 then
			country = uCase(n(0).text)
			'if the country abbreviation is XX then its an unknown address
			if country = "XX" then exit function
			geocodeClient.add "country", uCase(n(0).text)
			set n = xml.getElementsByTagName("gml:coordinates")
			if n.length > 0 then
				latlng = split(n(0).text, ",")
				old = setLocale("en-us")
				geocodeClient.add "lng", str.parse(latlng(0), 0.0)
				geocodeClient.add "lat", str.parse(latlng(1), 0.0)
				setLocale(old)
			else
				geocodeClient.add "lng", empty
				geocodeClient.add "lat", empty
			end if
		end if
	end function

end class
%>