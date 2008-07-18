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
	
	'**************************************************************************************************************
	'' @SDESCRIPTION:	Gets geoinformation about the clients ISP according to its IP address like e.g. country, coordinates, etc.
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
	''					- <em>countryname</em>: Uppercased countryname. (EMPTY if unknown)
	''					- <em>lat</em>: Latitude (EMPTY if unknown)
	''					- <em>lng</em>: Longitude (EMPTY if unknown)
	''					All the different failures are handled as followed:
	''					- NOTHING is returned on a failure (xml parsing errors & network errors e.g. timeout, service down, etc.)
	''					- DICTIONARY with no keys (<em>count = 0</em>) is returned if the location is unknown (service is available but could not determine location).
	'**************************************************************************************************************
	public function getClientInfo(timeout)
		set getClientInfo = nothing
		ip = request.serverVariables("REMOTE_HOST")
	end function

end class
%>