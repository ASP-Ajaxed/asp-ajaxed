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

end class
%>