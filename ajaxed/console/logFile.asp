<!--#include virtual="/ajaxed/ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-05-05 22:21
'* Description: views a given logfile
'* Input:		-
'******************************************************************************************

'we dont want any logs on of this page because of "watch logfile feature"
lib.logger.logLevel = 0
set page = new AjaxedPage
with page
	.plain = true
	.onlyDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	path = page.RF("file")
	if path = "" then exit sub
	path = server.mapPath(path)
	if not lib.fso.fileExists(path) then
		str.write("")
		exit sub
	end if
	set stream = lib.fso.openTextFile(path, 1, false)
	lineFeed = lib.iif(page.RF("moreSpace") = "1", "<br/><br/>", "<br/>")
	logsContent = replace(str(stream.readAll()), vbNewLine, lineFeed)
	if page.RF("removeAscii") = "1" then
		str.write(removeAsciiCodes(logsContent))
	else
		str.write(logsContent)
	end if
	stream.close()
	'delete the file if we watch it
	if page.RFHas("live") then lib.fso.deleteFile(path)
end sub

'******************************************************************************************
'* removeAsciiCodes 
'******************************************************************************************
function removeAsciiCodes(val)
	set r = new Regexp
	r.global = true
	r.pattern = chr(27) & "\[[\d;]{1,8}m"
	removeAsciiCodes = r.replace(val, "")
end function
%>