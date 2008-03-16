<!--#include file="../ajaxed.asp"-->
<%
set fso = server.createObject("scripting.filesystemobject")
for each file in fso.getFolder(server.mappath("/ajaxed/test/")).files
	if str.startsWith(file.name, "test_") then
		str.write("<a href=/ajaxed/test/" & file.name & ">" & file.name & "</a><br>")
	end if
next
%>