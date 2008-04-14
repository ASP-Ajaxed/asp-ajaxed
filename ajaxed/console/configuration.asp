<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-08 21:38
'* Description: configuration
'* Input:		-
'******************************************************************************************

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
	content()
end sub

'******************************************************************************************
'* getComponents 
'******************************************************************************************
function getComponents()
	'TODO: add more components here, which should be checked
	comps = array("JMail.Message", "Persits.MailSender", _
		"w3.Upload", "Persits.Jpeg", "Persits.Upload", "Persits.Pdf", "StringBuilderVB.StringBuilder")
	found = array()
	for each c in comps
		on error resume next
		set tmp = server.createObject(c)
		failed = err <> 0
		on error goto 0
		if not failed then
			redim preserve found(ubound(found) + 1)
			found(ubound(found)) = c
		end if
	next
	getComponents = found
end function

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>
	
	<div style="float:left;width:40%">
		<table cellpadding="3">
		<% row "Ajaxed version", lib.version & "&nbsp;<a href=""../changes.txt"">changes</a>", false %>
		<% row "Environment", lib.ENV, true %>
		<% row "Scripting Engine", ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion, false %>
		<% row "IIS", request.ServerVariables("SERVER_SOFTWARE"), false %>
		<% row "Server name", request.ServerVariables("SERVER_NAME"), false %>
		<% row "Components", str.parse(str.arrayToString(getComponents(), ", "), "-"), false %>
		<% row "Server time", now(), false %>
		</table>
	</div>
	<div style="float:left;width:60%">
		<table cellpadding="3">
		<% row "Default DB", lib.iif(isEmpty(db.defaultConnectionString), "no DB configured. (DB can be configured in your /ajaxedConfig/config.asp)", db.defaultConnectionString), false %>
		<%
		if not isEmpty(db.defaultConnectionString) then
			on error resume next
				db.openDefault()
				failed = err <> 0
				if failed then description = err.description
			on error goto 0
			if failed then
				row "DB Connected", "could not connect. (" & description & ")", false
			else
				row "DB Connected", "successfully.	", false
				row "DBMS", db.connection.properties("DBMS Name") & " " & db.connection.properties("DBMS Version"), false
				row "Provider", db.connection.properties("Provider Name") & " " & db.connection.properties("Provider Version"), false
				row "OLE DB version", db.connection.properties("OLE DB Version"), false
			end if
		end if
		%>
		<% row "ADO version", server.createobject("adodb.connection").version, false %>
		</table>
	</div>
	<div style="clear:both">&nbsp;</div>

<% end sub %>

<% sub row(name, value, htmlEncode) %>

	<tr valign="top">
		<td nowrap><strong><%= name %>:</strong></td>
		<td><%= lib.iif(htmlEncode, str.HTMLEncode(value), value) %></td>
	</tr>
	
<% end sub %>