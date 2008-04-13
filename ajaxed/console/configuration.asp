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
		<% row "Ajaxed version", lib.version %>
		<% row "Environment", lib.ENV %>
		<% row "Scripting Engine", ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %>
		<% row "IIS", request.ServerVariables("SERVER_SOFTWARE") %>
		<% row "Server name", request.ServerVariables("SERVER_NAME") %>
		<% row "Components", str.parse(str.arrayToString(getComponents(), ", "), "-") %>
		<% row "Server time", now() %>
		</table>
	</div>
	<div style="float:left;width:60%">
		<table cellpadding="3">
		<% row "Default DB", lib.iif(isEmpty(db.defaultConnectionString), "no DB configured. (DB can be configured in your /ajaxedConfig/config.asp)", db.defaultConnectionString) %>
		<%
		if not isEmpty(db.defaultConnectionString) then
			on error resume next
				db.openDefault()
				failed = err <> 0
				if failed then description = err.description
			on error goto 0
			if failed then
				row "DB Connected", "could not connect. (" & description & ")"
			else
				row "DB Connected", "successfully.	"
				row "DBMS", db.connection.properties("DBMS Name") & " " & db.connection.properties("DBMS Version")
				row "Provider", db.connection.properties("Provider Name") & " " & db.connection.properties("Provider Version")
				row "OLE DB version", db.connection.properties("OLE DB Version")
			end if
		end if
		%>
		<% row "ADO version", server.createobject("adodb.connection").version %>
		</table>
	</div>
	<div style="clear:both">&nbsp;</div>

<% end sub %>

<% sub row(name, value) %>

	<tr valign="top">
		<td nowrap><strong><%= name %>:</strong></td>
		<td><%= str.HTMLEncode(value) %></td>
	</tr>
	
<% end sub %>