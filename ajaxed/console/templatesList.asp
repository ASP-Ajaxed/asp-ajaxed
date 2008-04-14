<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-14 19:19
'* Description: lists all templates on the server
'* Input:		-
'******************************************************************************************

numberTemplates = 0
set page = new AjaxedPage
with page
	.onlyDev = true
	.plain = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* drawFiles  
'******************************************************************************************
function drawFiles(folder, caption)
	if lcase(folder.name) = "ajaxed" then exit function
	for each file in folder.files
		if str.endsWith(file.name, ".template") then
			path = unmapPath(file.path, "/")
			str.writeln(str.format(caption, array(path, str.shorten(file.name, 25, "..."))))
			numberTemplates = numberTemplates + 1
		end if
	next
	for each subfolder in folder.subfolders
		drawFiles subFolder, caption
	next
end function

'******************************************************************************************
'* unMapPath 
'******************************************************************************************
function unMapPath(aPath, virtualRoot)
	unMapPath = replace(replace(lCase(aPath), lCase(server.mappath(virtualRoot)), ""), "\", "/")
end function

'******************************************************************************************
'* main 
'******************************************************************************************
sub main() %>

	<% c = "<a href=""javascript:void(0)"" onclick=""ajaxed.callback('load', loaded, {name:'{0}'}, null, 'templates.asp')"" title=""{0}"">{1}</a>" %>
	<% drawFiles lib.fso.getFolder(server.mappath("/")), c %>
	
	<br><br>
	<%= numberTemplates %> templates found.

<% end sub %>