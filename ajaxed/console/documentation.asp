<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-08 21:38
'* Description: documentation
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
'* content 
'******************************************************************************************
sub content() %>

	<a href="../index.html" target="_blank"><strong>ajaxed API</strong></a>
	full programmers class reference.
	<br><br>
	<a href="http://groups.google.com/group/asp-ajaxed" target="_blank">
		<strong>asp-ajaxed google discussion group</strong></a>
	discuss issues with other ajaxed developers and contributors.
	<br><br>
	<a href="http://www.webdevbros.net/category/classic-asp/ajaxed/" target="_blank">
		<strong>Tutorials</strong></a>
	browse through different ajaxed tutorials on webdevbros.
	
	<br><br><br>
	
	<div style="float:left;width:50%">
		<strong>Want to generate your own documentation?</strong>
		<br>
		Enter the virtual path of the folder (on your server) you want to create the documentation for.
		<br>
		If it's not there copy it to your webroot ;)<br>
		<br>
		<script>
			genDoc = function() {
				new Ajax.Updater('docOutput', 'documentor/default.asp', {
					parameters: {folder: $F('docpath')}
				})
			}
		</script>
		<label>Documentation of:</label>
		<input type="Text" id="docpath" size="20" value="/ajaxed/">
		<button type="button" onclick="genDoc()">generate...</button>
		
		<br><br>
		<a href="documentor/manual.asp" target="_blank">
			<strong>Documentor manual</strong></a>
		read through a quick manual how to document your code.
	</div>
	<div style="float:left;width:50%">
		<div class="console" id="docOutput">&gt; Ready.</div>
	</div>
	
	<div style="clear:both">&nbsp;</div>

<% end sub %>
