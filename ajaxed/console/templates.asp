<!--#include file="../ajaxed.asp"-->
<!--#include file="../class_textTemplate/TextTemplate.asp"-->
<%
set page = new AjaxedPage
with page
	.onlyDev = true
	.plain = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* callback 
'******************************************************************************************
sub callback(action)
	set t = new TextTemplate
	t.fileName = page.RF("name")
	
	if action = "load" then
		page.returnValue "name", t.filename
		page.returnValue "content", t.content
		exit sub
	end if
	
	if not str.startsWith(t.filename, "/") or not str.endsWith(lCase(t.filename), ".template") then
		page.return "Template name must start with a '/' and has the extension '.template'"
		exit sub
	end if
	
	desc = ""
	select case action
		case "delete"
			on error resume next
				t.delete()
				desc = err.description
			on error goto 0
			page.return desc
		case "save"
			t.content = page.RF("template")
			on error resume next
				t.save()
				desc = err.description
			on error goto 0
			page.return desc
	end select
end sub

'******************************************************************************************
'* main 
'******************************************************************************************
sub main() %>

	<script>
		loaded = function(t) {
			$('templateName').value = t.name;
			$('templateContent').value = t.content;
		}
		done = function(r) {
			if (r == "") {
				alert("successfully done.");
				loadList();
			} else {
				alert(r);
			}
		}
		loadList = function() {
			new Ajax.Updater('templates', 'templatesList.asp');
		}
	</script>
	
	<form id="frm">
	
	Templates listed here are used with the <code>TextTemplate</code> component. 
	Best use for the templates is in combination with emails. It seperates email content from the actual code.
	Thus changing email content does not require touching the code anymore.
	<br><br>
	
	<div style="float:left;width:75%">
		<strong>Template:</strong> (virtual path &amp; filename)<br>
		<input type="Text" name="name" id="templateName" style="width:70%" value="" /><br>
		<br>
		<strong>Template content:</strong><br>
		<textarea class="code" wrap="off" rows="20" id="templateContent" style="width:90%" name="template"></textarea>
		<br><br>
		<button type="button" id="save" onclick="ajaxed.callback('save', done, null, null, 'templates.asp')">Save</button>
		<button type="button" id="delete" onclick="if(confirm('Are you sure you want to delete the selected Template?\nThis cannot be undone!'))ajaxed.callback('delete', done, null, null, 'templates.asp')">Delete</button>
	</div>
	<div style="float:left;width:25%">
		<strong>Existing Templates:</strong>
		<div id="templates"></div>
		<br>
	</div>
	<div style="clear:both">&nbsp;</div>
	
	</form>
	
	<% page.execJS("loadList();") %>
	
	
<% end sub %>