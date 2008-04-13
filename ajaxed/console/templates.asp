<!--#include file="../ajaxed.asp"-->
<!--#include virtual="/gab_Library/class_textTemplate/TextTemplate.asp"-->
<%
set page = new AjaxedPage
with page
	.onlyDev = true
	.plain = true
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
'* callback 
'******************************************************************************************
sub callback(action)
	set t = new TextTemplate
	t.fileName = page.RF("file")
	
	select case action
		case "selected"
			page.returnValue "content", t.content
			page.returnValue "name", t.filename
		case "delete"
			t.delete()
			page.return true
		case "save"
			t.filename = page.RF("name")
			t.content = page.RF("template")
			t.save()
			page.return true
	end select
end sub

'******************************************************************************************
'* sub  
'******************************************************************************************
sub drawFileSelector()
	set FS = new FileSelector
	with FS
		.height = "400"
		.multipleSelection = false
		.name = "file"
		.onItemClicked = "ajaxed.callback('selected', loadTemplate, $('frm').serialize(true))"
		.sourcePath = consts.userFiles & "emailTemplates/"
		.draw()
	end with
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<script>
		function loadTemplate(template) {
			$('name').value = template.name;
			$('template').value = template.content;
			
			$("name").readonly = true;
			$("save").disabled = false;
			$("delete").disabled = false;
		}
		function newT() {
			$("save").disabled = false;
			$("name").readOnly = false;
			$("delete").disabled = true;
			$("template").activate();
		}
		function saved(ret) {
			if (ret) {
				alert("Successfully saved template.");
				window.location.reload();
			}
		}
		function deleted(ret) {
			if (ret) {
				alert("Successfully deleted template.");
				window.location.reload();
			}
		}
	</script>
	
	<h1>Manage Templates</h1>

	<form class="form" id="frm">
	
	<div id="error"></div>
	<table>
	<tr valign="top">
		<td width="200">
			<fieldset style="width:auto">
				<legend>Templates</legend>
				<div class="content">
				<% drawFileSelector() %>
				<br>
				<button class="button" type="button" onclick="newT()">+ add new</button>
				</div>
			</fieldset>
		</td>
		<td>
			<table>
			<tr>
				<td class="label">Name:</td>
				<td><input type="Text" readonly name="name" size="50"></td>
			</tr>
			</table>
			<div><textarea class="code" wrap="off" rows="30" style="width:100%" name="template"></textarea></div>
			<div class="endline">
				<button type="button" class="button" id="save" disabled onclick="ajaxed.callback('save', saved, $('frm').serialize(true))">Save</button>
				<button type="button" class="button" id="delete" disabled onclick="if(confirm('Are you sure you want to delete the selected Template?\nCannot be undone!'))ajaxed.callback('delete', deleted, $('frm').serialize(true))">Delete</button>
				<% frm.drawPrintButton() %>
				<% frm.drawCancelButton("about:blank") %>
			</div>
		</td>
	</tr>
	</table>
	
	</form>

<% end sub %>