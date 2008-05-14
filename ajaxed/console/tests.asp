<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-08 21:33
'* Description: lists all tests
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
'* prints test files
'* - the ajaxedFolder param is here because virtual folders are not recognized
'* when walking through the FS with FSO. thats why it needs to be done manually.
'* Its only done for the ajaxed folder because this is likely to be virtual on some 
'* environments.
'******************************************************************************************
function drawFiles(folder, caption, ajaxedFolder)
	if lcase(folder.name) = "ajaxed" and not ajaxedFolder then exit function
	for each file in folder.files
		if str.startsWith(file.name, "test_") and str.endsWith(file.name, ".asp") then
			if ajaxedFolder then
				path = "/ajaxed" & unmapPath(file.path, "/ajaxed/")
			else
				path = unmapPath(file.path, "/")
			end if
			str.writeln(str.format(caption, array(path, str.shorten(file.name, 25, "..."))))
		end if
	next
	for each subfolder in folder.subfolders
		drawFiles subFolder, caption, ajaxedFolder
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
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<script>
		runAll = function() {
			$('output').update('');
			$$('#tests a').each(function(el) {
				run(el.title);
			});
		}
		run = function(file) {
			new Ajax.Updater('output', file, {
				insertion: Insertion.Bottom,
				onSuccess: function(t) {
					t.responseText = t.responseText.gsub(/(\([^0][\d| ]* (errors|failed)\))/, function(match) {
						return '<span class="error">' + match[1] + '</span>'
					}) + '<br>';
				},
				onComplete: function(t) {
					scrollToFloor('console');
				}
			});
		}
		runSelected = function() {
			$$('#tests .selected').each(function(el) {
				if (el.checked) run(el.value);
			});
		}
	</script>
	
	Here you see all unit tests located on your server. A test is recognized if the file
	is prefixed with <code>test_</code> and has the extension <code>.asp</code>. The console
	indicates test errors and failures with a red background color. 
	Take a look at the documentation of the <code>TestFixture</code> class if you want to write
	your own tests. 
	<br><br>
	
	<div style="float:left;width:30%">
		<div style="height:300px;overflow:auto;padding:5px;" id="tests">
			<%
			link = "<input type=""Checkbox"" class=""selected"" value=""{0}"">&nbsp;" & _
					"<a href=""javascript:void(0)"" title=""{0}"" onclick=""run(this.title)"">{1}</a>&nbsp;&nbsp;" & _
					"<br>"
			%>
			<div id="ajaxedTests">
				<strong>
					<input type="Checkbox" onclick="$$('#ajaxedTests .selected').each(function(el){el.checked = !el.checked})">
					Ajaxed tests:
				</strong><br>
				<% drawFiles lib.fso.getFolder(server.mappath("/ajaxed/")), link, true %>
			</div>
			<br>
			<div id="otherTests">
				<strong>
					<input type="Checkbox" onclick="$$('#otherTests .selected').each(function(el){el.checked = !el.checked})">
					Other tests:
				</strong><br>
				<% drawFiles lib.fso.getFolder(server.mappath("/")), link, false %>
			</div>
		</div>
		<button type="button" onclick="runSelected()">run selected</button>
		<button type="button" onclick="runAll()">run all</button>
	</div>
	<div style="float:left;width:70%">
		<div class="console" id="console" style="height:300px;">
			<div>&gt; Ready.</div>
			<div id="output"></div>
		</div>
		<a href="javascript:void(0)" onclick="$('output').update('');">clear console</a>
	</div>
	<div style="clear:both">&nbsp;</div>

<% end sub %>