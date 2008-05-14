<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-08 21:38
'* Description: logs
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
'* callback 
'******************************************************************************************
sub callback(a)
	if a = "clear" then
		lib.logger.clearLogs()
		page.return(true)
	end if
end sub

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
	
	Ajaxed offers the possibility of easy logging. The <code>lib.loggger</code> property
	holds a ready-to-use <code>Logger</code> instance. All logfiles are stored in
	<code><%= lib.logger.path %></code> named after the current environment.
	
	<br><br>
	
	<script>
		loadFile = function() {
			if ($('live').checked) {
				if (watcher) {
					watcher.stop();
					watcher = null;
				}
				watcher = new Ajax.PeriodicalUpdater('output', 'logFile.asp', {
					parameters: $('frm').serialize(true),
					insertion: Insertion.Bottom,
					onSuccess: function(r) {
						scrollToFloor('console');
					}
				});
			} else {
				if (watcher) {
					watcher.stop();
					watcher = null;
					return;
				}
				new Ajax.Updater('output', 'logFile.asp', {
					parameters: $('frm').serialize(true),
					onComplete: function(t) {
						scrollToFloor('console');
					}
				});
			}
		}
	</script>
	
	<form id="frm">
	<div style="float:left;width:25%">
		<div style="height:230px;overflow:auto;padding:5px;" id="logs">
			<% if lib.fso.folderExists(server.mapPath(lib.logger.path)) then %>
				<% for each f in lib.fso.getFolder(server.mappath(lib.logger.path)).files %>
					<div>
						<% id = lib.getUniqueID() %>
						<input type="Radio" name="file" id="f<%= id %>" value="<%= lib.logger.path & f.name %>" onclick="loadFile()">
						<label for="f<%= id %>"><a href="javascript:void(0)"><%= f.name %></a></label>
					</div>
				<% next %>
			<% else %>
				<em>No log files found.</em>
			<% end if %>
		</div>
		<input type="Checkbox" onchange="loadFile()" name="live" value="1" id="live">
		<label for="live">live watch (BETA!)</label>
		<br>
		<input type="Checkbox" onchange="loadFile()" checked value="1" name="removeAscii" id="ascii">
		<label for="ascii">remove ASCII codes</label>
		<br>
		<input type="Checkbox" onchange="loadFile()" checked value="1" name="moreSpace" id="moreSpace">
		<label for="moreSpace">extended line space</label>
		<br><br>
		<button onclick="ajaxed.callback('clear', function(r){alert('successfully cleared.');$('logs').update('');}, null, null, 'logs.asp')" type="button">
			clear all logs
		</button>
	</div>
	<div style="float:left;width:75%">
		<div class="console" id="console" style="height:300px;">
			<div id="output"></div>
		</div>
	</div>
	<div style="clear:both">&nbsp;</div>
	</form>
	
<% end sub %>