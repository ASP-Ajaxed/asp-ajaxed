<!--#include file="../ajaxed.asp"-->
<%
regexError = ""
set page = new AjaxedPage
with page
	.onlyDev = true
	.defaultStructure = true
	.plain = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* callback 
'******************************************************************************************
sub callback(a)
	if not a = "exec" then exit sub
	
	session("ajaxedConsoleRegexPattern") = page.RF("pattern")
	session("ajaxedConsoleRegexSourceString") = page.RF("searchstring")
	if page.RF("type") = "execute" then
		set result = getResult(true)
		if not result is nothing then
			i = 1
			for each m in result
				ret = ret & "Match " & i & " '" & m.value & "' starts at index " & m.firstIndex & " with a length " & m.length & "<br>"
				i = i + 1
			next
			page.returnValue "result", ret
			if result.count = 0 then
				page.returnValue "result", "No matches found."
			end if
		end if
	else
		result = getResult(false)
		page.returnValue "result", result
	end if
	page.returnValue "err", regexError
	if page.RFHas("showCode") then
		code = 	"set rgx = new Regexp<br>" & _
				"rgx.pattern = """ & page.RFE("pattern") & """<br>" & _
				lib.iif(page.RFHas("ignorecase"), "rgx.ignoreCase = true<br>", "") & _
				lib.iif(page.RFHas("global"), "rgx.global = true<br>", "")
		if page.RF("type") = "test" then
			code = code & "result = rgx.test(""" & page.RFE("searchstring") & """)<br>"
		elseif page.RF("type") = "execute" then
			code = code & "set matches = rgx.execute(""" & page.RFE("searchstring") & """)<br>"
		else
			code = code & "result = rgx.replace(""" & page.RFE("searchstring") & """, """ & page.RF("replaceby") & """)<br>"
		end if
		page.returnValue "code", "Code used:<br>" & code & "set rgx = nothing"
	end if
end sub

'******************************************************************************************
'* rExec 
'******************************************************************************************
function getResult(expectingObject)
	set r = new Regexp
	r.pattern = page.RF("pattern")
	r.ignoreCase = page.RFHas("ignorecase")
	r.global = page.RFHas("global")
	if expectingObject then set getResult = nothing
	on error resume next
	select case page.RF("type")
		case "test"
			if r.test(page.RF("searchstring")) then
				getResult = "Pattern found in search string."
			else
				getResult = "Pattern NOT found in search string."
			end if
		case "execute"
			set getResult = r.execute(page.RF("searchstring"))
		case "replace"
			getResult = r.replace(page.RF("searchstring"), page.RF("replaceby"))
	end select
	if err <> 0 then
		regexError = err.description
		if expectingObject then
			set getResult = nothing
		else
			getResult = ""
		end if
	end if
	on error goto 0
end function

'******************************************************************************************
'* content 
'******************************************************************************************
sub main() %>

	This helps you quickly build and test a regular expression using the <code>VBScript Rexgexp</code> object.
	Play around and tune your regular expression for the use within your ASP ajaxed web applications.
	<br><br>
	
	<form id="frm">
	
	<div style="float:left;width:50%">
		<strong>Pattern:</strong><br>
		<div><textarea id="tPattern" rows="2" style="width:90%;" class="code" name="pattern"><%= session("ajaxedConsoleRegexPattern") %></textarea></div>
		<small><a href="javascript:void(0)" onclick="$('tPattern').rows += 3">larger</a></small>
		<br>
		<a href="http://msdn2.microsoft.com/en-us/library/ms974570.aspx" target="_blank">Regex Reference</a>
		<br><br>
		<input id="id012" type="Checkbox" name="ignorecase" value="1" checked>
		<label for="id012" title="Tick this option to perform a case insensitive test.">Ignore case</label>
		<br>
		<input id="id010" type="Checkbox" name="global" value="1" checked>
		<label for="id010" title="Tick this if the search should not be stopped after the first match.">Global (e.g. replace will be executed on each match)</label>
		<br>
		<input id="id019" type="Checkbox" name="showCode" value="1">
		<label for="id019" title="Tick this option to display the used VBScript in order to make the regular expression execution.">Show used VBScript code</label>
		<br><br>
		<input id="id001" type="radio" name="type" value="test" checked>
		<label for="id001">Test (true if pattern matches the search string)</label>
		<br>
		<input id="id002" type="radio" name="type" value="execute">
		<label for="id002">Execute (returns matches for the pattern)</label>
		<br>
		<input id="id003" onclick="$('rBy').activate()" type="radio" name="type" value="replace">
		<label for="id003">Replace each match with:</label>
		<input type="Text" id="rBy" onfocus="$('id003').checked = true" size="20" name="replaceby">
		
		<br><br>
		<script>
			executed = function(r) {
				if (r.code) writeConsole(r.code);
				if (r.result) writeConsole(r.result);
				if (r.err) writeConsole(r.err);
				scrollToFloor('regexConsole');
			}
			writeConsole = function(msg) {
				if (msg == '') return;
				var c = $('regexConsole');
				c.insert({bottom: '<br>&gt; ' + msg});
			}
		</script>
		<button type="button" onclick="ajaxed.callback('exec', executed, null, null, 'regex.asp')">
			Execute
		</button>	
	</div>
	<div style="float:left;width:50%">
		<strong>Search string:</strong><br>
		<div><textarea rows="2" class="code" style="width:100%" id="tSearchstring" name="searchstring"><%= session("ajaxedConsoleRegexSourceString") %></textarea></div>
		<small><a href="javascript:void(0)" onclick="$('tSearchstring').rows += 3">larger</a></small>
		<br><br>
		<div class="console" id="regexConsole">&gt; Ready.</div>
		<a href="javascript:void(0)" onclick="$('regexConsole').update('&gt; Ready.')">clear console</a>
	</div>
	<div style="clear:both">&nbsp;</div>
	
	</form>

<% end sub %>