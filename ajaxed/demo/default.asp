<!--#include file="../ajaxed.asp"-->
<%
set page = new AjaxedPage
with page
	.title = "ajaxed demo - classic asp library"
	.defaultStructure = true
	.draw()
end with

'******************************************************************************************
'* init 
'******************************************************************************************
sub init()
	delay = str.parse(page.RF("delay"), 0)
	if delay > 0 and delay < 5 then lib.sleep delay
end sub

'******************************************************************************************
'* callback - this is the function which will be processed when any ajaxed.callback
'* javascript function is executed. The action is here to differentiate between
'* all your ajaxed.callback functions in the page.
'******************************************************************************************
sub callback(action)
	if action = "add" then
		'str.parse is VERY handy function which converts a variable into
		'the desired type. if it cannot be converted then the alternative value is returned
		page.return(str.parse(page.RF("a"), 0) + str.parse(page.RF("b"), 0))
		
		'see also page.returnValue - with this you can return even more values in one go
		'e.g. a recordset with data and a boolean flag which indicates something else
	end if
end sub

'******************************************************************************************
'* header  
'******************************************************************************************
sub header()
	page.loadCSSFile "../console/std.css", empty
end sub

'******************************************************************************************
'* sub  
'******************************************************************************************
sub pagePart_one() %>
	
	here is some content of the page part one.<br>
	this HTML is simply located within a sub called "pagePart_one".
	<strong>Thats all you need ;)</strong>	
	
<% end sub %>

<% sub pagePart_two() %>

	and here is another part. Check the source code to see how it works!<br>
	This could be some tabs in an application ... maybe. ajaxed rocks on asp :)
	
<% end sub %>

<% sub main() %>
	
	<style>
		body {
			padding:100px;
			padding-top:20px;
		}
	</style>

	<script>
		function added(sum) {
			$('c').value = sum;
		}
	</script>
	
	<h1>ajaxed v<%= lib.version %> Demo</h1>
	
	<form id="frm">
	<div style="padding:20px;background:#eee;">
		<strong>1. Calling a server side procedure</strong>
		<p>
			Test ajaxed functionality by adding up two number on server side using AJAX.
			Enter two numbers and press calculate to get the total. You will see that
			the page is not reloaded. Also recognize the "loading" indicator at the top right of the page. 
		</p>
		
		<input type="Text" name="a" id="a" size="4">+
		<input type="Text" name="b" id="b" size="4">=
		<input type="Text" name="c" id="c" size="4">
		<button onclick="ajaxed.callback('add', added, null, null, 'default.asp')" type="button">calculate</button>
	
		<br><br>
		<strong>2. Using page parts</strong>
		<p>
			Ajaxed also supports loading whole HTML fragments. This functionality is called "page parts". Page parts 
			are simply sub's in your code which return HTML. This markup is then inserted automatically into a container
			of your choice. See a simple example:
		</p>
		
		<a href="javascript:void(0)" onclick="ajaxed.callback('one', 'content', null, null, 'default.asp')">Part one</a>
		&nbsp;&nbsp;
		<a href="javascript:void(0)" onclick="ajaxed.callback('two', 'content', null, null, 'default.asp')">Part two</a>
		<div id="content" style="margin-top:10px;"></div>
	</div>
	
		<br><br>
		<strong>You dont see the loading indicator?</strong> Set some seconds delay for the calculation:
		<input type="Text" name="delay" size="4" value="0"> seconds (1-5)
	</form>
		
	<br>
	Take also a look at the ajaxed developer console
	<a href="/ajaxed/console"><strong>ajaxed console</strong></a> which supports you while you develop.
	
	<br><br>
	<a href="http://www.ajaxed.org/">asp ajaxed</a>
	
	
	
<% end sub %>