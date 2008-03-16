<!--#include file="../ajaxed.asp"-->
<%
set page = new AjaxedPage
page.draw()

sub init() : end sub

sub callback(action)
	if action = "add" then
		page.return(add(page.RF("a"), page.RF("b")))
	end if
end sub

function add(a, b)
	if not isnumeric(a) then a = 0
	if not isnumeric(b) then b = 0
	add = cint(a) + cint(b)
end function

sub main() %>

	<% for each f in request.form %>
		<%= f %>: <%= left(request.form(f), 10) %><br>
	<% next %>
	
	<script>
		function added(sum) {
			$('c').value = sum;
		}
	</script>
	
	<style>
		.ajaxLoadingIndicator {
			background-color: #cc0000;
			font-family:arial;
			color: #fff;
			margin:2px;
			right: 0px;
		}
	</style>

	<br><br><br><br><br><br><br>
	<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
	
	<form id="frm">
		<input type="Text" name="a" id="a">+
		<input type="Text" name="b" id="b">=
		<input type="Text" name="c" id="c">
		<button onclick="ajaxed.callback('add', added, null, null, 'default.asp')" type="button">calculate</button>
	</form>
	
<% end sub %>