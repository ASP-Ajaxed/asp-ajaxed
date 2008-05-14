<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="de" lang="de">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>some test file for testing</title>
	</head>
	<body>
		the content of the testfile<br>
		<%= request.queryString %><br>
		<% for each f in request.form %>
			<%= f %>: <%= request.form(f) %><br>
		<% next %>
	</body>
</html>