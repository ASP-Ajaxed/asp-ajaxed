<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="de" lang="de">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title><%= title %></title>
		<% 'if there is a header sub, then it will be called. e.g. for custom styles, etc.  %>
		<% ajaxedHeader(array()) %>
		<% lib.exec "header", empty %>
	</head>
	<body>