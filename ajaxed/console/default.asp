<!--#include file="../ajaxed.asp"-->
<!--#include file="../class_RSS/RSS.asp"-->
<%
set page = new AjaxedPage
with page
	.onlyDev = true
	.defaultStructure = true
	.title = "console"
	.draw()
end with

'******************************************************************************************
'* header  
'******************************************************************************************
sub header()
	str.write("<link rel=""shortcut icon"" href=""img/fav.png"" type=""image/ico"" />")
	page.loadCSSFile "std.css", empty
	page.loadCSSFile "screen_borderstyles.css", empty
	page.loadJSFile "console.js"
	page.loadJSFile "../script.aculo.us/scriptaculous.js"
end sub

'******************************************************************************************
'* getVersion 
'******************************************************************************************
function getVersion()
	getVersion = ""
	set r = new RSS
	r.url = "http://www.ajaxed.org/ajaxed/console/version.asp"
	r.load()
	if not r.failed then getVersion = r.description
end function

'******************************************************************************************
'* main  
'******************************************************************************************
sub main() %>
	
	<% version = getVersion() %>
	<% if lib.version <> getVersion() and version <> "" then %>
		<div class="rounded blue" style="position:absolute;top:5px;right:10px;height:30px">
			<b class="top"><b><b><b></b></b></b></b>
			<div class="content">
				&nbsp;&nbsp;
				version <strong><%= version %></strong> is available.
				<a href="http://www.ajaxed.org/"><strong>Update now!</strong></a>
				&nbsp;&nbsp;
			</div>
			<b class="bottom"><b><b><b></b></b></b></b>
		</div>
	<% end if %>
	
	<h1 id="headline"><a href="http://www.ajaxed.org" target="_blank"><img border="0" src="../logo.png" alt="ajaxed logo"></a></h1>
	
	<ul class="tab">
		<li onclick="loadContent('configuration.asp', this)" id="tConfig">
			<a href="#"><span>Configuration</span></a>
		</li>
		<li onclick="loadContent('tests.asp', this)" id="tTests">
			<a href="#"><span>Tests</span></a>
		</li>
		<li onclick="loadContent('documentation.asp', this)" id="tDoc">
			<a href="#"><span>Documentation</span></a>
		</li>
		<li onclick="loadContent('source.asp', this)" id="tSource">
			<a href="#"><span>Source</span></a>
		</li>
		<li onclick="loadContent('templates.asp', this)" id="tTemplates">
			<a href="#"><span>Templates</span></a>
		</li>
		<li onclick="loadContent('logs.asp', this)" id="tLogs">
			<a href="#"><span>Logs</span></a>
		</li>
		<li onclick="loadContent('regex.asp', this)" id="tRegex">
			<a href="#"><span>Regex</span></a>
		</li>
	</ul>
	
	<div id="htmlcontent" style="padding:20px;margin:40px;background:#FFFFCC" title=""></div>
	
	<div class="small">
		<a href="http://www.ajaxed.org">ajaxed</a>
		project founded by Michal Gabrukiewicz
	</div>
	
	<script>
		loadContent('configuration.asp', $('tConfig'));
	</script>
	
<% end sub %>