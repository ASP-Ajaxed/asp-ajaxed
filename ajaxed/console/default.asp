<!--#include file="../ajaxed.asp"-->
<%
set page = new AjaxedPage
with page
	.onlyDev = true
	.defaultStructure = true
	.title = "ajaxedConsole"
	.draw()
end with

'******************************************************************************************
'* header  
'******************************************************************************************
sub header()
	page.loadCSSFile "std.css", empty
	page.loadJSFile "../script.aculo.us/scriptaculous.js"
end sub

'******************************************************************************************
'* main  
'******************************************************************************************
sub main() %>
	
	<h1 id="headline">ajaxed Console</h1>
	
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
		<li onclick="loadContent('regex.asp', this)" id="tRegex">
			<a href="#"><span>Regex</span></a>
		</li>
	</ul>
	
	<div id="content" style="padding:20px;margin:40px;background:#FFFFCC" title="">
		
	</div>
	
	<div class="small">
		<a href="http://www.webdevbros.net">ajaxed</a> project founded by Michal Gabrukiewicz
	</div>
	
	<script>
		function loadContent(file, sender) {
			new Ajax.Updater('content', file, {evalScripts:true});
			last = $($('content').readAttribute('title'));
			if (last) last.removeClassName('active');
			$(sender).addClassName('active');
			$('content').writeAttribute('title', sender.id);
		}
		loadContent('configuration.asp', $('tConfig'));
		new Effect.BlindDown('headline');
	</script>
	
<% end sub %>