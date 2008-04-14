<!--#include file="../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-08 21:38
'* Description: documentation
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
'* main 
'******************************************************************************************
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<a href="http://www.webdevbros.net/ajaxed/"><strong>download current version</strong></a>
	the current stable version can be downloaded on webdevbros.

	<br><br>
	<a href="http://code.google.com/p/asp-ajaxed/"><strong>ajaxed SVN</strong></a>
	the source of ajaxed is controlled with SVN. Therefore it's possible that you
	always get the latest version (edge version) directly from the repository. Good
	if you want to grab the latest changes even before a release has been made.
	
	<br><br>
	<strong>How to get the latest version from SVN?</strong>
	<ol>
		<li>Download and install <a href="http://tortoisesvn.tigris.org/">Tortoise SVN</a></li>
		<li>Right-mouse click in windows explorer on the folder you want to import ajaxed sources.</li>
		<li>Click "SVN Checkout..."</li>
		<li>Enter <code>http://asp-ajaxed.googlecode.com/svn/trunk/</code> as URL of repository</li>
	</ol>

	<strong>Interested in development and contribution?</strong>
	Get in touch with <a href="http://www.webdevbros.net/about-michal/">Michal</a>. He is always happy about new developers.

<% end sub %>
