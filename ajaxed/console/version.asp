<!--#include file="../ajaxed.asp"-->
<!--#include file="../class_RSS/RSS.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-14 20:48
'* Description: generates an RSS feed which populates the version of ajaxed
'******************************************************************************************

set page = new AjaxedPage
page.plain = true
page.onlyDev = true
page.draw()

sub main()
	set r = new RSS
	r.title = "ajaxed version feed"
	r.description = lib.version
	r.link = "http://www.webdevbros.net/ajaxed"
	r.publishedDate = now()
	r.language = "en"
	r.generate("RSS2.0", empty).save(response)
	set r = nothing
end sub
%>