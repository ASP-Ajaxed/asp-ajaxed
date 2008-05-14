<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	pages = array( _
		"", "ajaxed console", _
		"configuration.asp", "ajaxed version", _
		"tests.asp", "run all", _
		"documentation.asp", "generate...", _
		"source.asp", "SVN", _
		"templates.asp", "existing templates", _
		"logs.asp", "clear all", _
		"regex.asp", "pattern:" _
	)
	for i = 0 to uBound(pages) step 2
		tf.assertResponse "/ajaxed/console/" & pages(i), empty, pages(i + 1), pages(i) & " seem not to respond!"
	next
end sub
%>