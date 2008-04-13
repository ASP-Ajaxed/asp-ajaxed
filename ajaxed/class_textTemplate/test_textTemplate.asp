<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="textTemplate.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	
	set t = new TextTemplate
	with t
		.fileName = "testTemplate.template"
		.addVariable "version", "0.2"
		.addVariable "MODIFIED", "06.09.2006"
		.addVariable "NAME", "Jack Johnson"
		set block = new TextTemplateBlock
		block.addItem(array("WEEKDAY", "Monday", "VALUE", vbMonday))
		block.addItem(array("WEEKDAY", "Tuesday", "VALUE", vbTuesday))
		block.addItem(array("WEEKDAY", "Friday", "VALUE", vbFriday))
		.addVariable "WEEKDAYS", block
	end with
	
	tf.assertEqual t.getFirstLine(), "Email Subject Jack Johnson", "parsing first line"
	tf.assertMatch "<td>06.09.2006</td>", t.getAllButFirstLine(), "template.getAllButFirstLine()"
	tf.assertMatch "<strong>Tuesday</strong>", t.getAllButFirstLine(), "template.getAllButFirstLine()"
end sub
%>