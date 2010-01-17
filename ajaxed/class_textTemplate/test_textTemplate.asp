<% on error resume next %>
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
		.addVariable "notparsed", "this should not be added, because the placeholder is defined wrong in the template"
		set block = new TextTemplateBlock
		block.addItem(array("WEEKDAY", "Monday", "VALUE", vbMonday))
		block.addItem(array("WEEKDAY", "Tuesday", "VALUE", vbTuesday))
		block.addItem(array("WEEKDAY", "Friday", "VALUE", vbFriday))
		block.addItem(array("WEEKDAY", empty, "VALUE", vbsaturday))
		.addVariable "WEEKDAYS", block
	end with
	
	'test first line
	tf.assertEqual t.getFirstLine(), "Email Subject Jack Johnson", "parsing first line"
	
	'test common placeholders
	parsedTemplate = t.getAllButFirstLine()
	tf.assertMatch "<td>06.09.2006</td>", parsedTemplate, "template.getAllButFirstLine() not working"
	tf.assertMatch "<td>Jack Johnson</td>", parsedTemplate, "NAME placeholder not parsed correctly"
	tf.assertMatch "Modified by: Jack Johnson", parsedTemplate, "NAME placeholder not parsed correctly (multiple times)"
	tf.assertMatch "<td>0.2</td>", parsedTemplate, "VERSION placeholder not parsed correctly"
	tf.assertMatch "<td><<<NOTPARSED>>></td>", parsedTemplate, "NOTPARSED placeholder was parsed (although invalid)"
	
	'test default placeholder values
	tf.assertMatch "<td>John Doe</td>", parsedTemplate, "default walue not set. we did not added an APPROVER variable therefore default value should be passed through"
	tf.assert instr(parsedTemplate, "not used value") = 0, "default value should be cleaned out if a variable is available"
	tf.assertMatch "<h1>none</h1>", parsedTemplate, "WEEKDAYS saturday does not contain a value and therefore default should be used"
	
	'test first block
	tf.assertMatch "<td>Monday: " & vbMonday & "</td>", parsedTemplate, "first WEEKDAYS (monday) block not parsed correctly"
	tf.assertMatch "<td>Tuesday: " & vbTuesday & "</td>", parsedTemplate, "first WEEKDAYS (tuesday) not parsed correctly"
	tf.assertMatch "<td>Friday: " & vbFriday & "</td>", parsedTemplate, "first WEEKDAYS (friday) block not parsed correctly"
	
	'test second block (has same name)
	tf.assertMatch "<h1>Monday</h1>", parsedTemplate, "WEEKDAYS (monday) block not parsed correctly (placeholder multiple times should work)"
	tf.assertMatch "<strong>Monday</strong> \(" & vbMonday & "\)", parsedTemplate, "WEEKDAYS (monday) block not parsed correctly"
	
	tf.assertMatch "<h1>Tuesday</h1>", parsedTemplate, "WEEKDAYS (tuesday) block not parsed correctly (placeholder multiple times should work)"
	tf.assertMatch "<strong>Tuesday</strong> \(" & vbTuesday & "\)", parsedTemplate, "WEEKDAYS (tuesday) block not parsed correctly"
	
	tf.assertMatch "<h1>Friday</h1>", parsedTemplate, "WEEKDAYS (friday) block not parsed correctly (placeholder multiple times should work)"
	tf.assertMatch "<strong>Friday</strong> \(" & vbFriday & "\)", parsedTemplate, "WEEKDAYS (friday) block not parsed correctly"
	
	'unused blocks
	tf.assert instr(parsedTemplate, "CLEANEDOUT") = 0, "unused blocks should be cleaned out"
end sub

'test adding a recordset
sub test_2()
	set t = new TextTemplate
	with t
		.fileName = "testTemplate.template"
		set block = new TextTemplateBlock
		block.addRS(["R"](array( _
			array("weekday", "value"), _
			array("Monday", "first"), _
			array("Sunday", "last") _
		)))
		.addVariable "WEEKDAYS", block
	end with
	
	parsedTemplate = t.getAllButFirstLine()
	tf.assertMatch "<td>Monday: first</td>", parsedTemplate, "block not parsed with recordset"
	tf.assertMatch "<td>Sunday: last</td>", parsedTemplate, "block not parsed with recordset 2"
end sub

'test invalid filename and empty content
sub test_3()
    set t = new TextTemplate
    t.fileName = "invalid.template"

    'invalid filename and empty content
    on error resume next
    t.returnString()

    if Err.number = 0 then
        tf.fail("TextTemplate should throw an exception when filename is invalid and content is null")
    end if
    on error goto 0
    
    t.content = "Now, content have some text"
    tf.assertMatch "Now, content have some text", t.returnString(), "Template content not parsed correctly"
end sub

'test default parameter value <<< TAG | default value >>>
sub test_4()
    set t = new TextTemplate
    t.content = "Test <<< TEST1 | defaultvalue >>>"
    
    tf.assertMatch "Test defaultvalue", t.returnString(), "Template content not parsed correctly" 
    t.add "TEST1", ""
    tf.assertMatch "Test ", t.returnString(), "Template content not parsed correctly" 
    
    t.add "TEST1", "value"
    tf.assertMatch "Test value", t.returnString(), "Template content not parsed correctly" 
    
    t.content = "<<< TAG1 | defaultvalue1 >>> and <<< TAG2 | defaultvalue2 >>>"    
    
    tf.assertMatch "defaultvalue1 and defaultvalue2", t.returnString(), "Template content not parsed correctly"
    
    t.add "TAG1", "value1"
    tf.assertMatch "value1 and defaultvalue2", t.returnString(), "Template content not parsed correctly"
    t.add "TAG2", "value2"
    tf.assertMatch "value1 and value2", t.returnString(), "Template content not parsed correctly"
end sub

' Regression test to issue #15
' Replacing a block, we get an exception! It's wrong
sub test_5()
	set t = new TextTemplate
	set b = new TextTemplateBlock
	b.additem Array("item", "first")
	t.add "MYBLOCK", b
	b.additem Array("item", "another")
	t.add "MYBLOCK", b
    t.add "MYBLOCK", "value"
end sub
%>

