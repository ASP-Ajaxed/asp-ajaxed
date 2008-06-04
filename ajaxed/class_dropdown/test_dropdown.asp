<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="../class_database/testDatabases.asp"-->
<!--#include file="dropdown.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

'test with recordset - quick instantiation
sub test_1()
	db.open(TEST_DB_SQLITE)
	set dd = (new Dropdown)("SELECT * FROM person", "person", 0)
	performTests(dd)
end sub

'dictionary
sub test_2()
	set dd = (new Dropdown)(lib.newDict(array(array(1, "jack"), array(2, "britney"))), _
		"person", 0)
	performTests(dd)
end sub

'array
sub test_2()
	set dd = (new Dropdown)(array("jack", "britney"), _
		"person", 0)
	dd.valuesDatasource = array(1, 2)
	performTests(dd)
end sub

sub performTests(drpDown)
	'set some properties to test the properties afterwards
	with drpDown
		.cssClass = "cssClass"
		.id = "someID"
		.attributes = "onclick=""alert(2)"""
		.style = "color:red"
		.onItemCreated = "onItemCreated"
		dd = .toString()
	end with
	msg = "dropdown not rendering properly"
	tf.assertMatch "<select", dd, msg
	tf.assertMatch "value=""1""", dd, msg
	tf.assertMatch "value=""2""", dd, msg
	tf.assertMatch ">Jack&lt;strong&gt;<", dd, msg
	tf.assertMatch ">Britney&lt;strong&gt;<", dd, msg
	tf.assertMatch "name=""person""", dd, msg
	tf.assertMatch "class=""cssClass""", dd, msg
	tf.assertMatch "id=""someID""", dd, msg
	tf.assertMatch "onclick=""alert\(2\)""", dd, msg
	tf.assertMatch "style=""color:red""", dd, msg
	tf.assertMatch "</select>$", dd, msg
end sub

sub onItemCreated(it)
	'test if on item created works fine and if item text is being html encoded
	it.text = it.text & "<strong>"
end sub
%>