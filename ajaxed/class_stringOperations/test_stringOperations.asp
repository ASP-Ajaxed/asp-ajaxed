<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assertEqual str.parse("202", 0), 202, "str.parse parses int"
	tf.assertEqual str.parse("229029899809809", 0), 0, "str.parse should not be able to parse a too big int"
	tf.assertEqual str.parse("2022,22", 0.0), 2022.22, "str.parse parsing floats"
	tf.assertEqual year(str.parse("1/1/2007", now())), 2007, "str.parse parsing a date"
	tf.assertEqual str.parse("True", false), true, "str.parse should parse booleans"
end sub

'TODO: add more tests here for all the string functions
%>
