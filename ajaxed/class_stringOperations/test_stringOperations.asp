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

sub test_2()
	emails = array( _
		"checker@check.com", _
		"che_cker@check.com", _
		"jack.johnson@where-am-i.ac.at", _
		"hottie10@hotmail.com" _
	)
	for each e in emails
		tf.assert str.isValidEmail(e), e & " should be a valid email"
	next
	emails = array( _
		"checkercheck.com", _
		"che_cker@checkcom", _
		"jack.johnson@wh ere-am-i.ac.at", _
		"hotti e10@hotmail.com", _
		"@dump.com", _
		"a@a.a" _
	)
	for each e in emails
		tf.assertNot str.isValidEmail(e), e & " should be an invalid email"
	next
end sub

sub test_3()
	tf.assertEqual "michal", str.rReplace("michXXalX", "x", "", true), "str.rReplace does not work"
	tf.assertEqual "michXXalX", str.rReplace("michXXalX", "x", "", false), "str.rReplace does not work"
	tf.assertEqual "mich.XX.alX", str.rReplace("michXXalX", "(xx)", ".$1.", true), "str.rReplace does not work"
	tf.assertEqual "michXXalX", str.rReplace("michXXalX", "nothing", "", true), "str.rReplace does not work"
end sub

sub test_4()
	tf.assertEqual "ein 1 zwei {1}", str.format("ein {0} zwei {1}", 1), "str.format does not work"
	tf.assertEqual "ein 1 zwei 2", str.format("ein {0} zwei {1}", array(1, 2)), "str.format does not work"
	tf.assertEqual "ein {0} zwei {1}", str.format("ein {0} zwei {1}", array()), "str.format does not work"
end sub
'TODO: add more tests here for all the string functions
%>
