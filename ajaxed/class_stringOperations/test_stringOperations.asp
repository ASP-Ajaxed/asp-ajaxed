<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

sub test_1()
	tf.assertEqual str.parse("202", 0), 202, "str.parse parses int"
	tf.assertEqual str.parse("229029899809809", 0), 0, "str.parse should not be able to parse a too big int"
	tf.assertEqual year(str.parse("1/1/2007", now())), 2007, "str.parse parsing a date"
	tf.assertEqual str.parse("True", false), true, "str.parse should parse booleans"
	tf.assertEqual str.parse(str.format("2022{0}22", local.comma), 0.0), 2022.22, "str.parse parsing floats"
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
	tf.assertEqual "ein 1 zwei 2, 1", str.format("ein {0} zwei {1}, {0}", array(1, 2)), "str.format does not work"
	tf.assertEqual "{1},{0}", str.format("{0},{1}", array("{1}", "{0}")), "str.format should also allow placeholder as replacement values"
end sub

sub test_5()
	tf.assertEqual "Created on", str.humanize("created_ON"), "str.humanize does not work"
	tf.assertEqual "User", str.humanize("user_id"), "str.humanize does not work"
	tf.assertEqual "User", str.humanize("id_user"), "str.humanize does not work"
	tf.assertEqual "Main category", str.humanize("fk_main_category"), "str.humanize does not work"
	tf.assertEqual "Main category", str.humanize("main_category_fk"), "str.humanize does not work"
	tf.assertEqual "First name", str.humanize("  FIRST_NAME  "), "str.humanize does not work"
	tf.assertEqual "News", str.humanize("id___news_fk"), "str.humanize does not work"
	tf.assertEqual "", str.humanize(empty), "str.humanize does not work"
	tf.assertEqual "", str.humanize(""), "str.humanize does not work"
	tf.assertEqual "A", str.humanize("a"), "str.humanize does not work"
end sub

sub test_6()
	tf.assert str.isAlphabetic("a"), "isAlphabetic() does not work"
	tf.assert str.isAlphabetic("A"), "isAlphabetic() does not work"
	tf.assert str.isAlphabetic("x"), "isAlphabetic() does not work"
	tf.assertNot str.isAlphabetic("_"), "isAlphabetic() does not work"
	tf.assertNot str.isAlphabetic(empty), "isAlphabetic() does not work"
	tf.assertNot str.isAlphabetic(""), "isAlphabetic() does not work"
end sub
'TODO: add more tests here for all the string functions
%>
