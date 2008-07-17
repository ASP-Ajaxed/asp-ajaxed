<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assertHas array(",", "."), local.comma, "local.comma does not work"
	old = setLocale("en-gb")
	tf.assertEqual ".", local.comma, "local.comma does not work"
	setLocale("de")
	tf.assertEqual ",", local.comma, "local.comma does not work"
	setLocale(old)
end sub
%>