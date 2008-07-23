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

sub test_2()
	set d = local.geocodeClient(0)
	if d is nothing then
		tf.info "Could not test Localization.geocodeClient() because it seems there is no network connection."
		exit sub
	end if
	if d.count > 0 then
		tf.assert d("country") <> "", "Country not identified"
	else
		tf.info "Localization.geocodeClient() does not contain information about your IP."
	end if
end sub
%>