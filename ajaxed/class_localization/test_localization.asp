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
	'this IP is located in austria
	anIP = "213.129.245.250"
	
	code = local.locateClient(0, anIP)
	if isempty(code) then
		tf.info "Could not test Localization.geocodeClient() because it seems there is no network connection."
		exit sub
	end if
	if code = "XX" then
		tf.info "Localization.geocodeClient() does not contain information about your IP. " & anIP
	else
		tf.assert code <> "", "Country not identified"
	end if
	
	'private IP
	tf.assertEqual "XX", local.locateClient(0, "127.0.0.1"), "Private addresses should be marked as unknown"
end sub
%>