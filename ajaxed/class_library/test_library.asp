<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assertNotEqual lib.getGUID(), lib.getGUID(), "lib.getGUID should return unique IDs"
	tf.assert lib.getUniqueID() < lib.getUniqueID(), "lib.getUniqueID() should return increased numbers"
	tf.assertEqual lib.iif(2 > 1, true, false), true, "lib.iif 1st"
	tf.assertEqual lib.iif(2 < 1, true, false), false, "lib.iif 2nd"
	tf.assertEqual lib.init(empty, true), true, "lib.init should pass the alternative is value is empty"
	tf.assertEqual lib.init(false, true), false, "lib.init should NOT pass the alternative is value is NOT empty"
	tf.assert lib.URLDecode("%20") = " ", "libURLDecode should parse a %20 into a space"
	
	on error resume next
	lib.throwError("check")
	tf.assert err <> 0, "lib.throwError must produce an error"
	if err <> 0 then
		tf.assertEqual err.number, 1024, "Thrown custom error number should be 1024"
		tf.assertEqual err.description, "check", "Thrown custom error description should be 'checked'"
	end if
	on error goto 0
end sub

sub test_2()
	tf.assertEqual lib.range(1, 3, 1), array(1, 2, 3), "lib.range"
	tf.assertEqual lib.range(5, 0, -1), array(5, 4, 3, 2, 1, 0), "lib.range"
	tf.assertEqual lib.range(1, 3, 0.5), array(1, 1.5, 2, 2.5, 3), "lib.range"
	tf.assertEqual lib.range(1, 3.2, 0.5), array(1, 1.5, 2, 2.5, 3), "lib.range"
	tf.assertEqual lib.range(1, 1.2, 1), array(1), "lib.range"
	tf.assertEqual lib.range(1, 2, 1), array(1, 2), "lib.range"
	tf.assertHas lib.range(1, 20, 1), 1, "lib.range"
	tf.assertHas lib.range(20, 1, -1), 10, "lib.range"
	tf.assertHas lib.range(1, 20, 1), 20, "lib.range"
	tf.assertHas lib.range(1, 20, 1), 15, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), 0, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), -10, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), -5, "lib.range"
	tf.assertHas lib.range(-10, 10, 1), 10, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 10, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 50.5, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 100, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 0, "lib.range"
	tf.assertHasNot lib.range(0, 100, 0.5), 100.5, "lib.range"
	tf.assertHas lib.range(0.1, 10.1, 1), 10.1, "lib.range"
	tf.assertHas lib.range(0.1, 10.1, 1), 0.1, "lib.range"
	tf.assertHas lib.range(0.12, 10.12, 5.00), 5.12, "lib.range"
	tf.assertHas lib.range(-10, -5, 0.1), -5.2, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 9.6, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 20, "lib.range"
	tf.assertHas lib.range(0, 20, 0.100), 20, "lib.range"
	tf.assertHas lib.range(0, 0, 1), 0, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 17.6, "lib.range"
	tf.assertHas lib.range(1.0001, 10, 0.1), 9.9001, "lib.range"
	tf.assertHas lib.range(1, 2, 0.01), 1.99, "lib.range"
	tf.assertHasNot lib.range(1, 10.0001, 0.1), 11.1001, "lib.range"
end sub

sub test_3()
	tf.assert lib.contains(array(1, 2, 3), 2), "lib.contains"
	tf.assert lib.contains(array(1, 2, 3), "3"), "lib.contains"
	tf.assert lib.contains(array(1, 2, 3), 2.0), "lib.contains"
	tf.assert lib.contains(array("test", "who", 3), "3"), "lib.contains"
	tf.assertNot lib.contains("3", "3"), "lib.contains"
	tf.assertNot lib.contains(array(1, 2, 3, 10), 11), "lib.contains"
	
	set d = lib.newDict(empty)
	d.add 1, "someting"
	d.add 2, ""
	d.add "3", "yeah"
	d.add "not", empty
	tf.assert lib.contains(d, 2), "lib.contains with dictionary"
	tf.assert lib.contains(d, "2"), "lib.contains with dictionary"
	tf.assertNot lib.contains(d, 0), "lib.contains with dictionary"
	tf.assertNot lib.contains(d, 3.2), "lib.contains with dictionary"
	tf.assert lib.contains(d, "not"), "lib.contains with dictionary"
end sub

sub test_4()
	tf.assertEqual "scripting.dictionary", lib.detectComponent(array("bla", "scripting.dictionary")), "lib.detectComponent does not work"
end sub

sub test_5()
	on error resume next
		lib.require "SomeNonExistingClass", "Test_5"
	if err <> 0 then
		tf.assertEqual err.number, 1362, "lib.require raises a wrong error when requiring a not existing class"
	else
		tf.fail("lib.require did not raise an 1362 error when requiring a non existing class")
	end if
	on error goto 0
	on error resume next
		lib.require "TestFixture", "Test_5"
	if err <> 0 then tf.fail("lib.require raised an error on an existing class")
	on error goto 0
end sub

sub test_6()
	tf.assertEqual "/ajaxed/", lib.path(empty), "lib.path seem not to work"
	tf.assertEqual "/ajaxed/class_rss/rss.asp", lib.path("class_rss/rss.asp"), "lib.path seem not to work"
end sub
%>
