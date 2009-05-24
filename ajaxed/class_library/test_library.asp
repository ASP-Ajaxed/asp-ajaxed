<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.debug = true
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
	tf.assertEqual "scripting.dictionary", lib.detectComponent(array("bla", "scripting.dictionary")), "lib.detectComponent does not work"
end sub

sub test_4()
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

sub test_5()
	tf.assertEqual "/ajaxed/", lib.path(empty), "lib.path seem not to work"
	tf.assertEqual "/ajaxed/class_rss/rss.asp", lib.path("class_rss/rss.asp"), "lib.path seem not to work"
end sub

sub test_6()
	set o = ["O"](array("a", "b"), array("B", 1), empty)
	tf.assertEqual empty, o("a"), "Library.options() does not work"
	tf.assertEqual 1, o("b"), "Library.options() does not work"
	
	set o = ["O"](array("a", "b"), array("B", 1), empty)
	
	on error resume next
		set o = ["O"](array("a", ""), array("B", 1), empty)
	if err = 0 then tf.fail("Library.options() must raise an error")
	on error goto 0
	
	on error resume next
		set o = ["O"](empty, array("B", 1), empty)
	if err = 0 then tf.fail("Library.options() must raise an error")
	on error goto 0
	
	on error resume next
		set o = ["O"](array(), array("B", 1), empty)
	if err = 0 then tf.fail("Library.options() must raise an error")
	on error goto 0
	
	on error resume next
		set o = ["O"](array("B"), array("B"), empty)
	if err = 0 then tf.fail("Library.options() must raise an error")
	on error goto 0
	
	options = array("b", 2)
	["O"] array("a", "b"), options, empty
	tf.assert options.exists("a"), "Library.options() does not change the original value"
	
	options = empty
	["O"] array("a", "b"), options, 0
	tf.assertEqual 0, options("a"), "Library.options() does not change the original value"
	tf.assertEqual 0, options("b"), "Library.options() does not change the original value"
	
	set o = ["O"](array("a", "b", "c"), array(), array(1, 0))
	tf.assertEqual 1, o("a"), "Library.options() default values"
	tf.assertEqual 0, o("b"), "Library.options() default values"
	tf.assertEqual 0, o("c"), "Library.options() default values"
	
	set o = ["O"](array("a", "b", "c"), array(), empty)
	tf.assertEqual empty, o("a"), "Library.options() default values"
	tf.assertEqual empty, o("b"), "Library.options() default values"
	tf.assertEqual empty, o("c"), "Library.options() default values"
	
	set o = ["O"](array("a", "b", "c"), array(), array(2, empty))
	tf.assertEqual 2, o("a"), "Library.options() default values"
	tf.assertEqual empty, o("b"), "Library.options() default values"
	tf.assertEqual empty, o("c"), "Library.options() default values"
end sub

sub test_7()
	tf.assertEqual "x", []("x")(0), "Alias [] not working"
	tf.assertEqual "y", [](array("x", "y"))(1), "Alias [] not working"
	tf.assertEqual -1, ubound([](empty)), "Alias [] not working"
end sub

'test ["R"]()
sub test_8()
	set rs = ["R"](array( _
		array("firstname", "lastname", "createdOn", "age"), _
		array("John", "Doe", dateserial(2000, 1, 1), 80, "column ignored cause the definition does not have it"), _
		array("Tomy", "Foo", dateserial(1999, 1, 1)) _
	))
	tf.assert not rs.eof, "Must be at the beginning"
	tf.assertEqual 2, rs.recordcount, "wrong recordcount"
	
	tf.assertEqual "John", RS("firstname").value, "firstname column wrong value (1st record)"
	tf.assertEqual "Doe", RS("lastname").value, "lastname column wrong value (1st record)"
	tf.assertEqual dateserial(2000, 1, 1), RS("createdOn").value, "createdOn column wrong value (1st record)"
	tf.assertEqual 80, RS("age").value, "age column wrong value (1st record)"
	
	rs.movenext()
	tf.assertEqual "Tomy", RS("firstname").value, "firstname column wrong value (2nd record)"
	tf.assertEqual "Foo", RS("lastname").value, "lastname column wrong value (2nd record)"
	tf.assertEqual dateserial(1999, 1, 1), RS("createdOn").value, "createdOn column wrong value (2nd record)"
	tf.assertEqual null, RS("age").value, "age column wrong value (2nd record) - should have no value as age is not given"
	
	rs.movefirst()
	while not RS.eof
		tf.assert RS("firstname") <> "", "firstname column missing"
		tf.assert RS("lastname") <> "", "lastname column missing"
		tf.assert RS("createdOn") <> "", "createdOn column missing"
		tf.assert isDate(RS("createdOn")), "createdon column not a date"
		RS.movenext()
	wend
	
	'new one without records
	set rs = ["R"](array( _
		array("firstname", "lastname", "createdOn", "age") _
	))
	tf.assert RS.eof, "no records should work"
	rs.addNew array(0, 1, 2), array("John", "Doe", dateserial(2000, 1, 1))
	rs.movefirst()
	tf.assertEqual "John", RS("firstname").value, "firstname column wrong value (custom record)"
	tf.assertEqual "Doe", RS("lastname").value, "lastname column wrong value (custom record)"
	tf.assertEqual dateserial(2000, 1, 1), RS("createdOn").value, "createdOn column wrong value (custom record)"
end sub
%>
