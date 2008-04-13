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
	
	'TODO: michal: lib range is not working with floats perfectly. I am leaving it for now because with int it works fine.
	'i dunno now why it has a little offset after sometime.. uncomment next line to see the problem:
	'str.write(str.arrayToString(lib.range(-10, -5, 0.1), " - "))
	tf.assertHas lib.range(-10, -5, 0.1), -5.2, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 9.6, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 20, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 17.6, "lib.range"
end sub
%>
