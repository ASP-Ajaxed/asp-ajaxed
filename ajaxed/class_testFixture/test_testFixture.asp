<!--#include file="testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assertEqual "x", "x", "assertEqual"
	tf.assertEqual 2, 2, "assertEqual"
	tf.assertEqual 2, 2.0, "assertEqual"
	tf.assertEqual empty, empty, "assertEqual"
	tf.assertEqual array(1), array(1), "assertEqual"
	tf.assertEqual array(), array(), "assertEqual"
	tf.assertEqual array(2, 3, "four"), array(2, 3, "four"), "assertEqual"
	tf.assertEqual dateserial(year(date()), month(date()), day(date())), date(), "assertEqual"
end sub

sub test_2()
	tf.assertNotEqual "x", "y", "assertNotEqual"
	tf.assertNotEqual 2, 1, "assertNotEqual"
	tf.assertNotEqual 2, "2", "assertNotEqual"
	tf.assertNotEqual dateadd("d", 1, date()), date(), "assertNotEqual"
	tf.assertNotEqual array(), array(1), "assertNotEqual"
	tf.assertNotEqual array("a", "c"), array("a", "b"), "assertNotEqual"
end sub

sub test_3()
	tf.assertInstanceOf "string", "as", "assertInstanceOf"
	tf.assertInstanceOf "testFixture", new TestFixture, "assertInstanceOf"
end sub

sub test_4()
	tf.assertNothing nothing, "assertNothing"
end sub

sub test_5()
	tf.assertInDelta 1, 2, 1, "assertInDelta"
	tf.assertInDelta 0.5, 0.3, 0.2, "assertInDelta"
	tf.assertInDelta 200, 250, 50, "assertInDelta"
end sub

sub test_6()
	tf.assertMatch "rub. on ....s", "ruby on rails", "assertMatch"
end sub

sub test_7()
	tf.assert 1 = 1, "assert"
end sub

sub test_8()
	tf.assertHas array(1, 2, 3), 1, "assertHas"
	tf.assertHas array("some", "ads", "x"), "ads", "assertHas"
	tf.assertHas array(empty, null, 1), 1, "assertHas"
	tf.assertHasNot "x", "1", "assertHasNot"
	tf.assertHasNot array(1, 2, 3, 4), 0, "assertHasNot"
end sub

sub test_9()
	tf.assertInFile "test_file.txt", "test_testFixture.asp", "'test_testFixture.asp' should be in the test_file.txt"
	tf.assertNotInFile "test_file.txt.no", "test_testFixture.asp", "'test_file.txt.no' does not exists and so it should not contain any matches"
	tf.assertNotInFile "test_file.txt", "jack johnson", "test_file.txt should not contain 'jack johnson'"
end sub

sub test_10()
	testfile = "/ajaxed/class_testFixture/testfile.asp"
	tf.assertResponse testfile, empty, "some test file", "assertResponse() not working"
	tf.assertResponse testfile, array("id", 10), "id=10", "assertResponse() with GET params not working"
	tf.assertResponse array("get", testfile), array("id", 10), "id=10", "assertResponse() with GET params not working"
	tf.assertResponse array("post", testfile), array("x", 12), "x: 12", "assertResponse() with POST params not working"
end sub
%>
