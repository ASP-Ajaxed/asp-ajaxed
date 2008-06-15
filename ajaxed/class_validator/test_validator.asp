<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="validator.asp"-->
<%
set tf = new TestFixture
tf.debug = True
tf.run()

sub test_1()
	set v = new Validator
	tf.assertEqual empty, v.getErrorSummary("<div>", "</div>", "<li>", "</li>"), "validator.getErrorSummary() must return no summary if no errors"
	tf.assert v, "New validator must be valid"
	tf.assert v.add("firstname", "firstname wrong"), "validator.add() must return true on add()"
	tf.assertNot v.valid, "after adding an invalid field the validator must be invalid"
	tf.assertNot v.add("firstname", "somethign"), "validator.add() must return false if a field already exists with this name"
	tf.assert v.add("lastname", "lastname wrong"), "validator.add() must return true on add"
	tf.assert v.isInvalid("FIRSTname"), "validator.isInvalid() must return true if field is invalid"
	tf.assert v.isInvalid(array("some", "lastname")), "validator.isInvalid() must return true if at least one field is invalid"
	tf.assertNot v.isInvalid(array("some", "other")), "validator.isInvalid() must return false if no field is invalid"
	tf.assertEqual "lastname wrong", v.getDescription("lastname"), "Validator.getDescritpion() does not work"
	sum = v.getErrorSummary("<ul>", "</ul>", "<li>", "</li>")
	tf.assertMatch "<li>lastname wrong</li>", sum, "validator.getErrorSummary() not working"
	tf.assertMatch "<li>firstname wrong</li>", sum, "validator.getErrorSummary() not working"
	tf.assertMatch "^<ul>.*</ul>$", sum, "validator.getErrorSummary() must contain overall prefix and postfix"
	tf.assert v.reflect()("data").count = 2, "validator.reflect() must return 2 data items"
	tf.assertNot v.reflect()("valid"), "validator.reflect() must return an invalid validator"
end sub
%>
