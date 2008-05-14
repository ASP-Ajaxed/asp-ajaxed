<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set output = new StringBuilder
	for i = 1 to 10
		output("x")
	next
	tf.assertEqual "xxxxxxxxxx", output.toString(), "Stringbuilder does not work"
end sub
%>