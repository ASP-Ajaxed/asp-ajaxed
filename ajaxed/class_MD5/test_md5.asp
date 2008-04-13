<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="../class_MD5/md5.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set h = new MD5
	tf.assert h.hash("md5isnice") = "712a84bd75495d12946129e2bc34d4b7", "hash must be generated correctly."
end sub
%>