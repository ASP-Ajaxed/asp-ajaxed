<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="../class_SHA256/sha256.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set h = new SHA256
	tf.assert h.hash("The quick brown fox jumps over the lazy dog") = "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592", "hash must be generated correctly."
end sub
%>