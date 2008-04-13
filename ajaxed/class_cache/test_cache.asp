<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="../class_cache/cache.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set cash = new Cache
	with cash
		.name = "test"
		.interval = "s"
		.intervalValue = 10
		.store "val", 10
		tf.assert str.parse(.getItem("val"), 0) = 10, ""
	end with
end sub
%>
