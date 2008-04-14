<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="RSS.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set r = new RSS
	r.url = "http://www.webdevbros.net/feed/"
	r.load()
	tf.assertNot r.failed, "webdevbros RSS loading failed"
	tf.assert r.title <> "", "title of the feed must be set"
	tf.assert r.items.count > 0, "there must be some blog posts"
end sub

sub test_2()
	set r = new RSS
	r.url = "http://www.grafix.at/ajaxed/console/version.asp"
	r.load()
	tf.assertNot r.failed, "ajaxed updater is not available"
	tf.assertEqual "ajaxed version feed", r.title, "title of the feed does not match"
	tf.assert r.description <> "", "version released text must be there"
end sub
%>