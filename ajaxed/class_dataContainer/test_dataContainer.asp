<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="../class_database/testDatabases.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

sub test_1()
	set d = new DataContainer
	on error resume next
		d.contains("x")
		failed = err <> 0
	on error goto 0
	if not failed then tf.fail("It should not be possible to call a method without using the newWith() function")
end sub

sub test_2()
	set d = (new DataContainer)(array(1, 2, 3))
	tf.assert d.contains(2), "contains()"
	tf.assert d.contains(array(1, 3)), "contains() with an array"
	tf.assert d.contains(array(2)), "contains() with an array"
	tf.assertNot d.contains(array(4)), "contains() with an array"
	tf.assertNot d.contains(array(2, 10)), "contains() with an array"
	tf.assert d.contains("3"), "contains()"
	tf.assert d.contains(2.0), "contains()"
	set d = (new DataContainer)(array("test", "who", 3))
	tf.assert d.contains("3"), "contains()"
	tf.assertNot lib.contains("3", "3"), "lib.contains()"
	set d = (new DataContainer)(array(1, 2, 3, 10))
	tf.assertNot d.contains(11), "contains()"
	tf.assert (((new DataContainer)(array(1, 2, 3))).contains(3)), "noooo"
	
	set d = (new DataContainer)(lib.newDict(empty))
	d.data.add 1, "someting"
	d.data.add 2, ""
	d.data.add "3", "yeah"
	d.data.add "not", empty
	tf.assert d.contains(2), "contains() with dictionary"
	tf.assert d.contains("2"), "contains() with dictionary"
	tf.assertNot d.contains(0), "contains() with dictionary"
	tf.assertNot d.contains(3.2), "contains() with dictionary"
	tf.assert d.contains("not"), "contains() with dictionary"
end sub

sub test_3()
	set d = (new DataContainer)(array(1, 2, 3))
	tf.assertEqual 1, d.first, "first property does not work with array"
	tf.assertEqual 3, d.last, "last property does not work with array"
	tf.assertEqual 3, d.count, "count property does not work with array"
	tf.assertEqual "1, 2, 3", d.toString(", "), "toString() does not work with array"
	tf.assertEqual "", ((new DataContainer)(array())).toString(", "), "toString() does not work with array"
	
	'dictionary tests
	set d = (new DataContainer)(lib.newDict(empty))
	d.data.add 1, "jack"
	d.data.add 2, "madonna"
	d.data.add 3, "axel"
	tf.assertEqual 1, d.first, "first property with dictionary does not work"
	tf.assertEqual 3, d.last, "last property with dictionary does not work"
	tf.assertEqual 3, d.count, "count property with dictionary does not work"
	tf.assertEqual "1 -> jack, 2 -> madonna, 3 -> axel", d.toString(array(", ", " -> ")), "toString() does not work on dictionary"
	tf.assertEqual "1 jack, 2 madonna, 3 axel", d.toString(", "), "toString() does not work on dictionary"
	tf.assertEqual "", ((new DataContainer)(lib.newDict(empty))).toString(", "), "toString() does not work on dictionary"
	
	'recordset tests
	db.open(TEST_DB_ACCESS)
	set d = (new DataContainer)(db.getRS("SELECT * FROM person", empty))
	tf.assertNothing d, "newWith() must return nothing on db.getRS() recordsets"
	
	set d = (new DataContainer)(db.getUnlockedRS("SELECT * FROM person WHERE 1 = 2", empty))
	tf.assertEqual 0, d.count, "count property with recordset does not work"
	tf.assertEqual empty, d.last, "last property with recordset does not work"
	
	set d = (new DataContainer)(db.getUnlockedRS("SELECT * FROM person", empty))
	tf.assertEqual 2, d.count, "count property with recordset does not work"
	tf.assertEqual 2, d.last, "last property with recordset does not work"
	
	d.data.moveNext()
	pos = d.data.absolutePosition
	tf.assertEqual 1, d.first, "first property with recordset does not work after moving the position"
	tf.assert pos = d.data.absolutePosition, "position of recordset must be restored after using first property"
	
	d.data.moveNext()
	pos = d.data.absolutePosition
	tf.assertEqual 2, d.last, "last property with recordset does not work after moving the position"
	tf.assert pos = d.data.absolutePosition, "position of recordset must be restored after using last property"
end sub

sub test_4()
	set d = (new DataContainer)(lib.range(1, 100, 1))
	set pages = (new DataContainer)(d.paginate(10, 5, 3))
	
	tf.assert pages.contains(array(4, 5, 6)), "paginate() does not work"
	tf.assertEqual "...", pages.first, "paginate() does not work"
	tf.assertEqual "...", pages.last, "paginate() does not work"
	tf.assertEqual 5, pages.count, "paginate() does not work"
	
	set pages = (new DataContainer)(d.paginate(10, 10, 10))
	tf.assertNot pages.contains("..."), "paginate() does not work"
	tf.assertEqual 10, pages.count, "paginate() does not work"
	
	set pages = (new DataContainer)(d.paginate(10, 10, 3))
	tf.assertNotEqual "...", pages.last, "paginate() does not work"
	
	set pages = (new DataContainer)(d.paginate(50, 1, 3))
	tf.assert pages.contains(array(1, 2)), "paginate() does not work"
	tf.assertNot pages.contains("..."), "paginate() does not work"
	
	set pages = (new DataContainer)(d.paginate(10, empty, 3))
	tf.assertNot pages.last = "...", "paginate() does not work"
	tf.assert pages.first = "...", "paginate() does not work"
	
	'empty DataContainers
	set d = (new DataContainer)(array())
	set pages = (new DataContainer)(d.paginate(10, 5, 3))
	tf.assertEqual 1, pages.count, "paginate() on empty DataContainers"
	tf.assertEqual 1, pages.first, "paginate() on empty DataContainers"
	tf.assertEqual 1, pages.last, "paginate() on empty DataContainers"
	
	'recordset
	db.open(TEST_DB_ACCESS)
	set d = (new DataContainer)(db.getUnlockedRS("SELECT * FROM person", empty))
	set pages = (new DataContainer)(d.paginate(10, 5, 3))
	tf.assertEqual 1, pages.count, "paginate() with a recordset"
end sub

sub test_4()
	tf.assertEqual array(1, 2, 3, 4), ((new DataContainer)(array(1, 2, 3, 4, 1, 2, 3))).unique(false), "unique() is not working"
	tf.assertEqual array(), ((new DataContainer)(array())).unique(false), "unique() is not working"
	tf.assertEqual array("jack", "JACK"), ((new DataContainer)(array("jack", "JACK", "jack"))).unique(true), "unique() is not working"
	tf.assertEqual array("jack"), ((new DataContainer)(array("jack", "JACK", "jack"))).unique(false), "unique() case sensitive is not working"
	tf.assertEqual array("michaL"), ((new DataContainer)(array("michaL", "michal", "MICHAL"))).unique(false), "unique() case sensitive is not working"
end sub

'test adding
sub test_5()
	set dc = (new DataContainer)(array())
	dc.add("some value")
	tf.assertEqual 0, uBound(dc.data), "array length wrong"
	tf.assertEqual "some value", dc.last, "array field not addedd correctly (first value)"
	tf.assertEqual "some value", dc.first, "array field not addedd correctly (last value)"
	
	'same with a dictionary
	set dc = (new DataContainer)(["D"](empty))
	dc.add(array(1, "first"))
	tf.assertEqual 1, dc.count, "dictionary adding not working (count)"
	tf.assertEqual "first", dc.data(1), "dictionary add not working (get)"
	dc.add(array(2, "second"))
	tf.assertEqual 2, dc.count, "dictionary adding of second value not working (count)"
	tf.assert dc.contains(2), "dictionary adding not working (contains)"
	tf.assertEqual "second", dc.data(2), "dictionary get not working (after adding)"
	dc.add(array(2, "changed"))
	tf.assertEqual 2, dc.count, "count should stay 2 as key 2 already exists"
	tf.assertEqual "changed", dc.data(2), "should have updated the item with key 2"
	
	'now with a recordset
	set dc = (new DataContainer)(["R"](array( _
		array("firstname", "lastname") _
	)))
	dc.add(array("John", "Doe"))
	tf.assertEqual "John", dc.data("firstname").value, "record not added correctly (firstname)"
	tf.assertEqual "Doe", dc.data("lastname").value, "record not added correctly (lastname)"
	
	'chaining
	set dc = (new DataContainer)(array())
	dc.add("one").add("two").add("three")
	tf.assertEqual 3, dc.count, "chaining not working"
end sub
%>