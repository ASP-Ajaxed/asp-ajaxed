<!--#include file="../class_testFixture/testFixture.asp"-->
<% AJAXED_CONNSTRING = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath("test.mdb") & ";" %>
<%
set tf = new TestFixture
tf.run()

sub test_1()
	db.openDefault()
	db.connection.beginTrans()
	tf.assertInstanceOf "connection", db.connection, "db.connection"
	tf.assertInstanceOf "recordset", db.getRecordset("SELECT * FROM person"), "db.getRecordset"
	tf.assertEqual 2, db.getUnlockedRecordset("SELECT * FROM person").recordcount, "db.getUnlockedRecordset"
	tf.assert db.getRecordset("SELECT * FROM person WHERE id = 0").eof, "db.getRecordset"
	
	ID = 1
	ID = db.insert("person", array("firstname", "jack", "lastname", "kett", "age", 20))
	
	tf.assert ID > 0, "db.insert"
	tf.assertEqual 3, db.getUnlockedRecordset("SELECT * FROM person").recordcount, "db.getUnlockedRecordset"
	
	db.update "person", array("lastname", "Kett", "age", 69), ID
	
	tf.assertEqual 69, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), "db.getScalar"
	
	db.delete "person", id
	
	tf.assertEqual 2, db.count("person", empty), "in the end there should only be two persons"
	tf.assertEqual 1, db.count("person", "cool = true"), "only one should be cool"
	db.toggle "person", "cool", "cool = false"
	tf.assertEqual 2, db.count("person", "cool = true"), "togglin all not cool to cool should result in 2 cool"
	db.toggle "person", "cool", empty
	tf.assertEqual 0, db.count("person", "cool = true"), "in the end no cool people"
	
	tf.assertEqual 14, db.numberOfDBAccess, "db.numberOfDBAccess"
	db.connection.rollbackTrans()
	db.close()
	tf.assertNothing db.connection, "db.connection is nothing"
end sub
%>