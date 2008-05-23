<!--#include file="../class_testFixture/testFixture.asp"-->
<% AJAXED_CONNSTRING = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath("test.mdb") & ";" %>
<%
set tf = new TestFixture
tf.debug = true
tf.run()

'tests with default DB (ACCESS)
sub test_1()
	db.openDefault()
	performDBOperations()
	db.close()
	tf.assertNothing db.connection, "db.connection is nothing"
end sub

'SQLITE test
sub test_2()
	db.open("DRIVER=SQLite3 ODBC Driver;Database=" & server.mappath("test.sqlite") & ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;")
	performDBOperations()
	db.close()
end sub

sub performDBOperations()
	db.connection.beginTrans()
	
	tf.assertInstanceOf "connection", db.connection, "db.connection"
	tf.assertInstanceOf "recordset", db.getRS("SELECT * FROM person", empty), "db.getRecordset"
	tf.assertEqual 2, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, "db.getUnlockedRecordset"
	tf.assert db.getRS("SELECT * FROM person WHERE id = 0", empty).eof, "db.getRecordset"
	
	ID = 1
	ID = db.insert("person", array("firstname", "jack", "lastname", "kett", "age", 20))
	
	tf.assert ID > 0, "db.insert does not return a correct ID"
	tf.assertEqual 3, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, "db.getUnlockedRecordset"
	
	db.update "person", array("lastname", "Kett", "age", 69), ID
	
	tf.assertEqual 69, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), "db.getScalar"
	
	db.delete "person", id
	
	tf.assertEqual 2, db.count("person", empty), "in the end there should only be two persons"
	coolies = db.count("person", "cool = 1")
	tf.assertEqual 1, coolies, "only one should be cool but was " & coolies
	db.toggle "person", "cool", "cool = 0"
	tf.assertEqual 2, db.count("person", "cool = 1"), "togglin all not cool to cool should result in 2 cool"
	db.toggle "person", "cool", empty
	tf.assertEqual 0, db.count("person", "cool = 1"), "in the end no cool people"
	
	set RS = db.getRS("SELECT id FROM person", empty)
	tf.assert not RS.eof, "should have at least one person"
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", RS("id")).eof, "db.getRS should take only one param as argument"
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", array(RS("id"))).eof, "db.getRS should also take an array as arg"
	
	db.connection.rollbackTrans()
end sub
%>