<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="testDatabases.asp"-->
<%
AJAXED_CONNSTRING = TEST_DB_ACCESS

set tf = new TestFixture
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
	if not openDB(TEST_DB_SQLITE, "SQLite") then exit sub
	performDBOperations()
	db.close()
end sub

'MYSQL test
sub test_3()
	if not openDB(TEST_DB_MYSQL, "MySQL") then exit sub
	performDBOperations()
	db.close()
end sub

'MSSQL test
sub test_4()
	if not openDB(TEST_DB_MSSQL, "MS Sql Server") then exit sub
	performDBOperations()
	db.close()
end sub

function openDB(connStr, dbName)
	openDB = true
	on error resume next
		db.open(connStr)
		if err <> 0 then
			lib.logger.info err.description
			on error goto 0
			openDB = false
			tf.info(str.format("Could not open {0} Test Database. Therefore {0} was not tested. This test is only necessary if you develop the ajaxed core otherwise relax.", dbName))
			err.clear()
		end if
	on error goto 0
end function

sub performDBOperations()
	db.connection.beginTrans()
	
	tf.assertInstanceOf "connection", db.connection, "db.connection"
	tf.assertInstanceOf "recordset", db.getRS("SELECT * FROM person", empty), "db.getRecordset"
	tf.assertEqual 2, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, "db.getUnlockedRecordset"
	tf.assert db.getRS("SELECT * FROM person WHERE id = 0", empty).eof, "db.getRecordset"
	
	ID = 1
	ID = db.insert("person", array("firstname", null, "lastname", "kett" & str.clone("s", 300), "age", 20))
	
	tf.assertEqual 255, len(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("lastname")), "lastname must have been trimmed to the maximum accepted length of the column"
	tf.assert isNull(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("firstname")), "firstname must be null when set to null"
	tf.assert not isEmpty(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("firstname")), "firstname cannot be empty when defined as null"
	
	tf.assert ID > 0, "db.insert does not return a correct ID"
	tf.assertEqual 3, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, "db.getUnlockedRecordset"
	
	db.update "person", array("lastname", "Kett", "age", 69), ID
	
	tf.assertEqual 69, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), "db.getScalar"
	
	tf.assertEqual 3, db.count("person", empty), "we should have three persons"
	db.delete "person", id
	
	tf.assertEqual 2, db.count("person", empty), "in the end there should only be two persons"
	coolies = db.count("person", "cool = 1")
	tf.assertEqual 1, coolies, "only one should be cool but was " & coolies
	db.toggle "person", "cool", "cool = 0"
	tf.assertEqual 2, db.count("person", "cool = 1"), "togglin all not cool to cool should result in 2 cool"
	db.toggle "person", "cool", empty
	tf.assertEqual 0, db.count("person", "cool = 1"), "in the end no cool people"
	
	set RS = db.getRS("SELECT id FROM person", empty)
	id = RS("id")
	tf.assert not RS.eof, "should have at least one person"
	set RS = nothing
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", id).eof, "db.getRS should take only one param as argument"
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", array(id)).eof, "db.getRS should also take an array as arg"
	
	'batch update
	tf.assertEqual 2, db.update("person", array("firstname", "batch"), empty), "update() returns wrong"
	tf.assertEqual 2, db.count("person", "firstname = 'batch'"), "Batch update failed"
	tf.assertEqual 1, db.update("person", array("firstname", "notbatch"), 1), "update() returns wrong"
	tf.assertEqual 1, db.count("person", "firstname = 'batch'"), "Batch update failed"
	tf.assertEqual 1, db.count("person", "firstname = 'notbatch'"), "Batch update failed"
	tf.assertEqual 0, db.update("person", array("firstname", "notbatch"), "firstname = 'batman'"), "update() returns wrong"
	
	'test insertOrUpdate()
	ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 40), 0)
	tf.assertEqual 40, db.getScalar("SELECT age FROM person WHERE id = " & id, 0), "insertOrUpdate problem"
	
	ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 88), "id = -1")
	tf.assertEqual 88, db.getScalar("SELECT age FROM person WHERE id = " & id, 0), "insertOrUpdate problem"
	
	mirID = db.getScalar("SELECT id FROM person WHERE firstname = 'mirjam'", 0)
	if mirID <= 0 then tf.fail("Must get the record of 'Mirjam'")
	mirID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "checker", "age", 70), mirID)
	tf.assertEqual 70, db.getScalar("SELECT age FROM person WHERE firstname = 'mirjam'", 0), "insertOrUpdate() should really update the values if it says that it did the update."
	tf.assertEqual "checker", db.getScalar("SELECT lastname FROM person WHERE firstname = 'mirjam'", ""), "insertOrUpdate() should really update the values if it says that it did the update."
	
	on error resume next
		ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 26), empty)
		failed = err <> 0
	on error goto 0
	tf.assert failed, "update() must fail with empty condition"
	
	tf.assertEqual 0, db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 27), ID), "insertOrUpdate() must return 0 on updates"
	tf.assertEqual 27, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), "insertOrUpdate problem"
	tf.assertEqual 2, db.count("person", "firstname = 'mirjam'"), "should be still 2"
	
	tf.assertEqual 0, db.insertOrUpdate("person", array("age", 99), "firstname = 'mirjam'"), "insertOrUpdate() must return 0 on updates"
	tf.assertEqual 99, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), "insertOrUpdate problem"
	tf.assertEqual 2, db.count("person", "firstname = 'mirjam'"), "should be even now 2"
	tf.assertEqual 2, db.count("person", "age = 99"), "should be 2 with age of 99"
	
	'rollback all changes
	db.connection.rollbackTrans()
end sub
%>