<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="testDatabases.asp"-->
<%
AJAXED_CONNSTRING = TEST_DB_ACCESS

currentDbmsName = "unknown"
set tf = new TestFixture
tf.run()

'test with default DB (ACCESS)
sub test_1()
	testDbms empty, "MS Access"
end sub

'SQLITE test
sub test_2()
	testDbms TEST_DB_SQLITE, "SQLite"
end sub

'MYSQL test
sub test_3()
	testDbms TEST_DB_MYSQL, "MySQL"
end sub

'MSSQL test
sub test_4()
	testDbms TEST_DB_MSSQL, "MS Sql Server"
end sub

sub test_5()
	testDbms TEST_DB_POSTGRESQL, "PostgreSQL"
end sub

sub testDbms(connectionString, dbmsName)
	currentDbmsName = dbmsName
	if isEmpty(connectionString) then
		db.openDefault()
	else
		if not openDB(connectionString, dbmsName) then exit sub
	end if
	db.connection.beginTrans()
	on error resume next
		performDBOperations()
		if err <> 0 then tf.fail currentDbmsName & " failed to test (" & err.description & ")"
	on error goto 0
	db.connection.rollbackTrans()
	db.close()
	tf.assertNothing db.connection, "db.connection is nothing"
end sub

function openDB(connStr, dbmsName)
	openDB = true
	on error resume next
		db.open(connStr)
		if err <> 0 then
			lib.logger.info err.description
			on error goto 0
			openDB = false
			tf.info(str.format("Could not open {0} Test Database. Therefore {0} was not tested. This test is only necessary if you develop the ajaxed core otherwise relax.", dbmsName))
			err.clear()
		end if
	on error goto 0
end function

sub performDBOperations()
	'Note: Those asserts assume existing data within the "tested" database (db.connection)
	'Please refer to the test.sqlite database to see the test fixture data
	
	tf.assertInstanceOf "connection", db.connection, assertMsg("db.connection")
	tf.assertInstanceOf "recordset", db.getRS("SELECT * FROM person", empty), assertMsg("db.getRecordset")
	tf.assertEqual 2, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, "db.getUnlockedRecordset"
	tf.assert db.getRS("SELECT * FROM person WHERE id = 0", empty).eof, assertMsg("db.getRecordset")
	
	ID = 1
	ID = db.insert("person", array("firstname", null, "lastname", "kett" & str.clone("s", 300), "age", 20))
	
	tf.assertEqual 255, len(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("lastname")), _
		assertMsg("lastname must have been trimmed to the maximum accepted length of the column")
	tf.assert isNull(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("firstname")), _
		assertMsg("firstname must be null when set to null")
	tf.assert not isEmpty(db.getRS("SELECT * FROM person WHERE id = {0}", ID)("firstname")), _
		assertMsg("firstname cannot be empty when defined as null")
	
	tf.assert ID > 0, assertMsg("db.insert does not return a correct ID")
	tf.assertEqual 3, db.getUnlockedRS("SELECT * FROM person", empty).recordcount, _
		assertMsg("db.getUnlockedRecordset does not return the correct recordcount")
	
	db.update "person", array("lastname", "Kett", "age", 69), ID
	
	tf.assertEqual 69, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), assertMsg("db.getScalar")
	
	tf.assertEqual 3, db.count("person", empty), _
		assertMsg("we should have three persons")
	db.delete "person", id
	
	tf.assertEqual 2, db.count("person", empty), _
		assertMsg("in the end there should only be two persons")
	coolies = db.count("person", "cool = 1")
	tf.assertEqual 1, coolies, _
		assertMsg("only one should be cool but was " & coolies)
	db.toggle "person", "cool", "cool = 0"
	tf.assertEqual 2, db.count("person", "cool = 1"), _
		assertMsg("togglin all not cool to cool should result in 2 cool")
	db.toggle "person", "cool", empty
	tf.assertEqual 0, db.count("person", "cool = 1"), _
		assertMsg("in the end no cool people")
	
	set RS = db.getRS("SELECT id FROM person", empty)
	id = RS("id")
	tf.assert not RS.eof, assertMsg("should have at least one person")
	set RS = nothing
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", id).eof, _
		assertMsg("db.getRS should take only one param as argument")
	tf.assert not db.getRS("SELECT * FROM person WHERE id = {0}", array(id)).eof, _
		assertMsg("db.getRS should also take an array as arg")
	
	'batch update
	tf.assertEqual 2, db.update("person", array("firstname", "batch"), empty), _
		assertMsg("update() returns wrong value on batch")
	tf.assertEqual 2, db.count("person", "firstname = 'batch'"), _
		assertMsg("Batch update no.1 failed")
	tf.assertEqual 1, db.update("person", array("firstname", "notbatch"), 1), _
		assertMsg("update() returns wrong value")
	tf.assertEqual 1, db.count("person", "firstname = 'batch'"), _
		assertMsg("Batch update no.2 failed")
	tf.assertEqual 1, db.count("person", "firstname = 'notbatch'"), _
		assertMsg("Batch update no.3 failed")
	tf.assertEqual 0, db.update("person", array("firstname", "notbatch"), "firstname = 'batman'"), _
		assertMsg("update() returned wrong value")
	
	'test insertOrUpdate()
	ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 40), 0)
	tf.assertEqual 40, db.getScalar("SELECT age FROM person WHERE id = " & id, 0), _
		assertMsg("insertOrUpdate no.1 problem")
	
	ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "other", "age", 88), "id = -1")
	tf.assertEqual 88, db.getScalar("SELECT age FROM person WHERE id = " & id, 0), _
		assertMsg("insertOrUpdate no.2 problem")
	
	mirID = db.getScalar("SELECT id FROM person WHERE firstname = 'mirjam'", 0)
	if mirID <= 0 then tf.fail(assertMsg("Must get one of the records of the two existing 'Mirjam'"))
	tf.assertEqual 40, db.getScalar("SELECT age FROM person WHERE firstname = 'mirjam' ORDER BY age ASC", 0), _
		assertMsg("db.getScalar() must return the first record even if there are actually more returned.")
	mirID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "checker", "age", 70), mirID)
	tf.assertEqual 70, db.getScalar("SELECT age FROM person WHERE firstname = 'mirjam' AND age = 70", 0), _
		assertMsg("insertOrUpdate() no.1 should really update the values if it says that it did the update.")
	tf.assertEqual "checker", db.getScalar("SELECT lastname FROM person WHERE firstname = 'mirjam' AND age = 70", ""), _
		assertMsg("insertOrUpdate() no.2 should really update the values if it says that it did the update.")
	
	on error resume next
		ID = db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 26), empty)
		failed = err <> 0
	on error goto 0
	tf.assert failed, assertMsg("update() must fail with empty condition")
	
	tf.assertEqual 0, db.insertOrUpdate("person", array("firstname", "mirjam", "lastname", "gabru", "age", 27), ID), _
		assertMsg("insertOrUpdate() must return 0 on updates")
	tf.assertEqual 27, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), _
		assertMsg("insertOrUpdate no.3 problem")
	tf.assertEqual 2, db.count("person", "firstname = 'mirjam'"), _
		assertMsg("should be still 2")
	
	tf.assertEqual 0, db.insertOrUpdate("person", array("age", 99), "firstname = 'mirjam'"), _
		assertMsg("insertOrUpdate() must return 0 on updates")
	tf.assertEqual 99, db.getScalar("SELECT age FROM person WHERE id = " & ID, 0), _
		assertMsg("insertOrUpdate no.4 problem")
	tf.assertEqual 2, db.count("person", "firstname = 'mirjam'"), _
		assertMsg("should be even now 2")
	tf.assertEqual 2, db.count("person", "age = 99"), _
		assertMsg("should be 2 with age of 99")
end sub

function assertMsg(msg)
	assertMsg = msg & "(" & currentDbmsName & ")"
end function
%>