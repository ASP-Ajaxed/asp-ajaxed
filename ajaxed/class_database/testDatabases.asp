<%
'connection strings for all the test databases
TEST_DB_ACCESS = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath(lib.path("class_database/test.mdb")) & ";"
TEST_DB_SQLITE = "DRIVER=SQLite3 ODBC Driver;Database=" & server.mappath(lib.path("class_database/test.sqlite")) & ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"
%>