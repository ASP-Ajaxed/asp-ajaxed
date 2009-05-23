<%
'connection strings for all the test databases
TEST_DB_ACCESS 		= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath(lib.path("class_database/test.mdb")) & ";"
TEST_DB_SQLITE 		= "DRIVER=SQLite3 ODBC Driver;Database=" & server.mappath(lib.path("class_database/test.sqlite")) & ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"
TEST_DB_MSSQL 		= "Provider=sqloledb;Data Source=(local);Initial Catalog=ajaxedtest;Trusted_Connection=yes"
TEST_DB_MYSQL 		= "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=ajaxedtest;User=root;Option=3;"
TEST_DB_POSTGRESQL 	= "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=ajaxedtests;Uid=username;Pwd=password;"
%>