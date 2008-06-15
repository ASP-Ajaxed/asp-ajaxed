<%
'**************************************************************************************************************
'* Michal Gabrukiewicz Copyright (C) 2007 									
'* For license refer to the license.txt 									
'**************************************************************************************************************

'Uncomment the line if you want to use a configuration variable
'In 90% cases you don't need to change anything.
'(most of those settings are properties of classes and are initialized when the class is instantiated,
'therefore this means that you can change it after initialization also.)
'only the most common config variables are listed here. Check the documentation for even more configuration possibilities
'if you need a config var to be different in each environment then just place it into the envLIVE() and/or envDEV() sub

'**************************************************************************************************************

'The environment. choose between LIVE or DEV. no setting always results in DEV
	'AJAXED_ENVIRONMENT = "DEV"

'The vitual location to the ajaxed folder. must end AND start with an SLASH! (/)
	'AJAXED_LOCATION = "/ajaxed/"

'The text which appears when a callback is being performed and the user has to wait
	'AJAXED_LOADINGTEXT = "loading..."

'should the prototype JavaScript library be automatically loaded on every page?
	'AJAXED_LOADPROTOTYPEJS = true

'the caption for the errors used with error() method
	'AJAXED_ERRORCAPTION = "Erroro: "

'sets the response.buffer for each page.
	'AJAXED_BUFFERING = true

'ID of the form which should be used by default when no form is specified
	'AJAXED_FORMID = "frm"

'Advanced: should the codepage be set to the session directly?
'this is only recommended if you have IIS5 or lower which does not support
'setting the codepage directly. If you experience an codepage property error then turn this on.
	'AJAXED_SESSION_CODEPAGE = true

'EMAIL CONFIGS
	'AJAXED_MAILSERVER = "mail.domain.com"
	'AJAXED_EMAIL_SENDER = "email@domain.com"
	'AJAXED_EMAIL_SENDER_NAME = "your website"

'useful when developing
'if all emails should be sent to one email address (e.g. yours).
	'AJAXED_EMAIL_ALLTO = "youremail@host.com"

'turn dispatching of if you dont want to dispatch the messages 
	'AJAXED_EMAIL_DISPATCH = false

'DATABASE CONFIGS
'Should the database connection be established automatically on each page.
	'AJAXED_DBCONNECTION = true

'If you want to use a database with the ajaxed Library then configure a proper connectionstring
'to your database. Some are suggested below. uncomment if applicable
'MSSQL
	'AJAXED_CONNSTRING = "Driver={SQL Server};Server=localhost;Database=YOUR_DB;Uid=YOUR_USER;Pwd=YOUR_PASSWORD;"
'mySQL
	'AJAXED_CONNSTRING = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=YOUR_DB;User=YOUR_USER;Password=YOUR_PASSWORD;Option=3;"
'MS ACCESS 2008
	'AJAXED_CONNSTRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccess2007file.accdb;Persist Security Info=False;"
'MS ACCESS 2000-2003
	'AJAXED_CONNSTRING = "Driver={Microsoft Access Driver (*.mdb)};Dbq=test.mdb;"
'PostgreSQL
	'AJAXED_CONNSTRING = "Driver={PostgreSQL};Server=localhost;Port=5432;Database=YOUR_DB;Uid=YOUR_USER;Pwd=YOUR_PASSWORD;"
'Oracle
	'AJAXED_CONNSTRING = "Driver={Microsoft ODBC for Oracle};Server=localhost;Uid=YOUR_USER;Pwd=YOUR_PASSWORD;"
'Sqlite with ODBC - (odbc driver http://www.ch-werner.de/sqliteodbc/)
	'AJAXED_CONNSTRING = "DRIVER=SQLite3 ODBC Driver;Database=mydb.sqlite;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"

'CONFIG OVERRIDES ON THE ENVIRONMENTS (uncomment the sub)
'just place the config var into the sub if you want to have it different on that environment
'sub envDEV()
	'it makes sense to send all emails to you on the development 
	'AJAXED_EMAIL_ALLTO = "youremail@domain.com"
	'or dont dispatch emails at all
	'AJAXED_EMAIL_DISPATCH = false
'end sub

'sub envLIVE()
'end sub
%>