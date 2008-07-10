<!--#include file="../class_testFixture/testFixture.asp"-->
<% AJAXED_CONNSTRING = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath("../class_database/test.mdb") & ";" %>
<%
class Person
	public firstname	''[string] firstname
	public lastname ''[string] lastname
	public favNumbers   ''[array] persons favorite numbers
	public function reflect()
		set reflect = server.createObject("scripting.dictionary")
		with reflect
			.add "firstname", firstname
			.add "lastname", lastname
			.add "favNumbers", favNumbers
		end with
	end function
end class

set db = new Database

set tf = new TestFixture
tf.run()


sub test_1()
	'just for completion, that it is a test and generates output
	db.open(AJAXED_CONNSTRING)
end sub


set p = new Person
p.firstname = "John"
p.lastname = "Doe"
p.favNumbers = array(2, 7, 234)
set aDict = server.createObject("scripting.dictionary")
aDict.add "a", 2.2342348
aDict.add "b", array(7, 8, array(9, 5))
aDict.add "c", p
aDict.add "f", db.getRecordset("SELECT * FROM person ORDER BY created_on")
aDict.add "g", db.getRecordset("SELECT * FROM person WHERE firstname = 'sbas' ORDER BY created_on")

set rs = db.getRecordset("SELECT * FROM person ORDER BY created_on")
set empty_rs = db.getRecordset("SELECT * FROM person WHERE firstname = 'sbas' ORDER BY created_on")
%>
<script>
	var assertions = 1;
	assert = function(bool, msg) {
		var m = 'Success.';
		if (!bool) {
			alert('Failed: ' + ((msg) ? msg : 'no error help message given.'));
		}
		assertions++;
	}
	
	<%
	set s = lib.newDict(empty)
	s.add "germanumlaute", "äöüÄÖÜß"
	s.add "newline", vbNewLine
	s.add "specials", """\/"
	s.add "asian", ""
	%>
	var jsn = <%= (new JSON).toJSON(empty, s, false) %>;
	assert(jsn.germanumlaute == 'äöüÄÖÜß', 'german umlaute should work');
	assert(jsn.newline == '\r\n', 'new lines should work');
	assert(jsn.specials == '"\\/', 'special chars should work');
	
	var jsn = <%= (new JSON).toJSON("p", p, false) %>;
	assert(jsn.p.favNumbers[0] == 2);
	assert(jsn.p.favNumbers[2] == 234);
	assert(jsn.p.firstname == 'John');
	assert(jsn.p.lastname == 'Doe');
	
	var jsn = <%= (new JSON).toJSON("a", aDict, false) %>;
	assert(jsn.a.a == 2.2342348, 'Floats need to be recognized.');
	assert(jsn.a.b[2][0] == 9);
	assert(jsn.a.c.lastname == 'Doe');
	assert(jsn.a.c.favNumbers[2] == 234);
	assert(jsn.a.f[1].lastname == 'and the gang');
	assert(jsn.a.g.length == 0, 'empty RS within nested within a dictionary');
	
	<% rs.movefirst %>
	var jsn = <%= (new JSON).toJSON("rs", rs, false) %>;
	assert(jsn.rs.length == 2, '2 records must exist in person table')
	assert(jsn.rs[0].firstname == 'michal');
	assert(jsn.rs[0].lastname == 'gabrukiewicz');
	assert(jsn.rs[0].cool == false);
	assert(!jsn.rs[0].some_null);
	assert(jsn.rs[0].some_decimal == 3.789);
	assert(jsn.rs[1].firstname == 'cool');
	assert(jsn.rs[1].lastname == 'and the gang');
	
	var jsn = <%= (new JSON).toJSON("rs", empty_rs, false) %>;
	assert(jsn.rs.length == 0);
	
</script>
