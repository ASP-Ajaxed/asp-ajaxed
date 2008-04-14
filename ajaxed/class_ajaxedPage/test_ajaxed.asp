<% AJAXED_CONNSTRING = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & server.mappath("../class_database/test.mdb") & ";" %>
<!--#include file="../class_testFixture/testFixture.asp"-->
<%
'This page performs several tests in order to check the
'correct functionality of AJAXED.
'dont change anything!!!

initialized = false

class foo
	public name, age, country, favs
	function reflect()
		set reflect = lib.newDict(empty)
		reflect.add "name", me.name
		reflect.add "age", me.age
		reflect.add "Country", me.country
		reflect.add "favs", me.favs
	end function
end class

class fooUnknown
end class

set p = new AjaxedPage
p.plain = true
p.dbconnection = true
p.onlyDev = true

set tf = new TestFixture
tf.run()

sub test_1
	p.draw()
end sub


sub init()
	if request.querystring = "" and p.QS(empty) <> "" then p.error("No querystring should be here")
	
	initialized = true
end sub

sub callback(action)
	if not p.isPostback() then p.error("Must be postpack!")
	if not initialized then p.error("init() has not been exectuted")
	
	select case action
		case "bool"
			p.returnValue "t", true
			p.returnValue "f", false
		case "number"
			p.return(10)
		case "float"
			p.return(10.99)
		case "nothings"
			p.return(array(null, nothing, empty))
		case "string"
			p.returnValue "ABC", "ABC"
			p.returnValue "special", "äöü'""?!\/[]_²§()~"
		case "object"
			set f = new foo
			f.name = "michal"
			f.age = 26
			f.country = "austria"
			p.returnValue "custom", f
			p.returnValue "unknown", new fooUnknown
		case "object_simple"
			set f = new foo
			f.name = "michal"
			f.age = 26
			f.country = "austria"
			f.favs = array(2, 3, 4, "s")
			p.return(f)
		case "dict"
			p.return((new foo).reflect())
		case "dict2"
			set d1 = (new foo).reflect()
			p.returnValue "d1", d1
			set d2 = server.createObject("scripting.dictionary")
			d2.add "arr", array(2, 3, false)
			d2.add "d1", d1
			p.returnValue "d2", d2
		case "rs"
			set rs = db.getRecordset("SELECT * FROM person ORDER BY created_on")
			p.return(rs)
		case "rs_count"
			set rs = db.getUnlockedRecordset("SELECT * FROM person ORDER BY created_on")
			p.returnValue "data", rs
			p.returnValue "count", rs.recordCount
		case "empty_rs"
			set rs = db.getRecordset("SELECT * FROM person WHERE firstname = 'asdajs' ORDER BY created_on")
			p.return(rs)
	end select
end sub

sub main() %>

	<%
	tf.assertEqual lib.page.loadPrototypeJS, true, "prototypeJS must be loaded by default"
	tf.assert not p.isPostback(), "No postback possible here!"
	tf.assert initialized, "init() has not been exectuted"
	%>
	
	<script>
		t_dict2 = function(r) {
			l(r.d1.name == '');
			l(r.d1.age == '');
			l(r.d2.arr[0] == 2);
			l(r.d2.arr[1] == 3);
			l(r.d2.arr[2] == false);
			l(r.d2.d1.Country == '');
		}
		t_string = function(r) {
			l(r.ABC == 'ABC');
			l(r.special == 'äöü\'"?!\\/[]_²§()~');
		}
		t_bool = function(r) {
			l(r.t == true);
			l(r.f == false);
		}
		t_number = function(r) {
			l(r == 10);
		}
		t_float = function(r) {
			l(r == 10.99);
		}
		t_object = function(r) {
			l(r.unknown == "fooUnknown");
			l(r.custom.name == "michal");
			l(r.custom.age == 26);
			l(r.custom.Country == "austria");
		}
		t_object_simple = function(r){
			l(r.name == "michal");
			l(r.age == 26);
			l(r.Country == "austria");
			l(r.favs[3] == 's');
		}
		t_dict = function(r) {
			l(r.name == "");
			l(r.age == "");
			l(r.Country == "");
		}
		t_nothings = function(r) {
			for (i = 0; i < r.length; i++) {
				l(!r[0]);
			}
		}
		t_rs = function(r) {
			l(r.length == 2);
			l(r[0].firstname == 'michal');
			l(r[0].lastname == 'gabrukiewicz');
			l(r[0].age == 26);
			l(r[0].cool == false);
			l(r[0].created_on != '');
			l(!r[0].some_null);
			l(r[0].some_decimal == 3.789);
			
			l(r[1].firstname == 'cool');
			l(r[1].lastname == 'and the gang');
			l(r[1].age == 48);
			l(r[1].cool == true);
			l(r[1].created_on != '');
			l(!r[1].some_null);
			l(r[1].some_decimal == 3.345);
		}
		t_empty_rs = function(r) {
			l(r.length == 0);
		}
		t_rs_count = function(r) {
			l(r.data.length == 2);
			l(r.count == 2);
			l(r.data[1].firstname == 'cool');
		}
		
		l = function(msg) {
			if (!msg) $('log').innerHTML += "<br>" + msg;
		}
	</script>

	<div style="font-family:courier" id="log"></div>
	
	<script>
		var u = '/ajaxed/class_ajaxedPage/test_ajaxed.asp';
		ajaxed.callback('bool', t_bool, null, null, u);
		ajaxed.callback('string', t_string, null, null, u);
		ajaxed.callback('number', t_number, null, null, u);
		ajaxed.callback('float', t_float, null, null, u);
		ajaxed.callback('nothings', t_nothings, null, null, u);
		ajaxed.callback('object', t_object, null, null, u);
		ajaxed.callback('dict', t_dict, null, null, u);
		ajaxed.callback('dict2', t_dict2, null, null, u);
		ajaxed.callback('object_simple', t_object_simple, null, null, u);
		ajaxed.callback('rs', t_rs, null, null, u);
		ajaxed.callback('empty_rs', t_empty_rs, null, null, u);
		ajaxed.callback('rs_count', t_rs_count, null, null, u);
	</script>	

<% end sub %>