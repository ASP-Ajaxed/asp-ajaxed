<%
'the JSON class is used also as a standalone component on webdevbros.net
'and it also supports option explicit .. thus those tests
option explicit
%>
<!--#include file="JSON.asp"-->
<%
dim output, data, multi(1, 1), dict, RS
multi(0, 0) = "x"
multi(0, 1) = 7
multi(1, 0) = null
set dict = server.createObject("scripting.dictionary")
dict.add "some", "x"
dict.add "other", array(new JSON)

data = array(array(1, false), 1.2, null, nothing, multi, dict)
output = (new JSON)("root", data, false)
%>