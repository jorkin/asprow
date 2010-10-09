<!--#include file="inc/record_head.inc"-->
<%
'User Sub
sub user(myFunction)
	'myFunction: %function_name%_begin , %function_name%_end
End sub

'Sabitlerin Ayarlanmasi
CASES = "|"
REQUIRED_COLUMNS= "FIELD1 FIELD2 FIELD3"	'Doldurulmasi zorunlu fieldlar
RECORD_TITLE = "asd"	'Record Basligi
TABLENAME = "ASPROW"	'SELECT COLUMNS FROM ______
COLUMNS = "PVERSION, DESCRIPTION"	'SELECT _______ FROM TABLENAME
LIST_PAGE = "asprow_grid.asp"	'Liste Sayfasi
LINKED_COLUMN = "ASPROWID"	'Kayitlar hangi kolon adindan yollaniyor
%>
<!--#include file="inc/record_body.inc"-->
<table align='center'><!--#include file="menu.html"--></table>
