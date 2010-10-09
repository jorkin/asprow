<!--#include file="inc/grid_head.inc"-->
<%
sub user(myFunction)
	'myFunction: %function_name%_begin , %function_name%_end
End sub
'Sabitlerin Ayarlanmasi
LIST_TITLE = "ÇÖŞİĞÜı"	'Listenin Basligi
TABLENAME = "ASPROW"	'SELECT COLUMNS FROM ______
COLUMNS = "*"	'SELECT _______ FROM TABLENAME
RECORD_PAGE = "asprow_record.asp"	'Linki verilecek sayfa
CHANGEABLE_COLUMNS = "*"	'Kolon ayarlarinda cikacak fieldlari yazin f1,f2
LINKED_COLUMN = "ASPROWID"	'Linki verilecek sutun
%>
<!--#include file="inc/grid_body.inc"-->
<table align='center'><!--#include file="menu.html"--></table>
