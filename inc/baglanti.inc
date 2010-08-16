<%
Session.codepage = 65001
Dim conn,Rs
'-------------------------------- SQL SERVER BAGLANTISI -------------------------------
Set conn=Server.CreateObject("Adodb.Connection")
Conn.Open "Provider=SQLOLEDB;Password=01012000;User ID=sa;Initial Catalog=MOBASIST;Data Source=HP-PC\SQLEXPRESS;"
Set Rs = Server.CreateObject("Adodb.Recordset")
'----------------------------------------------------------------------------------------
%>