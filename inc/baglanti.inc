<%
Session.codepage = 65001
Dim conn,Rs
'-------------------------------- MS-ACCESS BAĞLANTISI -------------------------------
Set conn=Server.CreateObject("Adodb.Connection")
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="&Server.Mappath("\db\database.mdb")&";Persist Security Info=False"
Set Rs = Server.CreateObject("Adodb.Recordset")
'----------------------------------------------------------------------------------------
%>