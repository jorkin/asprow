<%
Session.codepage = 65001
Dim conn,Rs,gconnstr

'-------------------------------- CONNECTION STRINGS -------------------------------
//Access Connection
gconnstr = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="&Server.Mappath("\db\database.mdb")&";Persist Security Info=False"
//MySQL 3.51 Connection
'gconnstr = "Provider={MySQL ODBC 3.51 Driver}; UID=root;PWD=;SERVER=127.0.0.1;DATABASE=cancer;"
//MySQL 5.1 Connection
'gconnstr = "DRIVER={MySQL ODBC 5.1 Driver}; UID=root;PWD=;SERVER=127.0.0.1;DATABASE=cancer;"


//MS-SQL Connection
'gconnstr = "Provider=SQLOLEDB.1;Password=01012000;Persist Security Info=True;User ID=SA;Initial Catalog=MEDISOFT_NET;Data Source=127.0.0.1"
'----------------------------------------------------------------------------------------
'on error resume next
Set conn=Server.CreateObject("Adodb.Connection")
Conn.Open gconnstr
Set Rs = Server.CreateObject("Adodb.Recordset")
'if err<>0 then
'response.write err.description
'end if
%>