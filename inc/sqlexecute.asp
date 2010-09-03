<!--#include file="baglanti.inc"-->
<%
dim sql
Session.CodePage = 65001
sql=trim(request.form("sql"))
%>
<script src="/JS/createtable.js"></script>
<script>tableTemplate()</script>
<script>
	function writeMe(element)
	{
	document.getElementById("sql_text").value = "SELECT * FROM "+element.innerHTML+"\n"+document.getElementById("sql_text").value;
	}
</script>
<hr>
<%
Set adox = CreateObject("ADOX.Catalog")
adox.activeConnection = conn
response.write "Active Tables: "
for each table in adox.tables

	For Each idxADOX In table.Indexes
		if idxADOX.PrimaryKey then
		pk_field =  idxADOX.columns(0).name
		end if
	next
	
	if table.type="TABLE" then  
		response.write "<b style='cursor:hand;cursor:pointer;'><u><span id='"&table.name&"_btbl' onclick='writeMe(this)'>"&table.name & "</span></u></b> ("  
		set rs = conn.execute("SELECT COUNT(*) FROM " & table.name) 
		response.write rs(0) &","& pk_field&"), "
	end if 
next 
response.write "<br /><br />Active Views: "
for each table in adox.tables  
	if table.type="VIEW" then  
		response.write "<b style='cursor:hand;cursor:pointer;'><u><span id='"&table.name&"_btbl' onclick='writeMe(this)'>"&table.name & "</span></u></b> ("  
		set rs = conn.execute("SELECT COUNT(*) FROM " & table.name) 
		response.write rs(0) & "), "
	end if 
next 
%>
<hr>
<form action="sqlexecute.asp" method="post">
<table align=center border=0>
	<tr align="center"><td>
	<textarea style="width:600px;height:200px;" id="sql_text" name="sql" size="40"><% if request("sql")="" then response.write "SELECT * FROM WEBPAGES" else response.write request("sql")%></textarea><br>
	<input style="width:600px;height:20px;" type="submit" value="Gönder">
	</td></tr>
</table>
</form>
<br>
<%
Set Rs = Server.CreateObject("Adodb.Recordset")
if sql="" then
	response.end
end if
	on error resume next
	set rs = conn.Execute(sql)
if err<>0 then
%>
	<b>Database Hatası! , Kayıt Güncellenemedi!:</b>
	<br>Sebep : "<%=err.description%>"<br>
	<textarea cols = '20' rows = '5' disabled><%=sql%></textarea>
<%
else 
	response.write("<b>Kod Basarili</b>")
	
	if rs.Fields.Count>0 then
	%>
	<table border="1" width="100%" bgcolor="#fff5ee">
	<tr>
	<%for each x in rs.Fields
		response.write("<th align='left' bgcolor='#b0c4de'>" & x.name & "</th>")
	next%>
	</tr>
	<%do until rs.EOF%>
		<tr>
		<%for each x in rs.Fields%>
		   <td><%Response.Write(x.value)%></td>
		<%next
		rs.MoveNext%>
		</tr>
	<%loop
	rs.close
	conn.close
	end if
	%>
	</table>
	<%
end if
%>