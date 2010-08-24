<!--#include file="inc\baglanti.inc"-->
<%
dim sql
Session.CodePage = 65001
sql=trim(request.form("sql"))
%>
<script src="createtable.js"></script>
<script>tableTemplate()</script>
<hr>
<form action="sqlexecute.asp" method="post">
<table align=center border=0>
	<tr align="center"><td>
	<textarea style="width:600px;height:200px;" name="sql" size="40"><% if request("sql")="" then response.write "SELECT * FROM WEBPAGES" else response.write request("sql")%></textarea><br>
	<input style="width:600px;height:20px;" type="submit" value="Gönder">
	</td></tr>
</table>
</form>
<br>
<%
if sql="" then
	response.end
end if
	on error resume next
if instr(sql,"INSERT")>0 OR instr(sql,"DELETE")>0 OR instr(sql,"UPDATE")>0 OR instr(sql,"ALTER")>0 then
	conn.Execute sql
else
	rs.Open sql, conn
end if
if err<>0 then
%>
	<b>Database Hatası! , Kayıt Güncellenemedi!:</b>
	<br>Sebep : "<%=err.description%>"<br>
	<textarea cols = '20' rows = '5' disabled><%=sql%></textarea>
<%
else 
	response.write("<b>Kod Basarili</b>")
	if not (instr(sql,"INSERT")>0 OR instr(sql,"DELETE")>0 OR instr(sql,"UPDATE")>0 OR instr(sql,"ALTER")>0 OR instr(sql,"CREATE")>0 OR instr(sql,"DROP")>0) then
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
	%>
	</table>
	<%
	end if
end if
%>