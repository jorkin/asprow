<!--#include file="record_head.inc"-->
<%
	
	sub user(myFunction)
		'myFunction: %function_name%_begin , %function_name%_end
	End sub
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	file_path = Server.Mappath(".")&"\.."&"\"&request("PAGENAME")
	If (fs.FileExists(file_path))=true Then
		set t=fs.OpenTextFile(file_path,1,false)
		varexecute = false
		do while t.AtEndOfStream=false
			satir = replace(t.ReadLine,"  "," ")
			if instr(satir,"RECORD_TITLE")>0 or instr(satir,"TABLENAME")>0 or instr(satir,"COLUMNS")>0 or instr(satir,"LIST_PAGE")>0 or instr(satir,"LINKED_COLUMN")>0 or instr(satir,"REQUIRED_COLUMNS")>0 or instr(satir,"CASES")>0 then
				'response.write satir
				Execute(satir) 
			end if
		loop
		RECORD_ID = getID(TABLENAME)
		if RECORD_ID="" then RECORD_ID = LINKED_COLUMN end if
		if LINKED_COLUMN="" then LINKED_COLUMN = RECORD_ID end if
		if RECORD_ID<>"" then
			if not (Instr(replace(" "&COLUMNS&" ",","," ")," "&RECORD_ID&" ")>0 or COLUMNS="*") then
				EXTRA_COLUMN=","&RECORD_ID
			else
				EXTRA_COLUMN=""
			end if
		end if
		
		on error resume next
		sql = "SELECT top 1 "&COLUMNS&EXTRA_COLUMN&" FROM "&TABLENAME&" "&where_statement
		
		if instr(conn.Provider,"MSDASQL")>0 then 
			sql = "SELECT "&COLUMNS&EXTRA_COLUMN&" FROM "&TABLENAME&" "&where_statement& " LIMIT 0,1 " 
		end if
		
		Set Rs = Server.CreateObject("Adodb.Recordset")
		rs.open sql,conn
		if err<>0 then
			response.write "Error 1:"&err.description&"<hr> SQL: "&sql&"<hr>"
			response.write "<hr>Columns : "&COLUMNS&"<hr>"&EXTRA_COLUMN&"<hr>"
			response.write "Rs.eof : "&rs.eof
			response.end
		end if
	else
		response.write "File couldn't find.."
		response.end
	end if
%>
<form action="<%=request("PAGENAME")%>" method="post" onsubmit="return form_control()">
	<table cellspacing="0" cellpadding="0" border="0" align="center" width="700">
		<tr>
		  <td valign="top">
			<table cellspacing="0" cellpadding="0" border="0" class="Header">
			  <tr>
				<td class="HeaderLeft">&nbsp;</td> 
				<td class="th"><strong><%=RECORD_TITLE%></strong></td> 
				<td class="HeaderRight">&nbsp;</td> 
			  </tr>
			</table>
			<table class="Record" cellspacing="0" cellpadding="0">
		<%
			for each x in rs.Fields
				if LCase(x.name) = LCase(RECORD_ID) then
					if not (EXTRA_COLUMN<>"") then ' Secilmeyen IDnin yazilmasi sevimsiz olur
					%>
					<tr class="Controls"><td class="th"><label for="<%=x.name%>"><%=PrintTitle(x)%></label></td><td><input type="text" size="15" id="<%=x.name%>_id"  name="<%=x.name%>" value="{<%=x.name%>}" readonly></td></tr>
					<%
					else
					%>
					<input type="hidden" name="<%=x.name%>" value="NEWID" id="<%=x.name%>_id">	
					<%
					end if
				else
			%>
				<tr class="Controls"><td class="th"><label for="<%=x.name%>"><%=PrintTitle(x)%></label></td><td>
				<%
				response.write PrintField(x)
				%>
				</td></tr>
			<%
				end if
			next
		%>
		<tr class="Bottom"><td colspan="2"  align="right">
		<input type="checkbox" id="cogaltbutton" onclick="cogalt()"> Çoğalt
		<input name="kaydet" type="submit" class="Button" value="Kaydet">
		<input type="reset" class="Button" value="Tümünü Geri Al">
		<input type="submit" name="delete" class="Button" value="Sil">
		<input type="hidden" name="back_url" value="<%=replace(request("back_url"),"-","&amp;")%>"><input type="submit" name="back" class="Button" value="Geri Dön">
		</td></tr>
	</table>
</table>
</form>