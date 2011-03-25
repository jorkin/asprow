<%
'    This file is part of ASPROW.

'    ASPROW is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    ASPROW is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with ASPROW.  If not, see <http://www.gnu.org/licenses/>.

%>

<html>
	<head>
			<link rel="stylesheet" type="text/css" href="/css/multistyle/<%=request.querystring("custom_style")%>/Style_doctype.css">
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> 
	<script>
	var element_number;
	function ayarla()
	{
		element_number = document.getElementById("cc").value*1;
		var columns = Array(element_number);
		var sira = Array(element_number);
		var counter = 0;
		for(var k=1;k<=element_number;k++)
		{
			try{
				if (document.getElementById(k+"_c").checked)
				{
					if (document.getElementById(k+"_cc").value=="")
						document.getElementById(k+"_cc").value=1;
					sira[counter] = document.getElementById(k+"_cc").value;
					columns[counter] = document.getElementById(k+"_c").name.replace("_display","");
					counter++;
				}
			}catch(e){}
		}
	
		var final_columns = "";
		for(k=0;k<counter;k++)
		{
			min = sira[k];
			cnumber = k;
			for(i=k;i<counter;i++)
			{
				if (sira[i]<min)
				{
					min = sira[i];
					cnumber = i;
				}		
			}
			t = columns[k];
			columns[k] = columns[cnumber];
			columns[cnumber] = t;
			if (k==0)
			final_columns = columns[k];
			else
			final_columns = final_columns+","+columns[k];
		
		}
		document.getElementById("sonuc").value = final_columns
		return true;
	}
	function uncheckall(bool)
	{
		element_number = document.getElementById("cc").value*1;
		for(var sayac = 1;sayac<=element_number+1;sayac++)
		{
			try{
			input = document.getElementById(sayac+"_c");
			input.checked = bool;
			}catch(e){}
		}
	}
	</script>
	</head>
	<body>
<!--#include file="baglanti.inc"-->
<%
	Session.codepage = 1254
	page_name = request("page_name")
	
if request("delete_columns")=1 then
	Response.Cookies(page_name&"_column_order") = ""
	response.write "<h1>Column settings have been reset!<br><a href='../"&replace(request("back_url"),"-","&")&"'>Show</a></h1>"
end if
if request("edit_column")<>"" then
	if request("sonuc")<>"" then
	
		Response.Cookies(page_name&"_column_order") = request("sonuc")
		Response.Cookies(page_name&"_column_order").Expires = date+365
		response.write "Columns are settled <i onclick=""{document.getElementById('sqls').style.display='block'}"">!</i>"
		response.write "<div id='sqls' style='display:none;'><br>Select<br>"&request("sonuc")&"<br> FROM <BR>"&table&"<br><br></div>"
		response.write "<a href='../"&replace(request("back_url"),"-","&")&"'>Show</a>"
	end if
end if


	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	file_path = Server.Mappath(".")&"\.."&"\"&page_name
	changeable = false
	If (fs.FileExists(file_path))=true Then
		set t=fs.OpenTextFile(file_path,1,false)

		do while t.AtEndOfStream=false
			satir = replace(t.ReadLine,"  "," ")
			do while instr(satir,"  ")>0
				satir = replace(satir,"  "," ")
			loop
			satir = replace(satir,"	","")
			if Len(satir)>5 then
				if instrrev(satir,"'")>0 then
					tek = instrrev(satir,"'")
					cift = instrrev(satir,"""")
					if tek>cift then
						satir = Left(satir,tek-1)
					end if
				end if
				tek = instr(satir,"""")
				cift = instrrev(satir,"""")
				if Instr(satir,"CHANGEABLE_COLUMNS")>0 and Instr(satir,"CHANGEABLE_COLUMNS")<tek then
					satir = Mid(satir,tek+1,cift-tek)
					c_columns = replace(satir,"""","")
					if len(c_columns)>0 then
						changeable = true
					end if
				elseif Instr(satir,"COLUMNS")<tek and Instr(satir,"COLUMNS")>0 then
					satir = Mid(satir,tek+1,cift-tek)
					columns = replace(satir,"""","")
				elseif Instr(satir,"TABLENAME")<tek  and Instr(satir,"TABLENAME")>0 then
					satir = Mid(satir,tek,cift-tek)
					table = replace(satir,"""","")
				end if
				
			end if

		loop
	else
		response.write "There are not any page for this criteria!"
		response.end
	end if
	if changeable then
		columns = c_columns
		'response.write columns
	end if
	if table="" or columns="" then
		response.write "This page does not allow you to change column settings!<br /><a href='../"&replace(request("back_url"),"-","&")&"'>Back to Page</a>"
		'response.write table&"- "&columns
		response.end
	end if

	if request.Cookies(page_name&"_column_order")<>"" and (not changeable) then 
		columns = request.Cookies(page_name&"_column_order")
	end if
	form_caption = "<form action='column_editor.asp' method='post' onsubmit='return ayarla();'><table border='1' align='center' width='80%'><tr class='Footer'><td colspan='3' align='center'><a href='column_editor.asp?delete_columns=1&page_name="&page_name&"&back_url="&request("back_url")&"&custom_style="&request("custom_style")&"'>Reset Settings</a><br>Column settings for : <i>"&page_name&"</i></td></tr><tr><td><input id='call_b' type='checkbox' onchange='uncheckall(this.checked)' checked><b>Add to Columns</b></td><td><b>Column Name</b></td><td><b>Show Order</b></td></tr>"
if request("column_editor")=1 then
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Set Rs2 = Server.CreateObject("Adodb.Recordset")
	if Instr(table,".")>0 then
		gelen_tablolar = ""
		bulundu = true
		bitis = Instrrev(table,".")
		baslangic = Instrrev(table," ",bitis)
		do until (bulundu=false)
			if not (bitis>0 and baslangic>0) then
				bulundu = false
			else
				if not instr(gelen_tablolar,Mid(table,baslangic+1,bitis-baslangic))>0 then gelen_tablolar = gelen_tablolar &Mid(table,baslangic,bitis-baslangic)
				bitis = Instrrev(table,".",baslangic)
				if bitis>0 then baslangic = Instrrev(table," ",bitis)	end if
			end if
		loop
		gelen_tablolar = trim(gelen_tablolar)
		g_tablolar = split(gelen_tablolar)
		selected_columns = " "
		counter = 0
		counter2 = 0
		response.write form_caption
		Rs2.open "SELECT TOP 1 "&columns&" FROM "&table,conn
		
		if instr(conn.Provider,"MSDASQL")>0 then
			Rs2.open "SELECT "&columns&" FROM "&table& " LIMIT 0,1 ",conn
		end if
		
		for i=0 to ubound(g_tablolar)
			if not (instr(g_tablolar(i),"_1")>0 or instr(g_tablolar(i),"_2")>0) then
				Rs.open "SELECT top 1 * FROM "&g_tablolar(i),conn 
				if instr(conn.Provider,"MSDASQL")>0 then
					Rs2.open "SELECT "&columns&" FROM "&table& " LIMIT 0,1 ",conn
				end if
				if err=0 then 
					for each x in Rs.fields
						if Instr(selected_columns," "&x.name&" ")=0 then
						counter2 = counter2+1
							selected_columns = selected_columns & " " & x.name
							found = false
							for each y in Rs2.fields	
								if x.name = y.name then
									counter = counter+1
									response.write chr(13)&chr(13)&"<tr class='Altrow'><td><input id='"&counter2&"_c' type='checkbox' name='"&g_tablolar(i)&"."&x.name&"_display' checked>"
									found = true
									exit for
								end if
							next
							if not changeable or (changeable and found=true) then
								if found=false then
									response.write chr(13)&"<tr class='Row'><td><input id='"&counter2&"_c' type='checkbox' name='"&g_tablolar(i)&"."&x.name&"_display'>"
								end if
								response.write "</td><td> <i>"&g_tablolar(i)&"</i> <b>. "&x.name&"</b></td><td><input type='text' id='"&counter2&"_cc' name='"&g_tablolar(i)&"."&x.name&"_no' size='2'"
								if found=true then
									response.write " value='"&counter&"'"
								end if
								response.write "></td></tr>"&chr(13)
							end if
						end if
					next
					Rs.close
				end if
			else
				err = 0
			end if
		next
	else
		sql = "SELECT TOP 1 * FROM "&table
		
		if instr(conn.Provider,"MSDASQL")>0 then
			sql = "SELECT "&columns&" FROM "&table& " LIMIT 0,1 "
		end if
		
		sql2 = "SELECT TOP 1 "&columns&" FROM "&table
		
		if instr(conn.Provider,"MSDASQL")>0 then
			sql2 = "SELECT "&columns&" FROM "&table& " LIMIT 0,1 "
		end if
		
		on error resume next
		Rs.Open sql,conn
		Rs2.Open sql2,conn
		if err<>0 then
			response.write "Hatali parametrelerle karsilasildi: "&err.description&"<br>"
			response.write "SQL1: "&sql&"<br>SQL2:"&sql2&"<br>"
			response.end
		end if
		response.write form_caption
		selected_columns = ""
		counter2 = 0
		for each x in Rs.fields
			if Instr(selected_columns,x.name)=0 then
				selected_columns = selected_columns & " " & x.name
				counter = 0
				found = false
				counter2 = counter2 + 1
				
				for each y in Rs2.fields
					counter = counter+1
					if x.name = y.name then
						response.write chr(13)&chr(13)&"<tr class='AltRow'><td><input id='"&counter2&"_c' type='checkbox' name='"&x.name&"_display' checked>"
						found = true
						exit for
					end if
				next
				if not changeable or (changeable and found=true) then
					if found=false then
						response.write chr(13)&"<tr class='Row'><td><input id='"&counter2&"_c' type='checkbox' name='"&x.name&"_display'>"
					end if
					response.write "</td><td>"&x.name&"</td><td><input  id='"&counter2&"_cc' type='text' name='"&x.name&"_no' size='2'"
					if found=true then
						response.write " value='"&counter&"'"
					end if
					response.write "></td></tr>"&chr(13)
				end if
			end if
		next
	end if
	response.write chr(13)&"<tr align='center'><td colspan='3'><input type='hidden' name='sonuc' id='sonuc'><input type='hidden' id='cc' value='"&counter2&"'><input type='hidden' name='page_name' value='"&page_name&"'><input type='hidden' name='table' value="""&table&"""><input type='hidden' name='columns' value="""&columns&"""><input type='hidden' name='back_url' value="""&request("back_url")&"""><input type='submit' name='edit_column' value='Submit'></td></tr></table></form>"
end if
%>
</body>
</html>