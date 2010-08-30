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
<!--#include file="baglanti.inc"-->
<%
Session.codepage = "65001"
Session.Timeout=120
Dim newid
Dim GERIDON
Dim SIL
Dim KAYDET
Dim GERIAL
Dim sql
Dim fsql
Dim REQUEST_LINKED_COLUMN
Dim BEFORE_TABLE_HTML
Dim GENERICPAGE

Dim RECORD_TITLE 	'Record Basligi
Dim TABLENAME 	'SELECT COLUMNS FROM ______
Dim COLUMNS 	'SELECT _______ FROM TABLENAME
Dim REQUIRED_COLUMNS
Dim LIST_PAGE
Dim RECORD_ID 
Dim LINKED_COLUMN 
Dim THIS_PAGE
Dim UPLOAD_PERMISSIONS
Dim GENERICAJAX_PERMISSION
BEFORE_TABLE_HTML = ""
GERIDON=1
SIL=1
KAYDET=1
GERIAL=1
GERIAL=1
GENERICAJAX_PERMISSION=1


call user("record_head_begin")
public function useAjax(the_id,the_text,table_from,info,style,field)
	Dim html
	Dim RsF
	Dim lookSql
	call user("useAjax_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	lookSql = "SELECT TOP 1 "&the_text
	if info<>"" then lookSql = lookSql & ","&info end if
	lookSql = lookSql & " FROM "&table_from
on error resume next
	if field.value<>"" then
		lookSql = lookSql & " WHERE "&the_id&"="
		if IsNumeric(field.value) then
			lookSql = lookSql &field.value
		else
			lookSql = lookSql &"'"&field.value&"'"
		end if
	end if
	RsF.open lookSql,conn
		html = "<input type=""hidden"" value="""
		if not (REQUEST_LINKED_COLUMN="NEWID") then 
			html = html&field.value
		end if
		html = html&""" name="""&field.name&""" id="""&field.name&""" "&style&" /><input  "&style&" type=""text"" id="""&field.name&"_id"" onchange=""control_ajax(this.id)"""
	if not RsF.eof and not (REQUEST_LINKED_COLUMN="NEWID") then
		html =html& "value="""&RsF(0)&""""
	end if
		html = html &  "style='width:100%;' />"
	RsF.close
	inscript =  "var options_xml_"&field.name&" = {"&chr(13)
	inscript =	inscript &"script: function (input) { return """&THIS_PAGE&"?ID="&the_id&"&VIEW="&the_text&"&KEYWORD=""+document.getElementById('"&field.name&"_id').value+""&INFO="&info&"&TABLE="&table_from&"&AJAX=1&""; },"&chr(13)
	inscript =  inscript &"varname:""input""," & chr(13)
	inscript =  inscript &"callback: function (obj) { document.getElementById('"&field.name&"').value = obj.id; }," & chr(13)
	inscript =  inscript &"json: true"& chr(13)
	inscript =  inscript &" };" & chr(13)& chr(13)
	inscript =  inscript &"var as_xml_"&field.name&" = new bsn.AutoSuggest('"&field.name&"_id', options_xml_"&field.name&");" & chr(13)
	JAVASCRIPT = JAVASCRIPT&chr(13)&inscript
	call user("useAjax_end")
	useAjax = html
end function
	
public function showAjax(the_id,the_text,the_info,table_from,keyword)
'on error resume next
	Dim RsF
	Dim sql_info
	call user("showAjax_begin")
	the_text = ""&the_text
	Set RsF = Server.CreateObject("Adodb.Recordset")
	if the_info<>"" then
		sql_info = ","&the_info
	end if
	if keyword<>"*" then
		sql = "SELECT TOP 25 "&the_text&","&the_id&sql_info&" FROM "&table_from& " WHERE "&the_text&" LIKE '"&replace(replace(keyword,"'",""),"""","")&"%'"
	else
		sql = "SELECT TOP 25 "&the_text&","&the_id&sql_info&" FROM "&table_from
	end if
	RsF.open "SELECT TOP 1 "&the_text&" FROM "&table_from,conn
	if not (RsF.fields(0).type=201 or RsF.fields(0).type=203 or RsF.fields(0).type=205) then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		sql = sql&" ORDER BY "&the_text&" asc"
	end if
	RsF.close
	
	RsF.open sql,conn
	results =  "{ results: [ " & vbCrLf
	sayac = 0
		do until RsF.eof
			if the_info<>"" then
				results = results & "	{ id: """&RsF(1)&""", value: """&RsF(0)&""", info: """&RsF(2)&""" },"&vbCrLf
			else
				results = results & "	{ id: """&RsF(1)&""", value: """&RsF(0)&""", info: """" },"&vbCrLf
			end if
			RsF.Movenext
			sayac = sayac + 1
		loop
			results = left(results,Len(results)-3)&vbCrLf
		results = results&"] }"
	RsF.close
	call user("showAjax_end")
	showAjax = results
end function

public function save_new_record
	Dim RsF
	call user("save_new_record_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	RsF.open "SELECT top 1 "&COLUMNS&" FROM "&TABLENAME,conn
	
	sql = "INSERT INTO "&TABLENAME&" ("
	for each x in RsF.Fields
		if not uCASE(x.name)=uCASE(RECORD_ID) then
			if Instr(Request.Form,"&"&x.name&"=")<>"" or Request.Form(1)=x.name then
				if Request.Form(x.name)="" then
					sql = sql &x.name&","
				else
					if (x.type>1 and x.type<6) or (x.type>15 and x.type<22) then 'Bunlar integer
						sql = sql & x.name&","
					else
						sql = sql &x.name&","
					end if
				end if
			end if
		else
			if not x.Attributes=adParamNullable  and x.Attributes=adFldRowID then
				sql = sql&uCASE(x.name)&","
			end if
		end if
	next
	if Right(sql,1)="," then
		sql = Left(sql,Len(sql)-1)&") VALUES ("
	else
		response.write "En az 1 bilgiyi doldurmalisiniz.."
	end if
	'Olasi yeni id
	Set RsT = Server.CreateObject("Adodb.Recordset")
	RsT.Open "SELECT TOP 1 "&RECORD_ID&" FROM "&TABLENAME&" ORDER BY "&RECORD_ID&" DESC",conn
	if not RsT.eof then 
		newid = (RsT(RECORD_ID)+1)
	else
		newid = 1
	end if
	RsT.Close
	'DEVAM
	for each x in RsF.Fields
		if not uCASE(x.name)=uCASE(RECORD_ID) then
			if Instr(Request.Form,"&"&x.name&"=")<>"" or Request.Form(1)=x.name then
				if Request.Form(x.name)="" then
					sql = sql &"NULL,"
				else
					if (x.type>1 and x.type<6) or (x.type>15 and x.type<22) then 'Bunlar integer
						sql = sql & replace(Request.Form(x.name),"'","''")&","
					else
						sql = sql & "'"&replace(Request.Form(x.name),"'","''")&"',"
					end if
				end if
			end if
		else
			if not x.Attributes=adParamNullable  and x.Attributes=adFldRowID then
				sql = sql&newid&","
			end if
		end if
	next
	if Right(sql,1)="," then
		sql = Left(sql,Len(sql)-1)&")"
	else
		response.write "En az 1 bilgiyi doldurmalisiniz.."
	end if
	call user("before_execute_sql_save_record_new")
	on error resume next
	conn.Execute sql
	if err<>0 then
		response.write "Hata: "&err.description&"<BR>"&sql
		response.end
	else
		call user("save_new_record_end")
	end if
	
end function

public function getValue(the_text,table_from,filter)
	Dim RsF
	if not Instr(filter,"NEWID")>0 THEN
	Set RsF = Server.CreateObject("Adodb.Recordset")
	fsql = "SELECT "&the_text&" FROM "&table_from
	if filter<>"" then
		fsql = fsql& " WHERE "&filter
	end if
	on error resume next
	RsF.open fsql,conn
	if err<>0 then response.write err.description&"<br>"&fsql
	if not RsF.eof then
		getValue = RsF(0)
	else
		getValue = ""
	end if
	else
		getValue = ""
	end if
end function

public function save_record
	Dim RsF
	call user("save_record_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	RsF.open "SELECT top 1 "&COLUMNS&" FROM "&TABLENAME,conn
	sql = "UPDATE "&TABLENAME&" SET "
	for each x in RsF.Fields
		if not x.name=RECORD_ID then
			if Instr(Request.Form,"&"&x.name&"=")>0 or Instr(Request.Form,x.name&"=")=1 then
				sql = sql&" "&x.name&"="
				if Request.Form(x.name)="" then
					sql = sql &"NULL, "
				else
					if (x.type>1 and x.type<6) or (x.type>15 and x.type<22) then 'Bunlar integer
						sql = sql & replace(Request.Form(x.name),"'","''")&", "
					else
						sql = sql & "'"&replace(Request.Form(x.name),"'","''")&"', "
					end if
				end if
			end if
		end if
	next
	sql = Left(sql,Len(sql)-2)
	sql = sql&" WHERE "&RECORD_ID&"="
	if IsNumeric(Request.Form(RECORD_ID)) then
		sql = sql & Request.Form(UCASE(RECORD_ID))
	else
		sql = sql & "'"&Request.Form(RECORD_ID)&"'"
	end if
	RsF.close
	call user("before_execute_sql_save_record")
	on error resume next
	conn.execute sql
	if err<>0 then
		response.write err.description&"<br>"&sql
		response.end
		save_record = 0
	else
		call user("save_record_end")
		save_record = 1
	end if
end function

public function delete_record
	call user("delete_record_begin")
	if request(RECORD_ID)<>"" then
		sql = "DELETE FROM "&TABLENAME &" WHERE "&RECORD_ID&"="
		if IsNumeric(request(RECORD_ID)) then
			sql = sql&request(RECORD_ID)
		else
			sql = sql&"'"&request(RECORD_ID)&"'"
		end if
		conn.execute sql
	end if
	call user("delete_record_end")
end function

public function createGLookup(the_id,the_text,table_from,filter,style,field,connstr)
	Dim RsF,ConnF
	call user("createLookup_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	Set ConnF=Server.CreateObject("Adodb.Connection")
	ConnF.Open connstr

	fsql = "SELECT "&the_text&","&the_id&" FROM "&table_from
	if filter<>"" then
		fsql = fsql& " WHERE "&filter
	end if
	RsF.open "SELECT TOP 1 "&the_text&" FROM "&table_from,ConnF
	if not (RsF.fields(0).type=201 or RsF.fields(0).type=203 or RsF.fields(0).type=205) then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		fsql = fsql&" ORDER BY "&the_text&" asc"
	end if
	RsF.close
	RsF.open fsql,ConnF
	selection = 0
	do until RsF.eof
		result = result & "<option value='"&RsF(1)&"'"
		if not (REQUEST_LINKED_COLUMN="NEWID") then
			if RsF(1)=field.value  then
				selection = 1
				result= result&" selected"
			end if
		end if
		result = result & ">"&RsF(0)&"</option>"
		RsF.Movenext
	loop
	if selection=0 and filter="" then
		result = "<option value=''>Seciniz</option>"&result
	end if
	result = "<select name='"&field.name&"'  id='"&field.name&"_id' "&style&">"&result&"</select>"
	call user("createLookup_end")
	ConnF.close
	createGLookup = result
end function

public function createLookup(the_id,the_text,table_from,filter,style,field)
	Dim RsF
	call user("createLookup_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	fsql = "SELECT "&the_text&","&the_id&" FROM "&table_from
	if filter<>"" then
		fsql = fsql& " WHERE "&filter
	end if
	RsF.open "SELECT TOP 1 "&the_text&" FROM "&table_from,conn
	if not (RsF.fields(0).type=201 or RsF.fields(0).type=203 or RsF.fields(0).type=205) then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		fsql = fsql&" ORDER BY "&the_text&" asc"
	end if
	RsF.close
	RsF.open fsql,conn
	selection = 0
	do until RsF.eof
		result = result & "<option value='"&RsF(1)&"'"
		if IsNumeric(REQUEST_LINKED_COLUMN) then
			if RsF(1)=field.value  then
				selection = 1
				result= result&" selected"
			end if
		end if
		result = result & ">"&RsF(0)&"</option>"
		RsF.Movenext
	loop
	if selection=0 then
		result = "<option value=''>Seciniz</option>"&result
	end if
	result = "<select name='"&field.name&"'  id='"&field.name&"_id' "&style&">"&result&"</select>"
	call user("createLookup_end")
	createLookup = result
end function

'Added by Orhun Alp Oral, at 12/June/2010
public function createGenericLookup(the_id,the_text,table_from,filter,style,field)
	createGenericLookup = createLookup(the_id,the_text,table_from,filter,style,field)&"| <a href=""#""><img onclick=""RefreshCombo('"&field.name&"_id','"&the_id&"','"&the_text&"','"&table_from&"','"&filter&"')"" border='0' src='css/images/refresh.gif' alt='Refresh ComboBox'></a> <a href=""genericList.asp?gnrTABLENAME="&table_from&""" target=""_blank""><img border='0' src='css/images/edit.gif' alt='Edit ComboBox'></a>"
end function

public function genericAjax(the_id,the_text,table_from,filter)
	Dim RsF
	dim fsql
	call user("genericAjax_begin")
	response.clear
	Set RsF = Server.CreateObject("Adodb.Recordset")
	if GENERICAJAX_PERMISSION=1 then
		fsql = "SELECT "&the_text&","&the_id&" FROM "&table_from&" ORDER BY "&the_text
		if filter<>"" then
			fsql = fsql&" WHERE "&filter
		end if
		RsF.open fsql,conn
		jsonArray = "{""text"" : ""Seciniz"", ""value"" : """"}"
		if err<>0 then
			jsonArray = "{""text"" : ""SQL Error"", ""value"" : """"}"
		else
			do until RsF.eof
				jsonArray = jsonArray & ", {""text"" : """&RsF(0)&""", ""value"" : """&RsF(1)&"""}"
				RsF.Movenext
			loop
		end if
		genericAjax = "{ ""option"": ["&jsonArray&"]}"
	else
		genericAjax = " { ""option"": [ {""text"" : ""INVALID PERMISSION "", ""value"" : """" }]} "
	end if
end function

public function PrintTitle(field)
	Dim start,finish,zorunlu
	start = Instr(CASES,"["&field.name&"=")
	if instr(ucase(REQUIRED_COLUMNS),ucase(field.name))>0 then
		zorunlu = " <font color='red'>*</font>"
	else
		zorunlu = ""
	end if
	
	if start>0 then
		start = start + Len("["&field.name&"=")
		finish = Instr(start,CASES,"|")
		PrintTitle = (Mid(CASES,start,finish-start)&zorunlu)
	else
		PrintTitle = (CStr(UCase(Left(field.name,1))&Lcase(Right(field.name,Len(field.name)-1)))&zorunlu)
	end if
end function

public function PrintField(field)
	call user("PrintField_begin")
	print = ""
	field_style = getFieldFor(field)
	if field_style<>"" then
		if Left(field_style,1)="<" then
			print = field_style
		end if
	end if
	if print = "" then
			if field.type=7 or  field.type=133  or  field.type=135 then '133=adDBDate , 7 = adDate , 135=adDBTimeStamp
				print = "<script>DateInput('"&field.name&"', true, 'YYYY-MM-DD'"
				'style
				if field_style<>"" then 
					print = print &",'"&replace(field_style,"'","\'")&"'"
				else
					print = print &",''"
				end if
				' dd/mm/yyyy den mm/dd/yyyy ye cevirdik.
				on error resume next
				if field.value<>"" then
					ay = Month(field.value)
					gun = Day(field.value)
					yil = Year(field.value)
				end if
				if (REQUEST_LINKED_COLUMN="NEWID") then
					ay = Month(Date)
					gun = Day(Date)
					yil = Year(Date)
				end if
				if (REQUEST_LINKED_COLUMN="NEWID") or field.value<>"" then
					if Len(ay)=1 then ay="0"&ay end if
					if Len(gun)=1 then gun="0"&gun end if
					print = print &",'"&yil&"-"&ay&"-"&gun&"'"
				end if
				
				print = print & ")</script>"
			elseif field.type=201 or field.type=203 or field.type=205 then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
				if (Instr(field_style,"disabled")>0 or Instr(field_style,"readonly")>0) and HTMLEDITOR=1 then
					print = "<div style=""overflow:auto;border: 1px solid #666;height:250px;width:90%; padding:6px;"">"
					print = print&field.value
					print = print&"</div>"
				else
					print = "<textarea style='width:100%;height:250px;' id='"&field.name&"_id' "&field_style&""
					if not instr(field_style,"readonly")>0 then print = print & " name='"&field.name&"'" end if
					print = print & ">"
					if not (REQUEST_LINKED_COLUMN="NEWID") then
						print = print&field.value
					end if
					print = print &"</textarea>"&chr(13)
					if HTMLEDITOR=1 then
						print = print &"<script type=""text/javascript"">"&chr(13)
						print = print &"	CKEDITOR.replace( '"&field.name&"' );"&chr(13)
						print = print &"</script>"&chr(13)
					end if
				end if
			elseif field.type=11 then '11 = boolean
				print = "<input type='checkbox' name='c-"&field.name&"'  id='"&field.name&"_id' onclick='this.form."&field.name&".value=1-this.form."&field.name&".value;' "&field_style&""
				on error resume next
				if not (field.value="" or Isnull(field.value)) then
					if (field.value="True" or field.value=true or field.value=1) and not (REQUEST_LINKED_COLUMN="NEWID") then 
						print = print & " checked"
					end if
				end if
				print = print & "><input type='hidden' value='"
				if (field.value="True" or field.value=true or field.value=1) and not (REQUEST_LINKED_COLUMN="NEWID") then 
					print = print & "1"
				else
					print = print & "0"
				end if
				print = print & "' name='"&field.name&"'>"
			elseif field.type=3 then '3 = integer
			
				if field_style<>"" then
					if Left(field_style,1)="<" then
						print = field_style
					end if
				end if
				if print="" then
					print = "<input name='"&field.name&"' type='text' size='10'  id='"&field.name&"_id' value='"
					if not (REQUEST_LINKED_COLUMN="NEWID") then
						print = print&field.value
					end if	
					print = print&"' "&field_style&" >"
				end if
			else
				print = "<input name='"&field.name&"' type='text' size='50'  id='"&field.name&"_id' value='"
					if not (REQUEST_LINKED_COLUMN="NEWID") then
						print = print&field.value
					end if	
					print = print&"' "&field_style&" >"
			end if
	end if
	call user("PrintField_end")
	PrintField = print
end function

public function makeUploader(field,specialPermissions)
	dim fname,fvalue
	if UPLOAD_PERMISSIONS="" then
		UPLOAD_PERMISSIONS = "|UP|DOWN|DEL|"
	end if
	if specialPermissions<>"" then
		UPLOAD_PERMISSIONS = specialPermissions
	end if
	Session(replace(THIS_PAGE,".asp","")&"_"&field.name&"_"&request(LINKED_COLUMN)) = UPLOAD_PERMISSIONS
	on error resume next
	fvalue = field.value
	fname = field.name
	mup = "<div><input name='"&fname&"' type='text' size='30'  id='"&fname&"_id' value='"&fvalue&"'></div>"
	makeUploader = mup&"<iframe src=""inc/upload.asp?fname="&fname&"&ext="&fvalue&"&rid="&request(LINKED_COLUMN)&"&pagefrom="&replace(THIS_PAGE,".asp","")&""" style=""width:400px;height:40px;margin:0px;padding:0px;"" scrolling=""no"" frameborder=""no""></iframe>"
end function

public function getFieldFor(field)
	Dim innerTxt
	call user("getFieldFor_begin")
	start = InStr(CASES,"|"&field.name&"|")+Len(field.name)+2
	if start>(Len(field.name)+2) then
		finish = InStr(start,CASES,"|")
		innerTxt = Mid(CASES,start,finish-start)
		if (Instr(innerTxt,"makeUploader") or Instr(innerTxt,"useAjax") or Instr(innerTxt,"createLookup") or Instr(innerTxt,"createGLookup") or Instr(innerTxt,"createGenericLookup")) then
			
			getFieldFor = Eval(""&innerTxt&"")
		else
			getFieldFor = innerTxt
		end if
	else
		call user("getFieldFor_end")
		getFieldFor = ""
	end if
end function

Function getFileName()
  Dim lsPath, arPath
  lsPath = Request.ServerVariables("SCRIPT_NAME")
  arPath = Split(lsPath, "/")
  GetFileName =arPath(UBound(arPath,1))
End Function

Function getID(table)
	getID = ""
	Set adox = CreateObject("ADOX.Catalog")
	adox.activeConnection = conn
	for each table in adox.tables
		if ucase(table.name)=ucase(TABLENAME) then
			For Each idxADOX In table.Indexes
				if idxADOX.PrimaryKey then 
					getID =  idxADOX.columns(0).name 
					exit for
				end if
			next
		exit for
		end if
	next
end Function
%>