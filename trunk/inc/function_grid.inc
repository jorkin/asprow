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
<%
public function updateFields()
	Dim updated_records,RsF
	REC_IDS = Request.Form("REC_IDS")
	on error resume next
	if REC_IDS<>"" then
		updated_records = Split(REC_IDS)
		Set RsF = Server.CreateObject("Adodb.Recordset")
		RsF.open "SELECT TOP 1 "&COLUMNS&" FROM "&TABLENAME,conn
		for i=0 to ubound(updated_records)
			sql = "UPDATE "&TABLENAME&" SET "
			for each x in RsF.fields
				if not lcase(x.name)=lcase(LINKED_COLUMN) then
					if instr(Request.Form,"r_"&updated_records(i)&"_"&x.name)>0 or instr(Request.Form,"r_"&updated_records(i)&"_"&x.name)>0 then
						if (x.type>1 and x.type<6) or (x.type>15 and x.type<22) then 'Bunlar integer
							sql = sql&", "&x.name&" = "&int(Request.Form("r_"&updated_records(i)&"_"&x.name))
						else
							if Request.Form("r_"&updated_records(i)&"_"&x.name)="" then
								sql = sql&", "&x.name&" = NULL"
							else
								sql = sql&", "&x.name&" = '"&replace(Request.Form("r_"&updated_records(i)&"_"&x.name),"'","''")&"'"
							end if
						end if
					end if
				end if
			next
			sql = replace(sql,"UPDATE "&TABLENAME&" SET ,","UPDATE "&TABLENAME&" SET ")
			sql = sql &" WHERE "&LINKED_COLUMN&" = "&updated_records(i)
			on error resume next
			conn.execute sql
			response.write sql&"<br>"
			if err<>0 then
				response.write "Hatayla karsilasildi: "&err.description&"<br>"&sql
				response.end
			end if
		next
	end if
end function

public function PrintField(field,field_id)
	call user("PrintField_begin")
	field_style = "style = 'width:100%;height100%;'"
	if field.type=7 or  field.type=133  or  field.type=135 then '133=adDBDate , 7 = adDate , 135=adDBTimeStamp
		print = "<script>DateInput('r_"&field_id&"_"&field.name&"', true, 'MM/DD/YYYY'"
		'style
		if getFieldFor(field,0)<>"" then 
			print = print &",'"&replace(getFieldFor(field,0),"'","\'")&"'"
		else
			print = print &",''"
		end if
		' dd/mm/yyyy den mm/dd/yyyy ye cevirdik.
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
			print = print &",'"&ay&"/"&gun&"/"&yil&"'"
		end if
		
		print = print & ")</script>"
	elseif field.type=201 or field.type=203 or field.type=205 then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		if (Instr(field_style,"disabled")>0 or Instr(field_style,"readonly")>0) and HTMLEDITOR=1 then
			print = "<div style=""overflow:auto;border: 1px solid #666;height:250px;width:90%; padding:6px;"">"
			print = print&field.value
			print = print&"</div>"
		else
			print = "<textarea id='r_"&field_id&"_"&field.name&"_id' "&field_style&""
			if not instr(field_style,"readonly")>0 then print = print & " name='r_"&field_id&"_"&field.name&"'" end if
			print = print & ">"
			if not (REQUEST_LINKED_COLUMN="NEWID") then
				print = print&field.value
			end if
			print = print &"</textarea>"&chr(13)
			if HTMLEDITOR=1 then
				print = print &"<script type=""text/javascript"">"&chr(13)
				print = print &"	CKEDITOR.replace( '"&"r_"&field_id&"_"&field.name&"' );"&chr(13)
				print = print &"</script>"&chr(13)
			end if
		end if
	elseif field.type=11 then '11 = boolean
		print = "<input type='checkbox' name='c-"&field.name&"'  id='r_"&field_id&"_"&field.name&"_id' onclick='this.form.r_"&field_id&"_"&field.name&".value=1-this.form.r_"&field_id&"_"&field.name&".value;' "&field_style&""
		if field.value=true and not (REQUEST_LINKED_COLUMN="NEWID") then 
			print = print & " checked"
		end if
		print = print & "><input type='hidden' value='"
		if field.value=true and not (REQUEST_LINKED_COLUMN="NEWID") then 
			print = print & "1"
		else
			print = print & "0"
		end if
		print = print & "' name='r_"&field_id&"_"&field.name&"'>"
	elseif field.type=3 then '3 = integer
		print = "<input name='r_"&field_id&"_"&field.name&"' type='text' id='r_"&field_id&"_"&field.name&"_id' value='"
		if not (REQUEST_LINKED_COLUMN="NEWID") then
			print = print&field.value
		end if	
		print = print&"' "&field_style&" >"
		
	else
		print = "<input name='r_"&field_id&"_"&field.name&"' type='text' id='r_"&field_id&"_"&field.name&"_id' value='"
			if not (REQUEST_LINKED_COLUMN="NEWID") then
				print = print&field.value
			end if	
			print = print&"' "&field_style&" >"
	end if
	call user("PrintField_end")
	PrintField = print
end function

public function getValue(the_text,table_from,filter)
	Dim RsF
	if not Instr(filter,"NEWID")>0 THEN
	Set RsF = Server.CreateObject("Adodb.Recordset")
	sql = "SELECT "&the_text&" FROM "&table_from
	if filter<>"" then
		sql = sql& " WHERE "&filter
	end if
	RsF.open sql,conn
	if not RsF.eof then
		getValue = RsF(0)
	else
		getValue = ""
	end if
	else
		getValue = ""
	end if
end function

public function getURLGNR
	dim show_requests,url
	url = ""
	show_requests = Split(Request.Querystring,"&")
	for i = 0 to ubound(show_requests)
		if Left(lcase(show_requests(i)),Len("gnr"))="gnr" then
			url = url &"&"& show_requests(i)
		end if
	next
	getURLGNR = url
end function

public function getURL
	dim show_requests,url
	url = ""
	show_requests = Split(Request.Querystring,"&")
	for i = 0 to ubound(show_requests)
		if Left(lcase(show_requests(i)),Len("q_"))="q_" then
			url = url &"&"& show_requests(i)
		end if
	next
	if request.querystring("custom_style")<>"" then 
		url = url & "&custom_style="&request.querystring("custom_style")
	end if
	getURL = url&getShowURL&getURLGNR
end function

public function getShowURL
	dim show_requests,shows
	shows = ""
	show_requests = Split(Request.Querystring,"&")
	for i = 0 to ubound(show_requests)
		if Left(ucase(show_requests(i)),Len("SHOW_"))="SHOW_" then
			shows = shows & "&" & show_requests(i)
		end if
	next
	getShowURL = shows
end function

public function getBackURL
	dim backurl
	backurl = THIS_PAGE&"?"&replace(request.querystring,"&","-")
	getBackURL = backurl
end function

public function getURLform
	dim show_requests,show_requests2,url
	url = ""
	'if request("page")<>"" then url = url & "<input type='hidden' name='page' value='"&request.querystring("page")&"'>" end if
	'if request("page_counter")<>"" then url = url & "<input type='hidden' name='page_counter' value='"&request.querystring("page_counter")&"'>" end if
	if REQUEST_ORDER<>"" then url = url & "<input type='hidden' name='order' value='"&REQUEST_ORDER&"'>" end if
	if REQUEST_ORDERBY<>"" then url = url & "<input type='hidden' name='orderby' value='"&REQUEST_ORDERBY&"'>" end if
	url = ""
	show_requests = Split(replace(Request.Querystring,"&&","&"),"&")
	for i = 0 to ubound(show_requests)
		if Left(lcase(show_requests(i)),2)="q_" then
			show_requests2 = Split(show_requests(i),"=")
			arg1 = cstr(show_requests2(0))
			if ubound(show_requests2)>0 then
				arg2 = cstr(show_requests2(1))
				url = url &"<input type='hidden' name='"&arg1&"' value='"&arg2&"'>"
			end if
		end if
	next
	getURLform = url&getURLFormShow&getURLFormGNR
end function

public function getURLFormGNR
	dim url
	url = ""
	show_requests = Split(Request.Querystring,"&")
	for i = 0 to ubound(show_requests)
		if Left(lcase(show_requests(i)),Len("gnr"))="gnr" then
			show_requests2 = Split(show_requests(i),"=")
			arg1 = cstr(show_requests2(0))
			if ubound(show_requests2)>0 then
				arg2 = cstr(show_requests2(1))
				url = url &"<input type='hidden' name='"&arg1&"' value='"&arg2&"'>"
			end if
		end if
	next
	getURLFormGNR = url
end function

public function getURLFormShow
	dim url
	url = ""
	show_requests = Split(Request.Querystring,"&")
	for i = 0 to ubound(show_requests)
		if Left(ucase(show_requests(i)),Len("SHOW_"))="SHOW_" then
			show_requests2 = Split(show_requests(i),"=")
			arg1 = cstr(show_requests2(0))
			if ubound(show_requests2)>0 then
				arg2 = cstr(show_requests2(1))
				url = url &"<input type='hidden' name='"&arg1&"' value='"&arg2&"'>"
			end if
		end if
	next
	getURLFormShow = url
end function

public function tableInf
	tableInf = "Listing Records:"&((getPage-1)*PAGE_COUNT+1)&"-"&((getPage-1)*PAGE_COUNT+PAGE_COUNT)&" | Filtered Records: "&RECORD_COUNT&" | Total Records:"&TOTAL_RECORD_COUNT&" | Page Number: <b>"&getPage&"</b> of "&int(RECORD_COUNT/PAGE_COUNT)+1
End function

public function createPageCounter
	response.write "<form action="""&THIS_PAGE&""" method=""get"" id=""form_page_counter"">Size: <select name=""page_counter"" onchange=""form_page_counter.submit();"">"
for k=25 to 250 step 25
	response.write "<option value='"&k&"'"
	if request.querystring("page_counter")<>"" then
		if int(request.querystring("page_counter"))=k then response.write " selected" end if
	end if
	response.write ">"&k&"</option>"
next
	response.write "</select>"
	response.write getURLform
	if request.querystring("custom_style")<>"" then response.write "<input type='hidden' name='custom_style' value='"&request.querystring("custom_style")&"'>" end if
	response.write "</form>"
end function

public function createPager(recordCount,page)
	response.write "<form action="""&THIS_PAGE&""" method=""get"" id=""form_pager"">Page: <select name=""page"" onchange=""form_pager.submit();"">"
for k=1 to int(recordCount/page)+1
	response.write "<option value='"&k&"'"
	if int(request.querystring("page"))=k then response.write " selected" end if
	response.write ">"&k&"</option>"
next
	response.write "</select>"
	response.write getURLform
	if request.querystring("page_counter")<>"" then response.write "<input type='hidden' name='page_counter' value='"&request.querystring("page_counter")&"'>" end if
	if request.querystring("custom_style")<>"" then response.write "<input type='hidden' name='custom_style' value='"&request.querystring("custom_style")&"'>" end if
	response.write "</form>"
end function

public function getWhere
dim t
Set RsF = Server.CreateObject("Adodb.Recordset")
RsF.Open "SELECT TOP 1 "&COLUMNS&" FROM "&TABLENAME,conn
t=0
where = " WHERE 1=1 "
for each x in RsF.Fields
	
		if x.type = 7 or x.type = 133 or x.type = 135 then
			if conn.Provider = "Microsoft.Jet.OLEDB.4.0" then
				if request.querystring("q_"&x.name&"_1")<>"" then
					mydate1 = split(replace(replace(request.querystring("q_"&x.name&"_1"),".","/"),"'",""),"/")
					where = where & " AND "&x.name&" > #"&mydate1(0)&"/"&mydate1(1)&"/"&mydate1(2)&"#"
				end if
				if request.querystring("q_"&x.name&"_2")<>"" then
					mydate2 = split(replace(replace(request.querystring("q_"&x.name&"_2"),".","/"),"'",""),"/")
					where = where & " AND "&x.name&" < #"&mydate2(0)&"/"&mydate2(1)&"/"&mydate2(2)&"#"
				end if
			else
				if request.querystring("q_"&x.name&"_1")<>"" then
					where = where & " AND "&x.name&" > convert(datetime,'"&replace(replace(request.querystring("q_"&x.name&"_1"),".","/"),"'","")&"',103)"
				end if
				if request.querystring("q_"&x.name&"_2")<>"" then
					where = where & " AND "&x.name&" < convert(datetime,'"&replace(replace(request.querystring("q_"&x.name&"_2"),".","/"),"'","")&"',103)"
				end if
			end if
			
		else
			if request.querystring("q_"&x.name)<>"" then
					where = where & " AND "&x.name&" LIKE '"&request.querystring("q_"&x.name)&"%'"		
			end if
		end if
	
	t=t+1
next
COLUMN_COUNT = t
getWhere = where
end function

public function getOrderby(how)
	Dim ordered
	if REQUEST_ORDER<>"" and REQUEST_ORDERBY<>"" and how="" then
		order = "ORDER BY "& lcase(REQUEST_ORDER) & " "&REQUEST_ORDERBY
	else
		if REQUEST_ORDERBY="desc" then
				if how="desc" then 
					how = "asc" 
				else
					how = "desc"
				end if
			ordered = REQUEST_ORDER
		elseif REQUEST_ORDERBY="asc" then
			ordered = REQUEST_ORDER
		else
			Set RsF = Server.CreateObject("Adodb.Recordset")
			RsF.OPEN "SELECT TOP 1 "&COLUMNS&" FROM "&TABLENAME,conn
			for each x in RsF.Fields
				if x.type<>201 and x.type<>203 and x.type<>205 then
					ordered = x.name
					exit for
				end if
			next
			if ordered="" then	ordered = LINKED_COLUMN	end if
			RsF.Close
		end if
		order = " ORDER BY "&ordered&" "&how
	end if
	getOrderby = order
end function

public function getPage
	if request("page")<>"" then
		getPage = int(request("page"))
	else
		getPage = 1
	end if
end function

public function getPageCount
	if request("page_counter")<>"" then
		getPageCount = request("page_counter")
	else
		getPageCount = 25
	end if
end function

public function getRecordCount
	Dim RsF
	Dim k
	k=0
	Set RsF = Server.CreateObject("Adodb.Recordset")
	RsF.Open "SELECT count(*) as MIKTAR FROM (SELECT "&COLUMNS&" FROM "&TABLENAME&") AS TBL "&getWhere,conn
	k = RsF("MIKTAR")
	RsF.close
	getRecordCount = k
end function

public function getTotalRecordCount
	Dim RsF
	Dim k
	k=0
	Set RsF = Server.CreateObject("Adodb.Recordset")
	RsF.Open "SELECT count(*) as MIKTAR FROM "&TABLENAME,conn
	k = RsF("MIKTAR")
	RsF.close
	getTotalRecordCount = k
end function

public function PrintTitle(field)
	start = Instr(CASES,"["&field.name&"=")
	if start>0 then
		start = start + Len("["&field.name&":")
		finish = Instr(start,CASES,"|")
		PrintTitle = Mid(CASES,start,finish-start)
	else
		PrintTitle = CStr(UCase(Left(field.name,1))&Lcase(Right(field.name,Len(field.name)-1)))
	end if
end function

'Searchi gelistirmek icin recordlardan alinan bazi fonksiyonlar
public function getFieldFor(field,taborder)
	call user("getFieldFor_begin")
	start = InStr(CASES,"|"&field.name&"|")+Len(field.name)+2
	if start>(Len(field.name)+2) then
		finish = InStr(start,CASES,"|")
		call user("getFieldFor_end")
		if (Instr(Mid(CASES,start,finish-start),"useAjax") or Instr(Mid(CASES,start,finish-start),"createLookup")) then
			getFieldFor = replace(Eval(""&Mid(CASES,start,finish-start)&""),"id="""&field.name&"_id""","tabindex="""&taborder&""" id="""&field.name&"_id""")
		else
			getFieldFor = replace(Mid(CASES,start,finish-start),"id="""&field.name&"_id""","tabindex="""&taborder&""" id="""&field.name&"_id""")
		end if
	else
		call user("getFieldFor_end")
		getFieldFor = ""
	end if
end function

public function useAjax(the_id,the_text,table_from,info,style,field)
	Dim html
	Dim RsF
	Dim lookSql
	call user("useAjax_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	lookSql = "SELECT TOP 1 "&the_text
	if info<>"" then lookSql = lookSql & ","&info end if
	lookSql = lookSql & " FROM "&table_from
	RsF.open lookSql,conn
		html = "<input type=""hidden"""
		if request.querystring("q_"&field.name)<>"" then
			html =html& " value="""&request.querystring("q_"&field.name)&""""
		end if
		html = html&""" name=""q_"&field.name&""" id=""q_"&field.name&""" "&style&" /><input type=""text"" id=""q_"&field.name&"_id"" onchange=""control_ajax(this.id)"" "&style
	if request.querystring("q_"&field.name)<>"" then
		html =html& " value="""&request.querystring("q_"&field.name)&""""
	end if
		html = html &  " style=""width:100%;"" />"
	RsF.close
	inscript =  "var options_xml_q_"&field.name&" = {"&chr(13)
	inscript =	inscript &"script: function (input) { return """&THIS_PAGE&"?ID="&the_id&"&VIEW="&the_text&"&KEYWORD=""+document.getElementById('q_"&field.name&"_id').value+""&INFO="&info&"&TABLE="&table_from&"&AJAX=1&""; },"&chr(13)
	inscript =  inscript &"varname:""input""," & chr(13)
	inscript =  inscript &"callback: function (obj) { document.getElementById('q_"&field.name&"').value = obj.id; }," & chr(13)
	inscript =  inscript &"json: true"& chr(13)
	inscript =  inscript &" };" & chr(13)& chr(13)
	inscript =  inscript &"var as_xml_q_"&field.name&" = new bsn.AutoSuggest('q_"&field.name&"_id', options_xml_q_"&field.name&");" & chr(13)
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
	if the_text=the_id then sql = replace(sql,the_text&","&the_text,the_text) end if
	RsF.open "SELECT TOP 1 "&the_text&" FROM "&table_from,conn
	if not (RsF.fields(0).type=201 or RsF.fields(0).type=203 or RsF.fields(0).type=205) then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		sql = sql&" ORDER BY "&the_text&" asc"
	end if
	RsF.close
	if the_text<>the_id then
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
	else
		RsF.open sql,conn
		results =  "{ results: [ " & vbCrLf
		sayac = 0
			do until RsF.eof
				if the_info<>"" then
					results = results & "	{ id: """&RsF(0)&""", value: """&RsF(0)&""", info: """&RsF(1)&""" },"&vbCrLf
				else
					results = results & "	{ id: """&RsF(0)&""", value: """&RsF(0)&""", info: """" },"&vbCrLf
				end if
				RsF.Movenext
				sayac = sayac + 1
			loop
				results = left(results,Len(results)-3)&vbCrLf
			results = results&"] }"
		RsF.close
	end if
	call user("showAjax_end")
	showAjax = results
end function

public function createLookup(the_id,the_text,table_from,filter,style,field)
	Dim RsF
	call user("createLookup_begin")
	Set RsF = Server.CreateObject("Adodb.Recordset")
	if the_text<>the_id then	
		sql = "SELECT "&the_text&","&the_id&" FROM "&table_from
	else
		sql = "SELECT "&the_id&" FROM "&table_from
	end if
	if filter<>"" then
		sql = sql& " WHERE "&filter
	end if
	RsF.open "SELECT TOP 1 "&the_text&" FROM "&table_from,conn
	if not (RsF.fields(0).type=201 or RsF.fields(0).type=203 or RsF.fields(0).type=205) then '201 = adLongVarChar, 203 = adLongVarWChar, 205= adLongVarBinary
		sql = sql&" ORDER BY "&the_text&" asc"
	end if
	RsF.close
	RsF.open sql,conn
	selection = 0
	do until RsF.eof
		result = result & "<option value='"&RsF(the_id)&"'"
		if request.querystring("q_"&the_id)=RsF(the_id) then
			selection = 1
			result= result&" selected"
		end if
		result = result & ">"&RsF(0)&"</option>"
		RsF.Movenext
	loop
	
		result = "<option value=''>Seciniz</option>"&result
	
	result = "<select name=""q_"&field.name&"""  id=""q_"&field.name&"_id""  onchange=""f_form.submit()"" style=""width:100%;"" "&style&">"&result&"</select>"
	call user("createLookup_end")
	createLookup = result
end function

Function getFileName()
  Dim lsPath, arPath
  lsPath = Request.ServerVariables("SCRIPT_NAME")
  arPath = Split(lsPath, "/")
  GetFileName =arPath(UBound(arPath,1))&""
End Function

Function getID(mytable)
	getID = ""
	Set adox = CreateObject("ADOX.Catalog")
	adox.activeConnection = conn
	for each table in adox.tables
		if ucase(table.name)=ucase(mytable) then
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

Function setStyle()
	if PAGE_STYLE="" then
		if request.querystring("custom_style")<>"" then
			if IsNumeric(request.querystring("custom_style")) then
				PAGE_STYLE = getStyleFromNumber(request.querystring("custom_style"))
			else
				PAGE_STYLE = request.querystring("custom_style")
			end if
		else
			PAGE_STYLE = "Glassy"
		end if
	end if
end Function

Function getStyleFromNumber(style_no)
	dim fs,fo,x,sayac,style_return
	sayac = 0
	style_return = ""
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set fo=fs.GetFolder(Server.Mappath("css\multistyle"))
	
	for each x in fo.SubFolders
		sayac = sayac + 1
		if sayac=int(style_no) then
			style_return = x.Name
			exit for
		end if
	next
	if style_return="" then
		style_return = "Glassy"
	end if
	getStyleFromNumber = style_return
end Function

Function createStyleSelector
	dim fs,fo,x,sayac,html
	sayac = 0
	html = ""
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set fo=fs.GetFolder(Server.Mappath("css\multistyle"))
	html = "<form action="""&THIS_PAGE&""" method=""get"" id=""form_style"">"
	html = html & "<select name=""custom_style"" onchange=""form_style.submit();"" style=""width:100%;"">"
	for each x in fo.SubFolders
		sayac = sayac + 1
		html = html & "<option value="""&sayac&""""
		if PAGE_STYLE=x.Name then html = html & " selected" end if
		html = html & ">"&x.Name
		html = html & "</option>"
	next
	html = html & "</select>"
	html = html & getURLform
	html = html & "<input type='hidden' name='page_counter' value='"&getPageCount()&"'>"
	html = html & "<input type='hidden' name='page' value='"&getPage&"'>"
	html = html & "</form>"
	createStyleSelector = html
end Function
%>