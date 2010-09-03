<%
if SHOW_ONLYPAGE="" then SHOW_ONLYPAGE = request.querystring("SHOW_ONLYPAGE") end if
if SHOW_HEADER="" then SHOW_HEADER = request.querystring("SHOW_HEADER") end if
if SHOW_FOOTER="" then SHOW_FOOTER = request.querystring("SHOW_FOOTER") end if
if SHOW_TITLE="" then SHOW_TITLE = request.querystring("SHOW_TITLE") end if
if SHOW_PAGER="" then SHOW_PAGER = request.querystring("SHOW_PAGER") end if
if SHOW_PAGE_SIZER="" then SHOW_PAGE_SIZER = request.querystring("SHOW_PAGE_SIZER") end if
if SHOW_ALL_ICONS="" then SHOW_ALL_ICONS = request.querystring("SHOW_ALL_ICONS") end if
if SHOW_CAPTIONS="" then SHOW_CAPTIONS = request.querystring("SHOW_CAPTIONS") end if
if SHOW_FILTERS="" then SHOW_FILTERS = request.querystring("SHOW_FILTERS") end if
if SHOW_STYLER="" then SHOW_STYLER = request.querystring("SHOW_STYLER") end if
if SHOW_TABLE_INFO="" then SHOW_TABLE_INFO = request.querystring("SHOW_TABLE_INFO") end if
if SHOW_WIDTH="" then SHOW_WIDTH = request.querystring("SHOW_WIDTH") end if
if SHOW_TITLE="" then SHOW_TITLE = true end if
if SHOW_PAGER="" then SHOW_PAGER = true end if
if SHOW_PAGE_SIZER="" then SHOW_PAGE_SIZER = true end if
if SHOW_ALL_ICONS="" then SHOW_ALL_ICONS = true end if
if SHOW_CAPTIONS="" then SHOW_CAPTIONS = true end if
if SHOW_FILTERS="" then SHOW_FILTERS = true end if
if SHOW_STYLER="" then SHOW_STYLER = true end if
if SHOW_TABLE_INFO="" then SHOW_TABLE_INFO = true end if
if SHOW_HEADER="" then SHOW_HEADER = true end if
if SHOW_FOOTER="" then SHOW_FOOTER = true end if
if SHOW_ONLYPAGE="" then SHOW_ONLYPAGE = false end if
if SHOW_WIDTH="" then 
	SHOW_WIDTH = "80%" 
else
	SHOW_WIDTH = SHOW_WIDTH&"%" 
end if






if LINKED_COLUMN="" then LINKED_COLUMN = getID(TABLENAME) end if
THIS_PAGE = getFileName()
call setStyle()
if request("AJAX")=1 then
	response.clear
	response.write showAjax(request.querystring("ID"),request.querystring("VIEW"),request.querystring("INFO"),request.querystring("TABLE"),request.querystring("KEYWORD"))
	response.end
end if
if request.Cookies(THIS_PAGE&"_column_order")<>"" then COLUMNS = request.Cookies(THIS_PAGE&"_column_order") end if
REQUEST_ORDER = REQUEST("order")
REQUEST_ORDERBY = REQUEST("orderby")
RECORD_COUNT = getRecordCount
TOTAL_RECORD_COUNT = getTotalRecordCount
if instr(RECORD_PAGE,"?")>0 then
	LINK_FORMAT = RECORD_PAGE&"&"&LINKED_COLUMN&"="
else
	LINK_FORMAT = RECORD_PAGE&"?"&LINKED_COLUMN&"="
end if
if Instr(LINK_FORMAT,"www.")>0 and Left(LINK_FORMAT,7)<>"http://" then
	LINK_FORMAT = "http://"&LINK_FORMAT
end if
PAGE_COUNT = getPageCount	
call user("after_definations")
if SHOW_ONLYPAGE then
	response.clear()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8">
		<title>Liste</title>
		<link rel="stylesheet" type="text/css" href="css/multistyle/<%=PAGE_STYLE%>/Style_doctype.css">
		<link rel="stylesheet" href="css/autosuggest_inquisitor.css" type="text/css" media="screen" charset="utf-8" />
		<link type="text/css" href="css/datepicker/ui.all.css" rel="stylesheet" />
		<script type="text/javascript" src="js/bsn.AutoSuggest_2.1.3_comp.js" charset="utf-8">
		</script>
		<script type="text/javascript" src="js/datepicker/jquery-1.3.2.js">
		</script> 
		<script type="text/javascript" src="js/datepicker/ui.core.js">
		</script> 
		<script type="text/javascript" src="js/datepicker/ui.datepicker.js">
		</script> 
		<script type="text/javascript" src="js/datepicker/ui.datepicker-tr.js">
		</script> 
		<script type="text/javascript" src="js/calendarDateInput.js">
		</script>
		<script type="text/javascript">
		function control_ajax(gid)
			{
				if (document.getElementById(gid).value=="")
				{
					document.getElementById(gid.replace("_id","")).value = "";
				}
			}
		</script>
		<style type="text/css">
		.parlak  td { 
			font-size: 100%;
			padding: 3px;
			border-top: 1px solid #FFFFFF; border-right: 1px solid #FFFFFF;
			text-align: left;
			vertical-align: top;
			background-color: pink; 
			color: #000000;
			}
		tr.c4 {cursor:hand;}
		input.c3 {width:100%;}
		input.c5 {width:49%;}
		td.c2 {border:0px;}
		td.c1 {text-align:center;}
		</style>
	</head>
	<body>
<%
'(Siralama Linkine tiklandiginda..) Daha once artan siralandiysa simdi azalan sirala
if REQUEST_ORDERBY="desc" then
	orderby = "asc"
else
	orderby= "desc"
end if


'Son sayfada sayfa boyutundan kucuk miktardaki kayit sayilari icin..
if int(PAGE_COUNT)>int(RECORD_COUNT-((getPage-1)*PAGE_COUNT)) then
	LESS_COUNT = int(int(RECORD_COUNT)-int((getPage-1)*PAGE_COUNT))
else
	LESS_COUNT = int(PAGE_COUNT)
end if
if LESS_COUNT<1 then LESS_COUNT = 1 end if

if not (Instr(COLUMNS,LINKED_COLUMN)>0 or COLUMNS="*") and LINKED_COLUMN<>"" then
	EXTRA_COLUMN=","&LINKED_COLUMN
else
	EXTRA_COLUMN=""
end if
if EDITABLE=1 then
	call updateFields()
end if
'Sayfanin SQL'i sayfalamada sorun cikmamasi icin ic ice 3 sqlden olusur asc desc asc seklinde listeyi siralar
sql = "select top "&LESS_COUNT&" * from (SELECT top "&(PAGE_COUNT*getPage)&" * FROM (SELECT "&COLUMNS&EXTRA_COLUMN&" FROM "&TABLENAME&") AS TBL "&getWhere&getOrderby("asc")&") as TBL "&getOrderby("desc")
sql = "SELECT TOP "&PAGE_COUNT&" * FROM (" & sql & ") as TBL "&getOrderby("asc")
call user("after_sql_definations")
on error resume next
rs.open sql , conn
if err<>0 then
	response.write err.description&"<br>"
	response.write sql
	response.end
end if
sayac=1+((getPage-1)*PAGE_COUNT)
%>

		<table width="<%=SHOW_WIDTH%>%" align="center" cellspacing="0" cellpadding="0">
			<tr>
				<td valign="top">
					<% if SHOW_TITLE and SHOW_HEADER then%>
					<table cellspacing="0" cellpadding="0" border="0" class="Header" width="100%">
					  <tr>
						<td class="HeaderLeft"> </td> 
						<td class="th c1"><strong><%=LIST_TITLE%></strong></td>
						<td class="HeaderRight"> </td> 
					  </tr>
					</table>
					<% end if %>
			<table class="Grid" cellspacing="0" cellpadding="0">
				<% if SHOW_HEADER then %>
				<tr class="Footer">
					<td colspan="<%=(COLUMN_COUNT+1)%>" align="center">
						<table align="center" cellpadding="0" width="100%">
							<tr>
								<th align="left"><% if SHOW_PAGER then%><% call createPager(RECORD_COUNT,PAGE_COUNT) %><% end if %></th>
								<th>
									<% if SHOW_ALL_ICONS then%>
									<table align="center" border="0" width="50%">
										<tr>
											<td class="c2">
												<a href="<%=RECORD_PAGE%><%if instr(RECORD_PAGE,"?")>0 then response.write "&amp;" else response.write "?" end if%><%=LINKED_COLUMN&"=NEWID&amp;back_url="&getBackURL%>&amp;custom_style=<%=PAGE_STYLE%>"><img src="css/images/new1.png" border="0" alt="Yeni Kayıt" width="30" height="30"></a>
											</td>
											<td class="c2">
												<a href="<%=replace((THIS_PAGE&"?EXPORT_STYLE=xls&order="&request("order")&"&orderby="&request("orderby")&"&page_counter="&getPageCount&getURL&"&page="),"&","&amp;")%><%if int(request("page"))=0 then response.write "1" else response.write int(request("page")) end if%><%="&amp;gnrTABLENAME="&request("gnrTABLENAME")%>"><img src="css/images/exel1.png" border="0" alt="Verileri Excele Yolla" width="30" height="30"></a>
											</td>
											<td class="c2">
												<a href="<%=replace((THIS_PAGE&"?EXPORT_STYLE=doc&order="&request("order")&"&orderby="&request("orderby")&"&page_counter="&getPageCount&getURL&"&gnrTABLENAME="&request("gnrTABLENAME")),"&","&amp;")%>"><img src="css/images/word1.png" border="0" alt="Verileri Worde Yolla" width="30" height="30"></a>
											</td>
											<td class="c2">
												<a href="inc/column_editor.asp?column_editor=1&amp;page_name=<%=THIS_PAGE%>"><img src="css/images/setting1.png" border="0" alt="Sütunlari Duzenle" width="30" height="30"></a>
											</td>
										</tr>
									</table>
									<% end if%>
								</th>
								<th align="right"><% if SHOW_PAGE_SIZER then%><% call createPageCounter%><%end if%></th>
							</tr>
						</table>
					</td>
				</tr>
			<% end if %>
			<%
			if SHOW_CAPTIONS then
				'Filtreleme Inputlari + Siralama Linklerini Yazdir
				response.write "<tr class=""Caption""><th scope=""col"">#</th>"
				taborder=1
				for each x in rs.Fields
					if not (EXTRA_COLUMN<>"" and LCASE(x.name)=LCASE(LINKED_COLUMN)) then ' Secilmeyen IDnin yazilmasi sevimsiz olur
						response.write "<th scope=""col"" align=""center"">"
						'Bu tarz fieldlarda order by kullanilamaz
						if not (x.type=201 or x.type=203 or x.type=205) then
							response.write "<a href='"&replace(THIS_PAGE&"?order="&x.name &"&orderby="&orderby&getURL&"&page_counter="&getPageCount,"&","&amp;")
							if GENERICPAGE=1 then response.write "&amp;gnrTABLENAME="&trim(request.querystring("gnrTABLENAME")) end if
							response.write "'>"
						end if
						response.write PrintTitle(x)
						if REQUEST_ORDER=x.name and REQUEST_ORDERBY="asc" then
							response.write "<img border=""0"" src=""css/multistyle/"&PAGE_STYLE&"/Images/Asc.gif"" alt='Artan'>"
						elseif REQUEST_ORDER=x.name and REQUEST_ORDERBY="desc" then
							response.write "<img border=""0"" src=""css/multistyle/"&PAGE_STYLE&"/Images/Desc.gif"" alt='Azalan'>"
						end if
					
						if not (x.type=201 or x.type=203 or x.type=205) then
							response.write "</a></th>"
						end if
						
					end if
				next
				response.write "</tr>"&chr(13)
			end if
			if SHOW_FILTERS then
				response.write "<form action='"&THIS_PAGE&"' method='get' id='f_form'><tr width='100%'><td bgcolor=""white"" align=""left"">"
				if GENERICPAGE=1 then
					response.write "<input type='hidden' name='gnrTABLENAME' value='"&trim(request.querystring("gnrTABLENAME"))&"'>"
				end if
				response.write "<input type='hidden' name='order' value='"&REQUEST_ORDER&"'><input type='hidden' name='orderby' value='"&REQUEST_ORDERBY&"'>"
				response.write " <input type='hidden' name='page_counter' value='"&getPageCount()&"'></th>"&chr(13)
				'Filtreleme Inputlari
				for each x in rs.Fields
					if not (EXTRA_COLUMN<>"" and LCASE(x.name)=LCASE(LINKED_COLUMN)) then ' Secilmeyen IDnin yazilmasi sevimsiz olur
						response.write "<th>"
						'Gelismis aramalar icin combobox veya ajax
						got_field = getFieldFor(x,taborder)
						if got_field="" then 
							if x.type = 7 or x.type = 133 or x.type = 135 then 'Tarih
								response.write "<input type='text' tabindex="&taborder&" class='c5' name='q_"&x.name&"_1' id='q_"&x.name&"_1_id' value='"&request.querystring("q_"&x.name&"_1")&"' onkeypress='{if (event.keyCode==13)f_form.submit()}'>"
								response.write "<input type='text' tabindex="&(taborder+1)&" class='c5' name='q_"&x.name&"_2' id='q_"&x.name&"_2_id' value='"&request.querystring("q_"&x.name&"_2")&"' onkeypress='{if (event.keyCode==13)f_form.submit()}'></th>"&chr(13)
								taborder=taborder+1
								TARIHKUTULARI = TARIHKUTULARI & " q_"&x.name&"_1_id q_"&x.name&"_2_id"
							else
								response.write "<input type='text' tabindex="&taborder&" class='c3' name='q_"&x.name&"' value='"&request.querystring("q_"&x.name)&"' onkeypress='{if (event.keyCode==13)f_form.submit()}'></th>"&chr(13)
							end if
						else
							response.write got_field&"</th>"&chr(13)
						end if
						taborder=taborder+1
					end if
				next
				response.write "</tr></form>"&chr(13)
			end if
			if EDITABLE=1 then 	response.write "<form action="""&THIS_PAGE&""" method=""post"">" end if
			'Kayitlari Yazdir
			rec_Ids = ""
			do until rs.EOF
				if (sayac mod 2)=1 then
					color = "Row"
				else
					color = "Altrow"
				end if
				
				'Her kayidin icindeki sutunlar ve her kayidin link idsini baslamadan once temizle
				myTDs = ""
				LinkId=""
				rec_Ids = rec_Ids & Rs(LINKED_COLUMN)&" "
				for each x in rs.Fields
					if (LCase(x.name) = LCase(LINKED_COLUMN)) then LinkId = x.value end if
					if not (EXTRA_COLUMN<>"" and LCASE(x.name)=LCASE(LINKED_COLUMN)) then ' Secilmeyen IDnin yazilmasi sevimsiz olur
						myTDs=myTDs&"<td>"
						if EDITABLE=1 then
							
							if lcase(x.name)=lcase(LINKED_COLUMN) then
								print = "<input type=""disabled"" value="""&x.value&""">"
							else
								print = PrintField(x,rs(LINKED_COLUMN))
							end if
						else
							if x.type=7 or  x.type=133  or  x.type=135 then '133=adDBDate , 7 = adDate , 135=adDBTimeStamp
								ay = Month(x.value)
								gun = Day(x.value)
								yil = Year(x.value)
								if Len(ay)=1 then ay="0"&ay end if
								if Len(gun)=1 then gun="0"&gun end if
								if ay>0 and gun>0 and yil>0 then
									print = gun&"/"&ay&"/"&yil
								else
									print = "-"
								end if
							else
								print = x.value
							end if
							if Len(print)>1500 then
								print= Left(print,1500)&"..."
							end if
							print=print&"&nbsp;"
						end if
						myTDs=myTDs&print
					end if
				next
				
				'Satiri yaz
				%>
				<tr class="c4 <%=color%>" <% if EDITABLE<>1 AND LINKED_COLUMN<>"" then%>onclick="window.location.assign('<%=replace(LINK_FORMAT&LinkId&"&custom_style="&PAGE_STYLE&"&back_url="&getBackURL,"&","&amp;")%>');"<%end if%> onmouseover="this.className='parlak'" onmouseout="this.className='<%=color%>'"><td align='center'><b><%=sayac%></b></td><%=myTds%></tr>
				<%
				sayac = sayac+1
				rs.MoveNext
			loop
			if EDITABLE=1 then 
				rec_Ids = trim(rec_Ids)
				response.write "<input type=""hidden"" value="""&rec_Ids&""" name=""REC_IDS"">"
				%>
				<tr class="Footer"><td colspan="<%=(COLUMN_COUNT+1)%>" align="center"><input type="submit" value="Update Records"></td></tr>
				<%
				response.write "</form>" 
			end if
			%>
			
			<%if SHOW_FOOTER then%>
				<tr class="Footer"><td style="max-width:20px;"><%if SHOW_STYLER then%><%=createStyleSelector%><%end if%></td><td colspan="<%=(COLUMN_COUNT)%>" align="center"><b><%if SHOW_TABLE_INFO then%><%=tableInf%></b><%end if%></td></tr>
			<%end if%>
			</table>
		</table>
<script type="text/javascript">
	<%=JAVASCRIPT%>
</script>
<% if TARIHKUTULARI<>"" then %>
<script type="text/javascript">
	$(function() {
		$.datepicker.setDefaults($.extend({showMonthAfterYear: false,changeMonth: true,changeYear: true}, $.datepicker.regional['']));
	<% 
	TARIHKUTULARI = Right(TARIHKUTULARI,Len(TARIHKUTULARI)-1)
	tarih_kutucuklari = Split(TARIHKUTULARI)
	for i = 0 to ubound(tarih_kutucuklari)
		response.write "$(""#"&tarih_kutucuklari(i)&""").datepicker($.datepicker.regional['tr']);"&chr(13)
	next
	%>});
</script>
<%end if%>
	</body>
</html>
<% 
if REQUEST_EXPORT_STYLE<>"" then
	'Excel ise
	Response.Clear()
	Response.Buffer = True
	Response.AddHeader "Content-Disposition", "attachment;filename="&replace(LCASE(LIST_TITLE)," ","_")&"."&REQUEST_EXPORT_STYLE
	
	if REQUEST_EXPORT_STYLE="doc" then
		Response.ContentType = "application/vnd.ms-word"
	else
		Response.ContentType = "application/vnd.ms-excel"
	end if
	response.write "<meta http-equiv=""content-type"" content=""text/html; charset=UTF-8"">"
	response.write "<table border='1'><tr align='center'><td bgcolor='#CEE6A3'><b>#</b></td>"
	'Basliklar
	for each x in rs.Fields
		if not (EXTRA_COLUMN<>"" and LCASE(x.name)=LCASE(LINKED_COLUMN)) then ' Secilmeyen IDnin yazilmasi sevimsiz olur
			response.write "<td bgcolor='#CEE6A3'><b>"
			response.write UCase(Left(x.name,1))&Right(LCase(x.name),Len(x.name)-1)
			response.write "</b></td>"
		end if
	next
	response.write "</tr>"
	sayac=1
	do until rs.EOF
		'Her kayidin icindeki sutunlar ve her kayidin link idsini baslamadan once temizle
		if (sayac mod 2)=0 then
			mycolor="#C4C5EE"
		else
			mycolor="#EDEDED"
		end if
		myTDs = ""
		for each x in rs.Fields
			if not (EXTRA_COLUMN<>"" and LCASE(x.name)=LCASE(LINKED_COLUMN)) then ' Secilmeyen IDnin yazilmasi sevimsiz olur
				myTDs=myTDs&"<td bgcolor='"&mycolor&"' >"
				if x.type=7 or  x.type=133  or  x.type=135 then '133=adDBDate , 7 = adDate , 135=adDBTimeStamp
					ay = Month(x.value)
					gun = Day(x.value)
					yil = Year(x.value)
					if Len(ay)=1 then ay="0"&ay end if
					if Len(gun)=1 then gun="0"&gun end if
					if ay>0 and gun>0 and yil>0 then
						print = gun&"/"&ay&"/"&yil
					else
						print = "-"
					end if
				else
					print = x.value
				end if
				myTDs=myTDs& print
				myTDs=myTDs&"</td>"
			end if
		next
		'Satiri yaz
		response.write "<tr align='center'><td bgcolor='#CEE6A3'><b>"&sayac&"</b></td>"&myTds&"</tr>"
		sayac = sayac+1
		rs.MoveNext
	loop
	response.write "<tr align='center'><td bgcolor='#CEE6A3' colspan='"&(COLUMN_COUNT+1)&"'><b>"&tableInf&"</b></td></tr>"
	response.write "</table>"
	response.end
end if

if SHOW_ONLYPAGE then
	response.end
end if
%>

