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
THIS_PAGE = getFileName()
RECORD_ID = getID(TABLENAME)
if RECORD_ID="" then RECORD_ID = LINKED_COLUMN end if
JAVASCRIPT = ""	
REQUEST_LINKED_COLUMN = request(LINKED_COLUMN)
call user("after_definations_record_body")
if request("AJAX")=1 then
	if request("call")="generic" then
		response.write genericAjax(request.querystring("id_field"),request.querystring("text_field"),request.querystring("table_from"),request.querystring("filter"))
	else
		response.write showAjax(request.querystring("ID"),request.querystring("VIEW"),request.querystring("INFO"),request.querystring("TABLE"),request.querystring("KEYWORD"))
	end if
	response.end
end if

if Request.Form("back")<>"" then
	if Request.Form("back_url")<>"" then
		response.redirect Request.Form("back_url")
	else
		response.redirect LIST_PAGE
	end if
elseif Request.Form("kaydet")<>"" then
	if request(RECORD_ID)="NEWID" then
		if KAYDET=1 then call save_new_record end if
		if Request.Form("back_url")<>"" then
			response.redirect Request.Form("back_url")
		else
			response.redirect LIST_PAGE
		end if
	else
		if KAYDET=1 then call save_record end if
		if Request.Form("back_url")<>"" then
			response.redirect Request.Form("back_url")
		else
			response.redirect LIST_PAGE
		end if
	end if
	response.write "<table align='center'><tr><td align='center'><h2><font color='green'>Veriler Kaydedildi !</font></h2></td></tr></table>"
elseif Request.Form("delete")<>"" then
	if SIL=1 then call delete_record end if
	response.redirect LIST_PAGE
end if



if REQUEST_LINKED_COLUMN<>"" then
	if IsNumeric(REQUEST_LINKED_COLUMN) then
		where_statement = "WHERE "&LINKED_COLUMN &"="&REQUEST_LINKED_COLUMN
	else
		if REQUEST_LINKED_COLUMN="NEWID" then
			where_statement =""
		else
			where_statement = "WHERE "&LINKED_COLUMN &"='"&REQUEST_LINKED_COLUMN&"'"
		end if
	end if
	
	
	'3 ARALIK 2009 HATA TESPIT EDILDI //ALP
	'if not (Instr(COLUMNS,RECORD_ID)>0 or COLUMNS="*") then
	if RECORD_ID<>"" then
		if not (Instr(replace(" "&COLUMNS&" ",","," ")," "&RECORD_ID&" ")>0 or COLUMNS="*") then
			EXTRA_COLUMN=","&RECORD_ID
		else
			EXTRA_COLUMN=""
		end if
	end if
	
	on error resume next
	sql = "SELECT top 1 "&COLUMNS&EXTRA_COLUMN&" FROM "&TABLENAME&" "&where_statement
	Set Rs = Server.CreateObject("Adodb.Recordset")
	rs.open sql,conn
	
	
	
	if err<>0 then
		response.write "Error 1:"&err.description&"<hr> SQL: "&sql&"<hr>"
		response.write "<hr>Columns : "&COLUMNS&"<hr>"&EXTRA_COLUMN&"<hr>"
		response.write "Rs.eof : "&rs.eof
		response.end
	end if
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<% if HTMLEDITOR=1 then%><script type="text/javascript" src="js/ckeditor/ckeditor_basic.js"></script><%end if%>
		<% if Instr(CASES,"jscolor")>0 then%><script type="text/javascript" src="js/jscolor.js"></script><%end if%>
		<script type="text/javascript">
			function control_ajax(gid)
			{
				if (document.getElementById(gid).value=="")
				{
					document.getElementById(gid.replace("_id","")).value = "";
				}
			}
			function form_control()
			{
				var zorunlu_fieldlar="<%=REQUIRED_COLUMNS%>".split(" ");
				var durum = "true";
				var mesaj = "Lutfen asagidaki alanlari bos birakmayiniz:\n\n";
				for(var i=0;i<zorunlu_fieldlar.length;i++)
				{
					if (document.getElementById(zorunlu_fieldlar[i]+"_id")!=null)
					{
						if (document.getElementById(zorunlu_fieldlar[i]+"_id").value=="" || document.getElementById(zorunlu_fieldlar[i]+"_id").value=="0")
						{
							durum = "false";
							mesaj = mesaj + zorunlu_fieldlar[i].replace("_id","")+"\n";
						}
					}
				}
				if (durum=="false")
				{
					alert(mesaj);
					return false;
				}
			}
			<% if Instr(CASES,"createGenericLookup")>0 then%>
			var refreshProcess = " ";
			function RefreshCombo(select_id,id_field,text_field,table_from,filter)
			{
				if (refreshProcess.indexOf(select_id)>0){
					alert("This combo is not available to refresh now..");
					return;
				}
				refreshProcess += " "+select_id+" ";
				var myObject = document.getElementById(select_id);
				var originval = myObject.options[myObject.selectedIndex].value;
				while(myObject.selectedIndex!=-1)
					myObject.remove(myObject.selectedIndex);
				myObject.disabled = true;
				var y = document.createElement('option');
				y.text = 'Refreshing options, please wait..';
				try{
					myObject.add(y,null); 
				}
				catch(ex){
					myObject.add(y); // IE only
				}
				var xmlhttp;
				if (window.XMLHttpRequest)
					xmlhttp = new XMLHttpRequest();
				else
					xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
				xmlhttp.open("GET","<%=THIS_PAGE%>?ajax=1&call=generic&id_field="+id_field+"&text_field="+text_field+"&table_from="+table_from+"&filter="+filter+"&rnd_num="+Math.random(),false);
				xmlhttp.send();
				var myOptions = eval('('+xmlhttp.responseText+')');
				myObject.remove(myObject.selectedIndex);
				for(var i = 0;i<myOptions.option.length;i++)
				{
					var t = document.createElement('option');
					t.text = myOptions.option[i].text;
					t.value = myOptions.option[i].value;
					if (originval == t.value)
						t.selected = true;
					try{
						myObject.add(t,null); 
					}
					catch(ex){
						myObject.add(t); // IE only
					}
				}
				refreshProcess = refreshProcess.replace(" "+select_id+" ","");
				alert("Refresh Done!")
				myObject.disabled = false;
				myObject.focus();
			}
			<% end if %>
		</script>
		<link rel=stylesheet href="stylee.css" type="text/css">
		<script type="text/javascript" src="js/calendarDateInput.js">
		</script>
		<script type="text/javascript" src="js/bsn.AutoSuggest_2.1.3_comp.js" charset="utf-8" >
		</script>
		<meta http-equiv="content-type" content="text/html; charset=utf-8">
		<link rel="stylesheet" type="text/css" href="Glassy/Style_doctype.css">
		<link rel="stylesheet" href="css/autosuggest_inquisitor.css" type="text/css" media="screen" charset="utf-8">
		<title>Kayıt</title>
	</head>
	<body><%=BEFORE_TABLE_HTML%>
			<form action="<%=THIS_PAGE%>" method="post" onsubmit="return form_control()">
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
					if REQUEST_LINKED_COLUMN="NEWID" then ' yeni kayit
						for each x in rs.Fields
							if LCase(x.name) = LCase(RECORD_ID) then
								if not (EXTRA_COLUMN<>"") then ' Secilmeyen IDnin yazilmasi sevimsiz olur
								%>
								<tr class="Controls"><td class="th"><label for="<%=x.name%>"><%=PrintTitle(x)%></label></td><td><input type="text" size="15" id="<%=x.name%>_id"  name="<%=x.name%>" value="NEWID" readonly></td></tr>
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
					else ' update
						for each x in rs.Fields
							if LCase(x.name) = LCase(RECORD_ID) then
								if not (EXTRA_COLUMN<>"") then ' Secilmeyen IDnin yazilmasi sevimsiz olur
								%>
								<tr class="Controls"><td class="th"><label for="<%=x.name%>_id"><%=PrintTitle(x)%></label></td><td><input type="text" size="15" id="<%=x.name%>_id"  name="<%=x.name%>" value="<%=x.value%>" readonly></td></tr>
								<% else %>
								<tr class="Controls"><td colspan="2" class="th"><input type="hidden" name="<%=x.name%>" value="<%=x.value%>"></td></tr>
								<%
								end if
							else
						%>
							<tr class="Controls"><td class="th"><label for="<%=x.name%>_id"><%=PrintTitle(x)%></label></td><td><%=PrintField(x)%></td></tr>
					<%
							end if
						next
					end if
					%>
					<tr class="Bottom"><td colspan="2"  align="right">
					<% if KAYDET=1 then %><input name="kaydet" type="submit" class="Button" value="Kaydet"><% end if %>
					<% if GERIAL=1 then %><input type="reset" class="Button" value="Tümünü Geri Al"><% end if %>
					<% if SIL=1 then %><input type="submit" name="delete" class="Button" value="Sil"><% end if%>
					<% if GENERICPAGE=1 then %><input type="hidden" name="gnrTABLENAME"  value="<%=request("gnrTABLENAME")%>"><% end if%>
					<% if GERIDON=1 then %><input type="hidden" name="back_url" value="<%=replace(request("back_url"),"-","&amp;")%>"><input type="submit" name="back" class="Button" value="Geri Dön"><% end if %>
					</td></tr>
				</table>
			</table>
			</form>
	<% end if %>
		<script type="text/javascript">
		<%=JAVASCRIPT%>
		</script>
	</body>
</html>
<%
call user("record_body_end")
%>