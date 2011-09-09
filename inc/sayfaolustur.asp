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
Response.charset = "UTF-8"
Session.codepage = 65001
MyMapPath = Server.MapPath(".")
MyMapPath = replace(MyMapPath,"\inc","\")
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.CreateTextFile(MyMapPath&"test.txt",true)
%>
<!--#include file="baglanti.inc"-->
<%
if request("ajax")="" then
%>
<HTML>

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<script>
		var url,mekan,intler
		mekan="";
		var xmlhttp;
		if (window.XMLHttpRequest)
			xmlhttp=new XMLHttpRequest();
		else
			xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		
		function grid_selected()
		{
			url = "sayfaolustur.asp?ajax=1&grid_columns="+document.getElementById("grid_columns").value+"&grid_table="+document.getElementById("grid_table").value;
			mekan = "grid";
			
			showHint();
		}
		
		function record_selected()
		{
			url = "sayfaolustur.asp?ajax=1&record_columns="+document.getElementById("record_columns").value+"&record_table="+document.getElementById("record_table").value;
			mekan = "record";
			showHint();
		}
		
		function intleri_ogren()
		{
			url = "sayfaolustur.asp?ajax=1&intler=1&record_columns="+document.getElementById("record_columns").value+"&record_table="+document.getElementById("record_table").value;
			mekan = "intler";
		showHint();
		}
		
		function saveTemplate(fileName){
			xmlhttp.open("GET","templatecreator.asp?PAGENAME="+fileName,false);
			xmlhttp.send(null);
			var content = encodeURIComponent(xmlhttp.responseText)
			data = "fname="+fileName.replace(".asp","")+"&content="+content;
			xmlhttp.open("POST","savefile.asp",false);
			xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
			xmlhttp.send(data);
			if (xmlhttp.responseText=="TRUE"){
				alert("HTML Şablon oluşturuldu!")
			}
			
		}	
		
		function showHint()
		{
			
				xmlhttp.onreadystatechange=function()
				{
				if(xmlhttp.readyState==4)
				  {
					if (mekan=="record")
					{
						intleri_ogren();
					}
					else
					{
						if (mekan=="intler")
						{
							intler = xmlhttp.responseText
							document.getElementById("record_int").innerHTML = intler;
						}
						else
						{
							if (mekan=="grid")
							{
								document.getElementById("grid_sutun").innerHTML=xmlhttp.responseText;
							}
							else
							{
								if (mekan!="")
								{
									document.getElementById(mekan+"_td_0").innerHTML = xmlhttp.responseText.replace("name='#","name='"+mekan+"_0");
									document.getElementById(mekan+"_td_1").innerHTML = xmlhttp.responseText.replace("name='#","name='"+mekan+"_1");
									document.getElementById(mekan+"_td_3").innerHTML = xmlhttp.responseText.replace("name='#","name='"+mekan+"_3") + "/<input type='text' name='"+mekan+"_4'>";
								}
							}
						}
					}
					
				  }
				}
			xmlhttp.open("GET",url,false);
			xmlhttp.send(null);
			
		}
		
		function tablodene(asd)
		{
			url = "sayfaolustur.asp?ajax=1&record_ajax=1&mytable="+document.getElementById(asd+"_2").value;
			mekan = asd;
			showHint();
		}
		
		function sutun_ekle()
		{
			var x=document.getElementById("sutuncu")
			value_sutuncu = x.options[x.selectedIndex].text
			if (document.getElementById("r_c").value.indexOf(value_sutuncu+",")==-1)
			{
				document.getElementById("r_c").value +=  value_sutuncu+","
			document.getElementById("ekleme_yeri").innerHTML  += "<br>"+value_sutuncu+" : "+sutunu_koy(value_sutuncu)
			}
		}
		
		function sutunu_koy(asd)
		{
			wturn =  "<select name='"+asd+"_AL' id='"+asd+"_AL_id'><option value=''>Ekleme</option><option value='createLookup(' selected>Combo</option><option value='createGenericLookup('>Generic Combo</option><option value='useAjax('>Ajax</option><option value='makeUploader('>Uploader</option></select><br>";
			wturn += "<table align='center' border='1'><tr><td>Sutunu Adlandir</td><td>ID Column</td><td>Text Column</td><td>Table/Yetkiler(Uploader)<input type='button' value='DENE' onclick=\"tablodene('"+asd+"')\"></td><td>Info(Ajax)/Filter(Combo)</td></tr><tr><td><input type='text' name='"+asd+"_99'></td>";
			
			for(var t=0;t<4;t++)
			{
				wturn += "<td id='"+asd+"_td_"+t+"'>";
				if  (t==2)
				{
					wturn += "<input type='text' value='-UP-DOWN-DEL-' id='"+asd+"_2' name='"+asd+"_2' style='width:100%;'>";
				}
				wturn += "</td>";
				
			}
			wturn +="</tr></table>";
			return wturn;
		}
		
		function fsubmit()
		{
			var requireds = ["grid_ad_id","record_ad_id","grid_columns","grid_table","record_columns","record_table","grid_baslik_id","record_baslik_id"];
			var res = true;
			var msg = "Please fill the following inputs:\n\n";
			for(x in requireds)
			{
				if (document.getElementById(requireds[x]).value=="")
				{
					msg += requireds[x]+"\n"
					res = false;
				}
			}
			if (!res)
				alert(msg);
			return res;
		}
	</script>

</head>
<body bgcolor="#efefef">
<form action="sayfaolustur.asp" method="post" onsubmit="return fsubmit();">
<table align="center" border="1" style="border: 1px;" width="90%">
	<tr><td colspan="2" align="center"><b>Sayfa Adları</b><hr>Grid Sayfasi: <input type="text" size="15" id="grid_ad_id" name="grid_ad" value="<%=request("grid_ad")%>">.asp , Record Sayfasi: <input type="text" size="15" name="record_ad" id="record_ad_id" value="<%=request("record_ad")%>">.asp<br>&nbsp;</td></tr>
	<tr><td colspan="2" align="center">Add to Menu <input type="checkbox" name="addtomenu" checked>, Menu Name: <input type="text" name="menu_name" value="menu.html"></td></tr>
	<tr align="center">
		<td><b>Grid Sayfası icin SQL</b><hR><b>SELECT</b><br><textarea name="grid_columns" style="width:100%" id="grid_columns"><%=request("grid_columns")%></textarea><br><b>FROM</b><br><textarea style="width:100%" id="grid_table" name="grid_table"><%=request("grid_table")%></textarea> <input type="button" onclick="grid_selected()" value="Devam"></td>
		<td><input type="hidden" id="r_c" value=""><b>Record Sayfası icin SQL</b><hr><b>SELECT</b> <input type="checkbox" name="id_cikar_rec" checked>(-ID)<BR><textarea style="width:100%" name="record_columns" id="record_columns"><%=request("record_columns")%></textarea><BR><b>FROM</b><BR><textarea style="width:100%" name="record_table" id="record_table"><%=request("record_table")%></textarea> <input type="button" onclick="record_selected()" value="Devam"></td>
	</tr>
</table>
<BR>
<HR>
<BR>
<table align="center" border="0" width="90%">
<tr>
<td valign="top" width="50%">
<table align="center" width="100%" border="1" style="border:1px;">
	<tr><td colspan="2">Grid Tablosu</td></tr>
	<tr><td>Listenin Basligi</td><td><input type="text" name="grid_baslik" id="grid_baslik_id" value="<%=request("grid_baslik")%>"></td></tr>
	<tr><td>Kişiselleştirilebilen Fieldlar</td><td><input type="text" name="grid_kisisel" value="*" value="<%=request("grid_kisisel")%>"></td></tr>
	<tr><td>Linki Verilecek Sutun</td><td id="grid_sutun">-</td></tr>
</table>
</td>
<td valign="top">
<table align="center" width="100%" border="1" style="border:1px;">
	<tr><td colspan="2">Record Tablosu</td></tr>
	<tr><td colspan="2">Çoğalt<input type="checkbox" name="C_COGALT" checked>&nbsp;&nbsp;&nbsp;Kaydet<input type="checkbox" name="C_KAYDET" checked>&nbsp;&nbsp;&nbsp;Geri Al<input type="checkbox" name="C_GERIAL" checked>&nbsp;&nbsp;&nbsp;Sil<input type="checkbox" name="C_SIL" checked>&nbsp;&nbsp;&nbsp;Geri Dön<input type="checkbox" name="C_GERIDON" checked></td>
	<tr><td>Listenin Basligi</td><td><input type="text" name="record_baslik" id="record_baslik_id" value="<%=request("record_baslik")%>"></td></tr>
	<tr><td>Dold. Zorunlu Fieldlar</td><td><input type="text" name="record_zorunlu" value="FIELD1 FIELD2 FIELD3"></td></tr>
	<tr><td>HTML Şablondan Oluştur</td><td><input type="checkbox" name="C_SABLONDAN" value="1" checked>Şablonadı: [RECORDSAYFASI].html olacaktır.</td></tr>
	<tr><td colspan="2"><table align="center" width="100%">
	<tr><td id="record_int">-</td><td><input type="button" value="Ekle" onclick="sutun_ekle()"></td></tr>
	</table></td></tr>
</table>
</td>
</tr>
<tr><td colspan="2" id="ekleme_yeri"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" name="bitti" value="Sayfalari Olustur"></td></tr>
</table>
</form>
</body>
</html>
<%
end if
	if (request("ajax")=1) then
		if request("grid_columns")<>"" then
			columns = request("grid_columns")
			table = request("grid_table")
			name = "grid"
		end if
		if request("record_columns")<>"" then
			columns = request("record_columns")
			table = request("record_table")
			name = "record"
		end if
		
		if request("record_ajax")<>"" then
			table = request("mytable")
			on error resume next
			sql = "SELECT TOP 1 * FROM "&table
			if instr(conn.Provider,"MSDASQL")>0 then 
				sql = "SELECT * FROM "&table&" LIMIT 0,1"
			end if
			Set RsF2 = Server.CreateObject("Adodb.Recordset")
			RsF2.open sql,conn
			if err<>0 then
				response.write "?"
				response.end
			end if
			response.write "<select name='#'><option value=''>Seciniz</value>"
			for each x in RsF2.fields
					response.write "<option value='"&x.name&"'>"&x.name&"</option>"
			next
			RsF2.Close
			response.write "</select>"
			response.end
		else
		
			Set RsF = Server.CreateObject("Adodb.Recordset")
			sql = "SELECT TOP 1 * FROM "&table
			if instr(conn.Provider,"MSDASQL")>0 then 
				sql = "SELECT * FROM "&table&" LIMIT 0,1"
			end if
			on error resume next
			RsF.open sql,conn
			if err>0 then
				response.write "Hatali Sql: "&err.description
			else
					if (request("intler")=1) then
						response.write "<select name='"&name&"' id='sutuncu'>"
					else
						response.write "<select name='"&name&"'>"
					end if
						for each x in RsF.fields
							if (x.type=3) then
								response.write "<option value='"&x.name&"'>"&x.name&"</option>"
							end if
						next
						for each x in RsF.fields
							if not (x.type=3) then
								response.write "<option value='"&x.name&"'>"&x.name&"</option>"
							end if
						next			
							response.write "</select>"
				'	else
				'		response.write "<select name='"&name&"'>"
				'		for each x in RsF.fields
				'			response.write "<option value='"&x.name&"'>"&x.name&"</option>"
				'		next			
				'		response.write "</select>"
				'	end if
			end if
		end if
	end if
	
	if (request("bitti")<>"") then
		if (request("grid_columns")<>"" and request("grid_table")<>"" and request("record_columns")<>"" and request("record_table")<>"" and request("grid")<>"" and request("record")<>"" and request("grid_baslik")<>"" and request("record_baslik")<>"" and request("grid_ad")<>"" and request("record_ad")<>"" and request("record_ad")<>request("grid_ad")) then
			grid_texts = ""
			'Grid.asp
			grid_texts = grid_texts & ("<!--#include file=""inc/grid_head.inc""-->") & vbcrlf
			grid_texts = grid_texts & ("<"&"%") & vbcrlf
			grid_texts = grid_texts & ("sub user(myFunction)") & vbcrlf
			grid_texts = grid_texts & ("	'myFunction: %function_name%_begin , %function_name%_end") & vbcrlf
			grid_texts = grid_texts & ("End sub") & vbcrlf
			grid_texts = grid_texts & ("'Sabitlerin Ayarlanmasi") & vbcrlf
			grid_texts = grid_texts & ("LIST_TITLE = """&request("grid_baslik")&"""	'Listenin Basligi") & vbcrlf
			grid_texts = grid_texts & ("TABLENAME = """&request("grid_table")&"""	'SELECT COLUMNS FROM ______") & vbcrlf
			grid_texts = grid_texts & ("COLUMNS = """&request("grid_columns")&"""	'SELECT _______ FROM TABLENAME") & vbcrlf
			grid_texts = grid_texts & ("RECORD_PAGE = """&request("record_ad")&".asp""	'Linki verilecek sayfa") & vbcrlf
			grid_texts = grid_texts & ("CHANGEABLE_COLUMNS = """&request("grid_kisisel")&"""	'Kolon ayarlarinda cikacak fieldlari yazin f1,f2") & vbcrlf
			grid_texts = grid_texts & ("LINKED_COLUMN = """&request("grid")&"""	'Linki verilecek sutun") & vbcrlf
			grid_texts = grid_texts & ("%"&">") & vbcrlf
			grid_texts = grid_texts & ("<!--#include file=""inc/grid_body.inc""-->") & vbcrlf
			if request("addtomenu")<>"" and request("menu_name")<>"" then
				grid_texts = grid_texts & ("<table align='center'><!--#include file="""&request("menu_name")&"""--></table>") & vbcrlf
			end if
			call saveFile(MyMapPath&request("grid_ad")&".asp",grid_texts)
			response.write request("grid_ad")&".asp Dosyasi olusturuldu > <a href='../"&request("grid_ad")&".asp'>Link</a><br>" 
			'Record.asp
				'Casesi bul
				Set RsF = Server.CreateObject("Adodb.Recordset")
				REQ_RECORDS_COLUMN = request("record_columns")
				sql = "SELECT TOP 1 "&REQ_RECORDS_COLUMN&" FROM "&request("record_table")
				if instr(conn.Provider,"MSDASQL")>0 then 
					sql = "SELECT "&REQ_RECORDS_COLUMN&" FROM "&request("record_table")&" LIMIT 0,1"&table
				end if
				CASES="|"
				RsF.open sql,conn
				if request("id_cikar_rec")<>"" then
					REQ_RECORDS_COLUMN = ""
					for each x in RsF.fields
						if not ucase(x.name) = ucase(request("grid")) then
							REQ_RECORDS_COLUMN = REQ_RECORDS_COLUMN &x.name&", "
						end if
					next
					REQ_RECORDS_COLUMN = Left(REQ_RECORDS_COLUMN,Len(REQ_RECORDS_COLUMN)-2)
				end if
				for each x in RsF.fields
					
					if request(x.name&"_99")<>"" then
						CASES= CASES &"["&x.name&"="&request(x.name&"_99")&"|"
					end if
					if request(x.name&"_AL")<>"" then
						if request(x.name&"_AL")="makeUploader(" then
							CASES= CASES &x.name&"|"&request(x.name&"_AL")&	"field,"""""&request(x.name&"_2")&""""")|"
						else
							CASES= CASES &x.name&"|"&request(x.name&"_AL")
							for k=0 to 3
								CASES = CASES & """"""&request(x.name&"_"&k)&""""","		
							next
							CASES = CASES & """"""&replace(request(x.name&"_4"),"""","""""""""")&""""",field)|"
						end if
					end if
				next
			
			record_texts = ""
			record_texts = record_texts & ("<!--#include file=""inc/record_head.inc""-->") & vbcrlf
			record_texts = record_texts & ("<"&"%") & vbcrlf
			record_texts = record_texts & ("'User Sub") & vbcrlf
			record_texts = record_texts & ("sub user(myFunction)") & vbcrlf
			record_texts = record_texts & ("	'myFunction: %function_name%_begin , %function_name%_end") & vbcrlf
			record_texts = record_texts & ("End sub") & vbcrlf
			record_texts = record_texts & ("") & vbcrlf
			record_texts = record_texts & ("'Sabitlerin Ayarlanmasi") & vbcrlf
			record_texts = record_texts & ("CASES = """&CASES&"""") & vbcrlf
			record_texts = record_texts & ("REQUIRED_COLUMNS= """&UCase(replace(replace(request("record_zorunlu"),","," "),"""",""))&"""	'Doldurulmasi zorunlu fieldlar") & vbcrlf
			if request("C_COGALT")="" then record_texts = record_texts & ("COGALT=0") & vbcrlf end if
			if request("C_KAYDET")="" then record_texts = record_texts & ("KAYDET=0") & vbcrlf end if
			if request("C_GERIAL")="" then record_texts = record_texts & ("GERIAL=0") & vbcrlf end if
			if request("C_SIL")="" then record_texts = record_texts & ("SIL=0") & vbcrlf end if
			if request("C_GERIDON")="" then record_texts = record_texts & ("GERIDON=0") & vbcrlf end if
			record_texts = record_texts & ("RECORD_TITLE = """&request("record_baslik")&"""	'Record Basligi") & vbcrlf
			record_texts = record_texts & ("TABLENAME = """&request("record_table")&"""	'SELECT COLUMNS FROM ______") & vbcrlf
			record_texts = record_texts & ("COLUMNS = """&REQ_RECORDS_COLUMN&"""	'SELECT _______ FROM TABLENAME") & vbcrlf
			if request("C_SABLONDAN")<>"" then
				record_texts = record_texts & ("HTML_TEMPLATE = """&request("record_ad")&".HTML""	'DEFAULT IS EMPTY") & vbcrlf
			end if
			record_texts = record_texts & ("LIST_PAGE = """&request("grid_ad")&".asp""	'Liste Sayfasi") & vbcrlf
			record_texts = record_texts & ("LINKED_COLUMN = """&request("grid")&"""	'Kayitlar hangi kolon adindan yollaniyor") & vbcrlf
			record_texts = record_texts & ("%"&">") & vbcrlf
			record_texts = record_texts & ("<!--#include file=""inc/record_body.inc""-->") & vbcrlf
			if request("addtomenu")<>"" and request("menu_name")<>"" then
				record_texts = record_texts & ("<table align='center'><!--#include file="""&request("menu_name")&"""--></table>") & vbcrlf
			end if
			call saveFile(MyMapPath&request("record_ad")&".asp",record_texts)
			response.write request("record_ad")&".asp Dosyasi olusturuldu > <a href='../"&request("record_ad")&".asp'>Link</a><br>" 
			
			if request("C_SABLONDAN")<>"" then
				response.write "<script>saveTemplate('"&request("record_ad")&".asp')</script>"
			end if

			if request("addtomenu")<>"" then
				if request("menu_name")<>"" then
					set fs=Server.CreateObject("Scripting.FileSystemObject")
					set f=fs.OpenTextFile(MyMapPath&request("menu_name"),8,true)
					mystr = request("grid_ad")
					f.writeline("<td><a href="""&mystr&".asp"">"&Ucase(Left(mystr,1))&lcase(Right(mystr,len(mystr)-1))&"</a></td>")
					set f=nothing
					set fs=nothing
					response.write "Menu has been written"
				else
					response.write "Menu name required to write menu"
				end if
				
			end if
			
		end if
		
		
	end if
	
	function saveFile(fileName,texts)
	on error resume next
		dim objStream
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Open
		objStream.CharSet = "UTF-8"
		objStream.WriteText(texts)
		objStream.SaveToFile fileName , 1
		objStream.Close
		if err<>0 then
			response.write fileName&"<br>"
			response.write err.description
			response.end
		end if
	end function
%>
