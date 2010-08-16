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
<!--#include file="upload.inc"-->
<%
Dim uploadCamePage
Dim recordid
Dim extName
Dim fieldName

set o = new clsUpload

if request.querystring("pagefrom")<>"" then
	'response.write "Debug: Querystring found!<br>"
	uploadCamePage = request.querystring("pagefrom")
	recordid = request.querystring("rid")
	extName = request.querystring("ext")
	fieldName = request.querystring("fname")
else
	'response.write "Debug: Querystring not found!<br>"
	uploadCamePage = o.ValueOf("pagefrom")
	recordid = o.ValueOf("rid")
	extName = o.ValueOf("ext")
	fieldName = o.ValueOf("fname")
end if
extName = replace(extName,".","")
extName = replace(extName," ","")
if recordid="NEWID" then
	response.write "First, create the record for uploading file.."
	response.end
end if
Dim ndocname

Dim fs,f,dadi,satir
Set fs=Server.CreateObject("Scripting.FileSystemObject")
dadi = replace(Server.MapPath("."),"\inc","\"&uploadCamePage&".asp")
if fs.FileExists(dadi)=true then
	'response.write "Debug: dadi found<br>"
	Set f=fs.OpenTextFile(dadi, 1)
	found_r = false
	found_t = false
	col_satir = ""
	table_satir = ""
	do while f.AtEndOfStream = false
		satir = f.ReadLine
		if (instr(satir,"LINKED_COLUMN =")>0 or instr(satir,"LINKED_COLUMN=")>0) then
			found_r = true
			col_satir = satir
		end if
		if (instr(satir,"TABLENAME =")>0 or instr(satir,"TABLENAME=")>0) then
			found_t = true
			table_satir = satir
		end if
		if found_t and found_r then
			exit do
		end if
	loop
	if found_t and found_r then
		'response.write "Debug: found_t,found_r success<br>"
		tab_f1 = instr(table_satir,"""")+1
		tab_f2 = instr(tab_f1,table_satir,"""")
		mytablename = Mid(table_satir,tab_f1,tab_f2-tab_f1)
		
		tab_f1 = instr(col_satir,"""")+1
		tab_f2 = instr(tab_f1,col_satir,"""")
		mylcol = Mid(col_satir,tab_f1,tab_f2-tab_f1)
		
		ndocname = mytablename&"_"&fieldName&"_"&recordid&"_"&extName
		ndocname_raw = mytablename&"_"&fieldName&"_"&recordid&"_"
	else
		'response.write "Debug: found_t,found_r not success: found_t:"&found_t&" found_r:"&found_r&" <br>"
		ndocname = ""
	end if
	f.Close
	Set f=Nothing
else
	'response.write "Debug: dadi not found: "&dadi&"<br>"
	ndocname = ""
end if

'response.write "Debug: ndocname = "&ndocname&"<br>"
'response.write "Debug: uploadCamePage = "&uploadCamePage&" recordid = "&recordid&" extName = "&extName&"<br>"

if ndocname<>"" then
	fdizin = replace(Server.MapPath("."),"\inc","\docs\")
	on error resume next
	set fo=fs.GetFolder(fdizin)
	if err<>0 then
		response.write err.description&"<br>"&fdizin
		resonse.end
	end if
	for each x in fo.files
		if ucase(Left(x.Name,Len(ndocname&"."))) = Ucase(ndocname&".") then
			myFileName = x.Name
			exit for
		end if
		if ucase(Left(x.Name,Len(ndocname_raw)))=ucase(ndocname_raw) then
			set f=fs.GetFile(fdizin&x.Name)
			myFileName = ndocname&"."&right(f.name,len(f.name)-instrrev(f.name,"."))
			f.name = myFileName
			exit for
		end if
	next
	'response.write "Debug: uploadCamePage = "&uploadCamePage&" recordid = "&recordid&" extName = "&extName&"<br>"

	if fs.FileExists(fdizin&myFileName)=true then
		if request("act")="delete" then
			if instr(Session(uploadCamePage&"_"&fieldName&"_"&recordid),"DEL")>0 then
				fs.DeleteFile(fdizin&myFileName)
			end if
		end if
	end if
	if fs.FileExists(fdizin&myFileName)=true then
		
		if instr(Session(uploadCamePage&"_"&fieldName&"_"&recordid),"DOWN")>0 then
			response.write "<a href=""download.asp?rid="&recordid&"&path="&myFileName&""" target=""_blank"">Download</a>"
		else
			response.write "Download"
		end if
		
		if instr(Session(uploadCamePage&"_"&fieldName&"_"&recordid),"DEL")>0 then
			response.write " / <a href=""upload.asp?ext="&extName&"&rid="&recordid&"&fname="&fieldName&"&pagefrom="&uploadCamePage&"&act=delete&docname="&ndocname&""">ReUpload</a>"
		else
			response.write " / ReUpload"
		end if
		response.end
	end if

	if instr(Session(uploadCamePage&"_"&fieldName&"_"&recordid),"UP")>0 then
		if o.Exists("cmdSubmit") then
			if o.FileNameOf("txtFile")<>"" then
				
				if fs.FileExists(fdizin&myFileName)=true then
					fs.DeleteFile(fdizin&myFileName)
				end if
				
				'get client file name without path
				sFileSplit = split(o.FileNameOf("txtFile"), "\")
				sFile = sFileSplit(Ubound(sFileSplit))
				uzanti = Right(sFile,Len(sFile)-instr(sFile,"."))
				if lcase(uzanti)<>"php" and lcase(uzanti)<>"asp" and lcase(uzanti)<>"jsp" and lcase(uzanti)<>"swf" then
					o.FileInputName = "txtFile"
					o.FileFullPath = Server.MapPath(".") & "\..\docs\" & ndocname&"."&uzanti
					o.save
				else
					response.write "Script files are not allowed to upload!"
				end if
				if o.Error = "" then
					response.write "File saved !"
					response.end
				else
					response.write "Failed due to the following error: " & o.Error
				end if
			end if
		end if
	else
		response.write "You are not authorized to upload file!<br>"
		response.end
	end if
end if
set o = nothing
%>
<FORM ACTION = "upload.asp" ENCTYPE="multipart/form-data" METHOD="POST">
<INPUT NAME = "pagefrom" type="hidden" value="<%=uploadCamePage%>"></INPUT>
<INPUT NAME = "fname" type="hidden" value="<%=fieldName%>"></INPUT>
<INPUT NAME = "rid" type="hidden" value="<%=recordid%>"></INPUT>
<INPUT NAME = "ext" type="hidden" style="width:50px;" value="<%=extName%>"></INPUT>
<INPUT TYPE=FILE NAME="txtFile"><INPUT TYPE = "SUBMIT" NAME="cmdSubmit" VALUE="Upload">
</FORM>