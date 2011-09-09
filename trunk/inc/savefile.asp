<%
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

fname = request("fname")
content = request("content")


if (fname<>"" and content<>"") then
	MyMapPath = Server.MapPath(".")
	MyMapPath = replace(MyMapPath,"\inc","\")
	call saveFile(MyMapPath&fname&".html",content)
	response.write "TRUE"
end if

%>