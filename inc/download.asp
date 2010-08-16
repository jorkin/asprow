<%
	path = request.querystring("path")
	rid = request.querystring("rid")
	if IsNumeric(rid) and rid<>"" then
		response.redirect "../docs/"&path
	else
		response.write "Incorrect Path.."
	end if
%>