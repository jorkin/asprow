﻿<!--#include file="inc/grid_head.inc"-->
<%
	Dim g_tablename	'Table Name
	Dim g_IDColumn	'ID column
	Dim g_colsWOID	'Columns without ID field
	Dim g_listTitle	'Title of the Page
	
	g_listTitle = request.querystring("gnrTitle")
	g_tablename = trim(request.querystring("gnrTABLENAME"))
	if g_tablename="" then
		response.write "You must send gnrTABLENAME parameter in order to display the page.."
		response.end
	end if
	g_tablename = replace(g_tablename,"'","''")
	
	on error resume next
	Rs.open "SELECT TOP 1 * FROM "&g_tablename,conn
	if err<>0 then
		response.write "Error occured : "&err.description
		response.end
	else
		g_IDColumn = getID(g_tablename)
	end if
	Rs.Close
	
	
	g_colsWOID = request.querystring("gnrCOLUMNS")
	if g_colsWOID="" then
		g_colsWOID = "*"
	end if
	
	if g_listTitle="" then
		g_listTitle = "Generic Grid Page for "&g_tablename
	end if
%>
		
		<%
		sub user(myFunction)
		End sub
		LIST_TITLE = g_listTitle
		TABLENAME = g_tablename
		COLUMNS = g_colsWOID
		GENERICPAGE= 1
		RECORD_PAGE = "genericRecord.asp?gnrTABLENAME="&g_tablename
		CHANGEABLE_COLUMNS = " "
		LINKED_COLUMN = g_IDColumn
		%>
		<!--#include file="inc/grid_body.inc"-->