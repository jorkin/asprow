<!--#include file="baglanti.inc"-->
<!--#include file="function_grid.inc"-->
<%
'Sabitlerin Tanimi
Dim COLUMN_count
Dim LIST_TITLE
Dim TABLENAME
Dim COLUMNS
Dim RECORD_PAGE
Dim TOTAL_RECORD_COUNT
Dim RECORD_COUNT
Dim LINKED_COLUMN
Dim LINK_FORMAT
Dim PAGE_COUNT
Dim LESS_COUNT
Dim THIS_PAGE
Dim REQUEST_ORDER
Dim REQUEST_ORDERBY
Dim REQUEST_EXPORT_STYLE
Dim CASES
Dim GENERICPAGE
Dim JAVASCRIPT
Dim TARIHKUTULARI
TARIHKUTULARI = ""
JAVASCRIPT = ""
REQUEST_EXPORT_STYLE = REQUEST("EXPORT_STYLE")
if REQUEST_EXPORT_STYLE="" then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8">
		<title>Liste</title>
		<link rel="stylesheet" type="text/css" href="Glassy/Style_doctype.css">
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
end if
%>