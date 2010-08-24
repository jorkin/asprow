var DBTYPE = "ACCESS";
var access_types = new Array("TEXT(5)","TEXT(20)","TEXT(50)","TEXT(100)","TEXT(250)","MEMO","BYTE","INTEGER","LONG","COUNTER","SINGLE","DOUBLE","CURRENCY","GUID","DATETIME","YESNO","LONGBINARY")
function createtable()
{
	var table_name = document.getElementById("tablename").value;
	var all_elements = document.getElementById("create_form").elements;
	var colums = [];
	sayac = -1;
	for(var i = 0;i<all_elements.length;i++)
	{
		if ((i%3)==0)
		{
			colums.push({name:"",type:"",pk:false});
			sayac++;
		}
		if ((i%3)==0)
			colums[sayac].pk = (all_elements[i].checked);
		if ((i%3)==1)
			colums[sayac].name = (all_elements[i].value);
		if ((i%3)==2)
			colums[sayac].type = (all_elements[i].value);
	}
	var query = "";
	if (DBTYPE=="ACCESS")
	{
		query = "CREATE TABLE "+table_name+" ("
		for(var i = 0;i<colums.length;i++)
		{
			query += colums[i].name+" "+colums[i].type
			if (colums[i].pk)
				query +=  ", Constraint "+table_name+"_PK Primary Key ("+colums[i].name+")"
				
			if ((i+1)!=colums.length)
				query +=  ",";
		}
		query +=  ");";
		alert(query);
		return query;
	}
}
function tableTemplate()
{
	var html = "<table id='create_table' align='center' width='400'>"
	html += "<tr><td>"
	html += "DB: <select onchange='{DBTYPE=this.value;}' id='DBTYPE'><option value='ACCESS'>ACCESS</option><option value='ACCESS'>ACCESS</option></select>"
	html += "</td>"
	html += "<td>"
	html += "TABLE NAME: <input type='text' id='tablename'>"
	html += "</td>"
	html += "<td>"
	html += "<input type='button' onclick='addcolumn()' value='Add Column'>"
	html += "</td></tr>"
	html += "<tr><td colspan='2'><form id='create_form'><table width='100%' cellspacing='0' cellpadding='0' id='add_row_here'></table></form></td></tr>"
	html += "<tr><td colspan='3' align='center'>"
	html += "<input type='button' onclick='createtable()' value='Create Table'>"
	html += "</td></tr>"
	html += "</table>"
	document.write(html);
}
var col_count = 0;
function addcolumn()
{
	var x=document.getElementById('add_row_here').insertRow(col_count);
	var y=x.insertCell(0);
	col_count++;
	var ihtml = "<table width='100%' cellspacing='0' cellpadding='0'><tr>"
	ihtml += "<td><input type='checkbox'></td><td><input type='text'></td>"
	ihtml += "<td><select>"
	
	if (DBTYPE=="ACCESS")
	{
		ihtml += "<option value=''>Seciniz</option>"
		for (var i=0;i<access_types.length;i++)
			ihtml += "<option value='"+access_types[i]+"'>"+access_types[i]+"</option>"
	}
	ihtml += "</select></td>"
	ihtml += "</tr></table>";
	y.innerHTML = ihtml;
}