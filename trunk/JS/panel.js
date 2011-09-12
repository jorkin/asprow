var menuCollapsed = false;
var pageNo = 0;
var pages = [];
var menuPage = "";
function geid(id){
	return document.getElementById(id);
}
function collapse(){
	if (menuCollapsed){
		geid("leftMenuTd").style.width="208px";
		geid("closedPanelDiv").className = "closedPanel pasif";
		geid("menuTable").className = "";
	}else{
		geid("leftMenuTd").style.width="20px";
		geid("closedPanelDiv").className = "closedPanel";
		geid("menuTable").className = "pasif";
	}
	menuCollapsed = !menuCollapsed;
}
function openPage(pageId){
	var pageExist = false;
	for(var i=0;i<pages.length;i++)
		if (("f_"+pages[i][0])==pageId)
			pageExist = true;
	if (pageExist){
		for(var i=0;i<pages.length;i++){
			geid("f_"+pages[i][0]).style.display = "none";
			var cn = geid("d_"+pages[i][0]).childNodes;
			for(var j=0;j<cn.length;j++){
				cn[j].className = cn[j].className.replace("Active","");
			}
		}
		var cnn = geid("d_"+pageId.replace("f_","")).childNodes;
		for(var i=0;i<cnn.length;i++)
			if (cnn[i].className!="")
				if (cnn[i].className.indexOf("tab")==0)
					cnn[i].className = "Active"+cnn[i].className.replace("Active","");
		geid(pageId).style.display = "";
	}
}

function createPage(pageUrl,tabName){
	for(var i=0;i<pages.length;i++){
		if (pages[i][1]==pageUrl && pages[i][2]==tabName){
			openPage("f_"+pages[i][0]);
			return;
		}
	}
	var pageId = "f_"+pageNo;
	var dPageId = "d_"+pageNo;
	var buttonId = "b_"+pageNo;
	var contentHTML = "<iframe style=\"display:none;\" id=\""+pageId+"\" src=\""+pageUrl+"\" width=\"100%\" height=\"100%\" frameborder=\"0\"></iframe>";
	var buttonHTML = "<div id=\""+dPageId+"\" onclick=\"openPage('"+pageId+"')\" class=\"tabElement\"><div class='tabLeft'></div><div class='tabMiddle' id=\""+buttonId+"\">"+tabName+"</div><div class='tabRight'></div><div class='closeButton'onclick=\"closePage('"+dPageId+"')\"></div></div>"
	geid("pageContent").innerHTML += contentHTML;
	geid("menubuttons").innerHTML += buttonHTML;
	pages.push([pageNo,pageUrl,tabName]);
	pageNo++;
	openPage(pageId);
}

function closePage(dPageId){
	geid("menubuttons").removeChild(geid(dPageId));
	geid("pageContent").removeChild(geid(dPageId.replace("d_","f_")));
	var pageNo = parseInt(dPageId.replace("d_",""));
	for(var i=0;i<pages.length;i++){
		if (pages[i][0]==pageNo){
			pages.splice(i,1);
			return;
		}
	}
}

function getMenu(){
	if (menuPage!=""){
		var xmlhttp;
		if (window.XMLHttpRequest)
			xmlhttp=new XMLHttpRequest();
		else
			xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		xmlhttp.onreadystatechange=function()
		{
			if (xmlhttp.readyState==4 && xmlhttp.status==200)
			{
				geid("treeMenuContent").innerHTML=this.responseText;
				$("#tree").explr();
			}
		}
		xmlhttp.open("GET",menuPage,true);
		xmlhttp.send(null);
	}
}