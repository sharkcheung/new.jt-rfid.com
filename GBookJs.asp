<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：GBookJs.asp
'文件用途：留言框JS
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Id=Clng(Request.QueryString("Id"))
%>window.onresize = baiduResizeDiv;
window.onerror = function(){}
var divTop,divLeft,divWidth,divHeight,docHeight,docWidth,objTimer,i = 0;
var px = document.doctype?"px":0;
var scrollwidth = navigator.userAgent.indexOf("Firefox")>0?16:0;
var iframeheight = navigator.userAgent.indexOf("MSIE")>0?170-2:209-2;
String.prototype.Trim  = function(){return this.replace(/^\s+|\s+$/g,"");}
function baidu_collapse(obj){
	ct = document.getElementById('tab_c_iframe');
	if(ct.style.display=="none"){
		ct.style.display="";
		obj.src=obj.src.replace("b.gif","a.gif");
	} else {
		ct.style.display="none";
		obj.src=obj.src.replace("a.gif","b.gif");
	}
	baiduResizeDiv();
}

function baiduMsg()
{
	try{
		divTop = parseInt(document.getElementById("eMeng").style.top,10);
		divLeft = parseInt(document.getElementById("eMeng").style.left,10);
		divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10);
		divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10);

		var scrollPosTop,scrollPosLeft,docWidth,docHeight;
		if (typeof window.pageYOffset != 'undefined') { 
			scrollPosTop = window.pageYOffset; 
			scrollPosLeft = window.pageXOffset; 
			docWidth = window.innerWidth; 
			docHeight = window.innerHeight; 
		} else if (typeof document.compatMode != 'undefined' && document.compatMode != 'BackCompat') {
			scrollPosTop = document.documentElement.scrollTop;
			scrollPosLeft = document.documentElement.scrollLeft;
			docWidth = document.documentElement.clientWidth;
			docHeight = document.documentElement.clientHeight;
		} else if (typeof document.body != 'undefined') { 
			scrollPosTop = document.body.scrollTop;
			scrollPosLeft = document.body.scrollLeft;
			docWidth = document.body.clientWidth;
			docHeight = document.body.clientHeight;
		}

		document.getElementById("eMeng").style.top = parseInt(scrollPosTop,10) + docHeight + 10 + px;// divHeight
		document.getElementById("eMeng").style.left = parseInt(scrollPosLeft,10) + docWidth - divWidth - scrollwidth + px;
		document.getElementById("eMeng").style.visibility="visible";
		objTimer = window.setInterval("baidu_move_div()",10);
	}catch(e){}
}

function baiduResizeDiv()
{
	i+=1;
	try{
		divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10);
		divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10);

		var scrollPosTop,scrollPosLeft,docWidth,docHeight; 
		if (typeof window.pageYOffset != 'undefined') { 
			scrollPosTop = window.pageYOffset; 
			scrollPosLeft = window.pageXOffset; 
			docWidth = window.innerWidth; 
			docHeight = window.innerHeight; 
		} else if (typeof document.compatMode != 'undefined' && document.compatMode != 'BackCompat') {
			scrollPosTop = document.documentElement.scrollTop;
			scrollPosLeft = document.documentElement.scrollLeft;
			docWidth = document.documentElement.clientWidth;
			docHeight = document.documentElement.clientHeight;
		} else if (typeof document.body != 'undefined') { 
			scrollPosTop = document.body.scrollTop;
			scrollPosLeft = document.body.scrollLeft;
			docWidth = document.body.clientWidth;
			docHeight = document.body.clientHeight;
		}

		document.getElementById("eMeng").style.top = docHeight - divHeight + parseInt(scrollPosTop,10) + px;
		document.getElementById("eMeng").style.left = docWidth - divWidth + parseInt(scrollPosLeft,10) - scrollwidth + px;
	}catch(e){}
}

function baidu_move_div()
{

	var scrollPosTop,scrollPosLeft,docWidth,docHeight; 
	if (typeof window.pageYOffset != 'undefined') { 
		scrollPosTop = window.pageYOffset; 
		scrollPosLeft = window.pageXOffset; 
		docWidth = window.innerWidth; 
		docHeight = window.innerHeight; 
	} else if (typeof document.compatMode != 'undefined' && document.compatMode != 'BackCompat') {
		scrollPosTop = document.documentElement.scrollTop;
		scrollPosLeft = document.documentElement.scrollLeft;
		docWidth = document.documentElement.clientWidth;
		docHeight = document.documentElement.clientHeight;
	} else if (typeof document.body != 'undefined') { 
		scrollPosTop = document.body.scrollTop;
		scrollPosLeft = document.body.scrollLeft;
		docWidth = document.body.clientWidth;
		docHeight = document.body.clientHeight;
	}

	try{
		if(parseInt(document.getElementById("eMeng").style.top,10) <= (docHeight - divHeight + parseInt(scrollPosTop,10)))
		{
			window.clearInterval(objTimer);objTimer = window.setInterval("baiduResizeDiv()",1);
		}
		divTop = parseInt(document.getElementById("eMeng").style.top,10);
		document.getElementById("eMeng").style.top = divTop - 1 + px;
	}catch(e){}
}
function baiduMessbox(shape,color){
	var styles='position:absolute;top:0px;left:0px;z-index:99999;visibility:hidden;';
	var copic='<img src="<%=SiteDir%>Images/ico_'+color+'a.gif" align="absmiddle" class="ioc'+color+'" onClick="baidu_collapse(this)">';
	if(shape>1) window.onload = baiduMsg;
	else {styles='';copic=''}
	var s;
	if(shape==2) {
		s='<div id=eMeng style="Z-INDEX:99999;LEFT:0px;POSITION:absolute;TOP:0px;VISIBILITY:hidden;">'
			+ '<table width="216" border="0" cellpadding="0" background="<%=SiteDir%>Images/texttop.gif" cellspacing="0" class="" id="tab_'+(color+3)+'" style="border:0px;">'
			+'<tr>'
			+'<td align="right"></td>'
			+'</tr>'
			+'  <tr>'
			+'    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">'
			+'        <tr>'
			+'          <td width="21" rowspan="2" valign="bottom"></td>'
			+'        </tr>'
			+'        <tr>'
			+'         <td align=right style="font-size:12px; line-height:21px;height:21px;padding-right:6px;color:#FFFFFF;" onDblClick="baidu_collapse(document.all.baidu_Tu)"><img src="<%=SiteDir%>Images/ico_'+color+'a.gif" align="absmiddle" class="ioc'+color+'" id="baidu_Tu" onClick="baidu_collapse(this)">'
			+'		  </td>'
			+'        </tr>'
			+'      </table>'
			+'	  </td>'
			+'    </tr>'
			+ '</table>'
			+ '<iframe src="<%=SiteDir%>GBookFrame.asp?Id=<%=Id%>" width="216" height="' + iframeheight + '" frameborder="0" id="tab_c_iframe"></iframe>'
			+'</div>';
	}
	document.writeln(s);
}
baiduMessbox(2,1);

<!--#Include File="Code.asp"-->
