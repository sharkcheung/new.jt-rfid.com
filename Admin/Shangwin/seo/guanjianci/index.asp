<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
'Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
'Response.Expires=-999
Session.Timeout=999
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html oncontextmenu="return false" xmlns="http://www.w3.org/1999/xhtml" xmlns:xn="http://www.xiaonei.com/2009/xnml">
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<meta content="IE=7" http-equiv="X-UA-Compatible" />
<!-- ymPrompt组件 -->
<script type="text/javascript" src="/admin/winskin/ymPrompt.js"></script>
<link rel="stylesheet" type="text/css" href="/admin/winskin/qq/ymPrompt.css" /> 
<!-- ymPrompt组件 -->
<!--#Include File="../../loginchk.asp"-->
<!--include file=function.asp-->
<!-- include file=../../inc.asp -->
<%
'参数设置------------
yurl="http://index.baidu.com/main/word.php?word="
if request("kw")="" then
words=request("words")
else
words=request("kw")
end if
addwords="<a href='?words="&words&"'>"&words&"</a> "

if words<>"" then
	if instr(Request.Cookies("guanjiancikeywords"),addwords)>0 then
	else
		Response.Cookies("guanjiancikeywords")=addwords&Request.Cookies("guanjiancikeywords")
		Response.Cookies("guanjiancikeywords").Expires=#May 10,2020#
	end if
end if
if request("clear")<>"" then
Response.Cookies("guanjiancikeywords")=""
end if

'参数设置------------

response.write "<title>与【"&words&"】—SEO关键词联想</title>"
%>
<link href="column.css" rel="stylesheet" type="text/css" />
</head>
<body oncontextmenu="return false">
<div id="main">
	<table id="table1" border="1" bordercolor="#AAC7E9" style="border-collapse: collapse" width="99%">
		<tr>
			<td class="td0">
			<form action="index.asp" method="POST">
				输入进行联想的关键词（越短越好）：<input name="kw" size="20" type="text" value="<%=words%>"><input name="B1" type="submit" value="关键词联想"> <span class="tinfo"><b>提示：</b>单击列表中的关键词即可自动添加到关键词库</span>
			</form>
			</td>
		</tr>
	</table>
<%
response.write "<div class=top>您使用过的联想关键词："&Request.Cookies("guanjiancikeywords")&"　[<a href='?clear=yes'>清空</a>]</div>"
response.write html
response.write "</div>"
if request.form("kw")="" then
	kw=request("words")
else
	kw=request.form("kw")
end if
%>
<div class="div">
<iframe id="keywordsframe" marginwidth="0" marginheight="0" src="kwtest/?d=<%=kw%>" frameborder="0" width="100%" scrolling="no" height="25" onload="this.height=this.contentWindow.document.body.scrollHeight" name="I2"></iframe>
</div>
</div>
<div id="loading" style="display:;position:absolute; top:30%; left:49%; z-index:10000; ">
	 <img alt="" src="/admin/Images/loading.gif"></div>
<script language="javascript"> 
<!-- 
 
var frame = document.getElementById("keywordsframe"); 
frame.onreadystatechange = function(){ 
if( this.readyState == "complete" ) 
document.getElementById("loading").style.display="none";
} 
 
function loadingview(){
document.getElementById("loading").style.display="block";
}
//--> 
</script>
</body>
</html>
