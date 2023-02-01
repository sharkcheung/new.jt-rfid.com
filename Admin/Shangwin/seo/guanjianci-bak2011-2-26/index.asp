<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
'Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
'Response.Expires=-999
Session.Timeout=999
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:xn="http://www.xiaonei.com/2009/xnml"  oncontextmenu="return false">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=7" />

<!--#include file=function.asp-->

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

%>

<%


if request("htmlpage")<>"" then  '输出详细页面


url="http://info.china.alibaba.com/news/"&request("htmlpage")
html=getHTTPPage(url)

if instr(html,"<!--左侧导航list 引入一个detail的导购库-->")>0 then
	'html=strCut(html,"<!--内容-->","<!--标准底部 start-->",1)
	response.write "<div class='ad'>提示：此内容来源于阿里巴巴的具体商家产品图，不适合SEO收录，已屏蔽显示！如要继续请点<a href='"&url&"'>直接进入>></a></div>"
else
	html=strCut(html,"<!--content start-->","<!--pagelist end-->",1)
		if html="" then
			response.write "<div class='ad'>提示：此内容被系统视为广告内容，已屏蔽显示！请关闭本页。</div>"
		else
			html=strCut(html,"<h1>","<!--pagelist end-->",1)
			html=replace(html,"<h1>","<div class=page><h1>")
			html=replace(html,"</h1>","</h1></div>")
			html=replace(html,"/news/detail/","?htmlpage=detail/")
			html=replace(html,"http://info.china.alibaba.com","")
			html=replace(html,"#newsdetail-content","")
			response.write html
		end if
end if

else    '输出列表页面
url=yurl&words

html=getHTTPPage(url)
html=strCut(html,"<!--相关检索词分布-->","<!--页脚开始-->",1)

html=replace(html,"../","http://index.baidu.com/")
html=replace(html,"人群属性分布","")
html=replace(html,"相关检索词","关键词联想，搜索最多的前10")
'html=replace(html,"./word.php?type=0&word=","?words=")
html=replace(html,"./word.php?type=0&word=","#")
html=replace(html,"./word.php?","#")
html=replace(html,"当月","#")

response.write "<title>与【"&words&"】—SEO关键词联想</title>"
%>
<link href="baidu.css" rel="stylesheet" type="text/css">
<script type="text/javascript"> 
	//suggestion数据服务url
	//var sugSvr = 'http://jx-ba-test04.jx.baidu.com:8100/';
	var sugSvr   =  'http://nssug.baidu.com/';
</script>
<script type="text/javascript" src="http://index.baidu.com/script/core.js"></script>
<script type="text/javascript" src="http://index.baidu.com/script/region.js"></script>
<script type="text/javascript" src="http://index.baidu.com/script/swf.js"></script>
<script type="text/javascript" src="http://index.baidu.com/script/tangram.js"> </script>
<link href="column.css" rel="stylesheet" type="text/css" />

</head><body oncontextmenu="return false">
<div id="main">

	<table border="1" width="99%" id="table1" style="border-collapse: collapse" bordercolor="#AAC7E9">
<tr>
			<td class="td0">
<form method="POST" action="index.asp">
输入要联想的关键词：<input type="text" name="kw" size="20" value="<%=words%>"><input type="submit" value="关键词联想" name="B1">
</form>
</td></tr></table>
<%
response.write "<div class=top>您使用过的联想关键词："&Request.Cookies("guanjiancikeywords")&"　[<a href='?clear=yes'>清空</a>]</div>"
response.write html
response.write "</div>"

end if

%>