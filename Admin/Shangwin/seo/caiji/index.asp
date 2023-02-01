<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:xn="http://www.xiaonei.com/2009/xnml"  oncontextmenu="return false">
<!--#include file=function.asp-->
<!-- include file=../../inc.asp -->
<%
'参数设置------------
yurl="http://info.china.alibaba.com/news/searchrss.htm?keywords="
if request("kw")="" then
keywords=request("keywords")
else
keywords=request("kw")
end if
add="<a href='?keywords="&keywords&"'>"&keywords&"</a> "

if keywords<>"" then
	if instr(Request.Cookies("caijiwords"),add)>0 then
	else
		Response.Cookies("caijiwords")=add&Request.Cookies("caijiwords")
		Response.Cookies("caijiwords").Expires=#May 10,2020#
	end if
end if
if request("clear")<>"" then
Response.Cookies("caijiwords")=""
end if
Response.Cookies("keywords")=""

'参数设置------------

%>

<%

pp=request("pp")
if pp="" or pp="-15" then pp="0"
ppp=pp+15
pppp=pp-15


if request("htmlpage")<>"" then  '输出详细页面

%>
<div class="div">
<%
url="http://info.china.alibaba.com/news/"&request("htmlpage")
html=getHTTPPage(url)

if instr(html,"<!--左侧导航list 引入一个detail的导购库-->")>0 then
	'html=strCut(html,"<!--内容-->","<!--标准底部 start-->",1)
	response.write "<div class='ad'>提示：此内容来源于阿里巴巴的具体商家产品图，不适合SEO收录，已屏蔽显示！如要继续请点<a href='"&url&"' target='_blank'>进入>></a>　<div class='ntop'><a href='javascript:window.history.go(-1);'>返回</a></div></div>"
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
			html=replace(html,"http://jiage.china.alibaba.com/html/js/swfobject.js","")
			html=replace(html,"http://style.china.alibaba.com/js/fdevlib/core/yui/get-min.js","")
			response.write "<div class='ntop'><a onclick='javascript:window.history.go(-1);'><< 返 回 </a>"
			response.write "<form action='function.asp?act=1' target='I1' method='post'><textarea style='display:none' rows='2' name='S1' cols='13'>"&html&"</textarea>"
				newstitle=strCut(html,"<h1>","</h1>",2)
			response.write "<input style='display:none' type='text' name='T1' value='"&newstitle&"' size='50'>"
			response.write "　<input class='ipt' type='submit' value='一键采集到网站指定栏目！' name='B1'>"
			response.write Request.Cookies("newsclass")
			response.write "</form></div>"
			response.write html
			response.write "<iframe name='I1' src='' marginwidth='1' marginheight='1' height='1' width='1'></iframe>"
		end if
end if
%>
</div></body>
<%
else    '输出列表页面
url=yurl&keywords

html=getHTTPPage(url)
html=strCut(html,"</webmaster>","</channel>",1)
html=replace(html,"</webmaster>","<div>")
html=replace(html,"</channel>","</div>")
html=replace(html,"<![CDATA[","")
html=replace(html,"]]>","")
html=replace(html,"pubdate","span")
html=replace(html,"<span>","<hr><span class=span>")
html=replace(html,"description","p")
html=replace(html,"<category>关键字订阅</category>","<div class='line'></div>")
html=replace(html,"<author>http://info.china.alibaba.com</author>","")
html=replace(html,"<item>","<li>")
html=replace(html,"</item>","</li>")
html=replace(html,"<title>","<b>")
html=replace(html,"</title>","</b>")
html=replace(html,"<link>","<a>")
html=replace(html,"</link>","</a>")
html=replace(html,"<div>","<div class=list>")
html=replace(html,"<a>","<a href='")
html=replace(html,"</a>","' target='_self'>详细>></a>")
html=replace(html,keywords,"<font color='#FF0000'>"&keywords&"</font>")
html=replace(html,"?tracelog=info_rss","")
html=replace(html,"http://info.china.alibaba.com/news/","?htmlpage=")
html=replace(html,"　","")

'html=html&" "&shuzi+1&"页 15条/页　　<span class='sxye'><a href='?pp="&pppp&"'>上一页</a><a href='?pp="&ppp&"'>下一页</a></span>"

response.write "<title>与【"&keywords&"】相关的资讯—SEO关键词资讯采集</title></head><body oncontextmenu='return false'>"
%>
<div class="td0">
<form method="POST" action="index.asp">
输入要采集信息的关键词：<input type="text" name="kw" size="20"><input type="submit" value=" 搜 一 下 ! " name="B1">
</form>
</div>
<%response.write "<div class=top>您使用过的关键词："&Request.Cookies("caijiwords")&"　[<a href='?clear=yes'>清空</a>]</div>"%>
<div class="div">
<%
 if keywords<>"" then
response.write html
end if
'response.write "</div><div class='fenye'>"

'call pagelist(shuzi)
response.write "</div>"

'response.write "</body></html>"

end if

%>