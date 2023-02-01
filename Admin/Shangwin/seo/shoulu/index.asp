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
<html xmlns="http://www.w3.org/1999/xhtml"  oncontextmenu="return false">
<!--#Include File="../../loginchk.asp"-->
<!-- #include file=../../inc.asp -->
<%
host=lcase(request.servervariables("HTTP_HOST")) 

Dim domain,Url,s,ty,Url1,strPage,StrPage1
Dim xmldom,SD,SITE,dimg

domain = request("url")

if left(domain,7)="http://" then
domain=right(domain,len(domain)-7)
end if
if instr(domain,"/")<>0 then
domain=left(domain,instr(domain,"/")-1)
end if

aa = Split(domain,".") 
jj = ubound(aa)
if jj>=2 then
domain =""& aa(jj-2) & "." & aa(jj-1) & "." & aa(jj) &""
else
domain =""& domain &""
end if

domain=LCase(domain)

if Not iswww(domain) Then
domain = "qebang.cn"
'domain=host
End if
if request("domain")="localhost" then
domain = "qebang.com"
End if

on error resume Next
Function iswww(strng)
 iswww = false
 Dim regEx, Match
 Set regEx = New RegExp
 regEx.Pattern = "^\w+((-\w+)|(\.\w+))*[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$" 
 regEx.IgnoreCase = True
 Set Match = regEx.Execute(strng)
 if match.count then iswww= true
End Function

set tnames = request.cookies("dnames")
if isnull(tnames) or len(trim(tnames))=0 then
tnames = domain&"|"
else
if instr(tnames,domain)>0 then
names = replace(tnames,domain&"|","")
else
tnames = domain&"|"&tnames
end if
end If

ttnames = split(tnames,"|")
tmpncontent = ""

if ubound(ttnames)>5 then
for tat=0 to 4
tmpncontent = tmpncontent&ttnames(tat)&"|"
next
else
tmpncontent=tnames
end If


%>
<html>
<head>
<title><%=domain%>站长工具,域名工具,收录查询工具<%response.write Request.Cookies("qebang.cn")("baidu")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css" />
<base target="_blank">
</head>
<body oncontextmenu="return false">
<div id="main">
<%
DIM myArray()
REDIM myArray(30)
myArray(1) = "baidu"
myArray(2)= "baidus"
myArray(3)= "google"
myArray(4)= "googles"
myArray(5)= "soso"
myArray(6)= "sosos"
myArray(7)= "sogou"
myArray(8)= "sogous"
myArray(9)= "yahoo"
myArray(10)= "yahoos"
myArray(11)= "yodao"
myArray(12)= "yodaos"
myArray(13)= "bing"
myArray(14)= "bings"

myArray(15)= "gourank"
myArray(16)= "mypr"
myArray(17)= "myalexa"
myArray(18)= "cnrank"


%>
	<table border="1" width="99%" id="table1" style="border-collapse: collapse" bordercolor="#AAC7E9">
		<tr>
			<td class="td0" colspan="8"><img src="images/td0.gif" style="display:none;" />
			搜索引擎SEO收录结果查询:　<span><%=domain%></span></td>
		</tr>
		<tr>
			<td class="td2" colspan="8">
			Sogou评级：<span id="gourank"><img src="/admin/Images/loading.gif" /></span>
			谷歌PR值：<span id="mypr"><img src="/admin/Images/loading.gif" /></span>
			全球综合排名：<span id="myalexa"><img src="/admin/Images/loading.gif" /></span>
			国内综合排名：<span id="cnrank"><img src="/admin/Images/loading.gif" /></span>
			</td>
		</tr>
		<tr>
			<td class="td1">搜索引擎</td>
			<td class="td1"><img src="images/baidu.gif" /> 
			百度</td>
			<td class="td1"><img src="images/google.gif" /> 
			谷歌</td>
			<td class="td1"><img src="images/soso.gif" /> 搜搜</td>
			<td class="td1"><img src="images/sogou.gif" /> 
			搜狗</td>
			<td class="td1"><img src="images/yahoo.gif" /> 
			雅虎</td>
			<td class="td1"><img src="images/youdao.gif" /> 
			有道</td>
			<td class="td1"><img src="images/bing.gif" /> 
			Bing</td>
		</tr>
		<tr>
			<td class="td1">收录数量</td>
			<% 
			i=1
			while i<14  %>
			<td>
			<div id="<%=myArray(i)%>" class="shuzi"><img src="/admin/Images/loading.gif"  /></div>
			<%
			i=i+2
			wend%>
		</tr>
		<tr>
			<td class="td1">反向链接</td>
			<% 
			i=2
			while i<15  %>
			<td>
			<div id="<%=myArray(i)%>" class="shuzi"><img src="/admin/Images/loading.gif"  /></div>
			<%
			i=i+2
			wend%>
		</tr>
		<tr>
			<td class="td1">收录提交</td>
			<td>
			<a href="http://www.baidu.com/search/url_submit.html">提交入口</a></td>
			<td>
			<a href="http://www.google.com/addurl/?continue=/addurl">提交入口</a></td>
			<td>
			<a href="http://www.soso.com/help/usb/urlsubmit.shtml">提交入口</a></td>
			<td>
			<a href="http://www.sogou.com/feedback/urlfeedback.php">提交入口</a></td>
			<td><a href="http://search.help.cn.yahoo.com/h4_4.html">提交入口</a></td>
			<td><a href="http://tellbot.youdao.com/report">提交入口</a></td>
			<td>
			<a href="http://cn.bing.com/webmaster/SubmitSitePage.aspx?mkt=zh-CN">提交入口</a></td>
		</tr>

		
	</table>
<div class="div">
<h4>搜索引擎收录?</h4>
<p>搜索引擎收录是搜索引擎收录一个网站页面具体的数量值，收录的数量越多，收录的时间越快，证明此网站对搜索引擎比较友好！</p>
<h4>反向链接?</h4>
<p>也叫外部链接，指的是其它站点上指向你站点的链接,对方网站知名度越高、访问量越大，对有利于快速提升你网站知名度和排名。</p>
</div>
<div class="div">
 <form action="" method="get" target="_self">
 查询其它站点的SEO收录情况：http://
 <input name="url" type="text" style="width:180px" value="<%=domain%>">
  <input type="submit" name="sub" class="submit" value="查询" onclick="checkDomain()"></form></div>

</div>

<%If domain<>"" Then%>
<%
 i=1
 while i<19  
 %>
<script language="javascript" type="text/javascript" src="inc/<%=myArray(i)%>.asp?d=<%=domain%>&s=<%=myArray(i)%>"></script>
<%
i=i+1
wend
%>
<%end if%>
</body>
</html>