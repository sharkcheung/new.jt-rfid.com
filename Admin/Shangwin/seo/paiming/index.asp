<!--#include file=function.asp-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:xn="http://www.xiaonei.com/2009/xnml"  oncontextmenu="return false">
<head>
<link href="column.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#Include File="../../loginchk.asp"-->
<!-- #include file=../../inc.asp -->
</head>
<body oncontextmenu="return false">
<%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
host=lcase(request.servervariables("HTTP_HOST")) 

Dim domain,Url,s,ty,Url1,strPage,StrPage1
Dim xmldom,SD,SITE,dimg

domain = request("url")
querywords= request("querywords")
jifang= request("jifang")


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
'domain = "qebang.cn"
'domain=host
End if
if request("domain")="localhost" then
domain = "qebang.cn"
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

domain=replace(domain,"www.","")
%>
<html>
<head>
<title><%=domain%>站长工具,域名工具,收录查询工具</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css" />
<base target="_blank">
</head>
<body oncontextmenu="return false">
<div id="main">
	<table border="1" width="99%" id="table1" style="border-collapse: collapse" bordercolor="#AAC7E9">
		<tr>
			<td class="td0"><img src="images/td0.gif" style="display:none;" />
			搜索引擎关键词排名查询:</td>
		</tr>
		<tr>
			<td class="td2"> <form action="index.asp" method="get" target="_self">
 关键词排名查询：http://
 <input name="url" type="text" style="width:180px" value="<%=domain%>">　关键词：<input id="querywords" class="inputtext" name="querywords" size="20" value="<%=querywords%>">  <input type="submit" name="sub" class="submit" value="排名查询"></form></td>
		</tr>
		<%if request("url")<>"" and request("querywords")<>"" then%>
		<tr>
			<td class="td2" style="height:50px;">在百度中排名第 <span id="paiming"><img src="/admin/images/loading.gif"  /></span>位，在Google中排名第 <span id="paiming2"><img src="/admin/images/loading.gif"  /></span>位</td>		
		</tr>	
		<%end if%>
	</table>
</div>
<%if request("url")<>"" and request("querywords")<>"" then%>
<script language="javascript" type="text/javascript" src="p-baidu.asp?domain=<%=request("url")%>&querywords=<%=request("querywords")%>"></script>
<script language="javascript" type="text/javascript" src="p-google.asp?domain=<%=request("url")%>&querywords=<%=request("querywords")%>"></script>
<%end if%>
</body>
</html>