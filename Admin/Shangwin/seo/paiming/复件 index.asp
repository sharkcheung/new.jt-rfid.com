<!--#include file=function.asp-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:xn="http://www.xiaonei.com/2009/xnml"  oncontextmenu="return false">
<head>
<link href="column.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
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
<div class="div" style="display:">
 <form action="index.asp" method="get" target="_self">
 关键词排名查询：http://
 <input name="url" type="text" style="width:180px" value="<%=domain%>">　关键词：<input id="querywords" class="inputtext" name="querywords" size="20" value="<%=querywords%>">　区域：<select name="jifang">
			<option value="广东东莞电信2">广东东莞电信2</option>
<option value="上海电信2">上海电信2</option>
	<option value="北京1">北京1</option>
	<option value="北京4">北京4</option>
	<option value="北京电信2">北京电信2</option>
	<option value="安徽蚌埠电信1">安徽蚌埠电信1</option>
	<option value="安徽合肥电信3">安徽合肥电信3</option>
	<option value="安徽合肥电信4">安徽合肥电信4</option>
	<option value="安徽芜湖电信1">安徽芜湖电信1</option>
	<option value="安徽芜湖电信2">安徽芜湖电信2</option>
	<option value="福建厦门电信1">福建厦门电信1</option>
	<option value="福建厦门电信3">福建厦门电信3</option>
	<option value="河南郑州网通">河南郑州网通</option>
	<option value="湖北鄂州电信1">湖北鄂州电信1</option>
	<option value="湖北潜江电信1">湖北潜江电信1</option>
	<option value="湖北十堰电信1">湖北十堰电信1</option>
	<option value="湖北仙桃电信1">湖北仙桃电信1</option>
	<option value="湖北咸宁电信1">湖北咸宁电信1</option>
	<option value="湖南长沙电信1">湖南长沙电信1</option>
	<option value="吉林省吉林市联通1">吉林省吉林市联通1</option>
	<option value="江苏淮安电信1">江苏淮安电信1</option>
	<option value="江苏南京电信1">江苏南京电信1</option>
	<option value="江苏无锡电信1">江苏无锡电信1</option>
	<option value="江苏无锡铁通1">江苏无锡铁通1</option>
	<option value="江苏扬州电信1">江苏扬州电信1</option>
	<option value="江苏镇江电信1">江苏镇江电信1</option>
	<option value="江苏镇江电信2">江苏镇江电信2</option>
	<option value="江西南昌电信2">江西南昌电信2</option>
	<option value="辽宁省沈阳联通1">辽宁省沈阳联通1</option>
	<option value="美国2">美国2</option>
	<option value="山西太原联通1">山西太原联通1</option>
	<option value="上海武胜电信1">上海武胜电信1</option>
	<option value="四川电信1">四川电信1</option>
	<option value="四川绵阳电信1">四川绵阳电信1</option>
	<option value="浙江杭州电信1">浙江杭州电信1</option>
	<option value="浙江金华电信1">浙江金华电信1</option>
	<option value="浙江绍兴机房1">浙江绍兴机房1</option>
	<option value="浙江温州电信1">浙江温州电信1</option>
	<option value="浙江温州电信2">浙江温州电信2</option>
	</select>
  <input type="submit" name="sub" class="submit" value="查询"></form></div>


	<table border="1" width="99%" id="table1" style="border-collapse: collapse" bordercolor="#AAC7E9">
		<tr>
			<td class="td0"><img src="images/td0.gif" style="display:none;" />
			搜索引擎关键词排名结果:</td>
		</tr>
		<tr>
			<td class="td2">
			<span id="paiming">
			
<%If domain<>"" Then
wd=domain
querywords=request("querywords")
jifang=request("jifang")
BaiduUrl="http://i.linkhelper.cn/paiming.asp?url="&wd&"&querywords="&querywords&"&jifang="&jifang&""
html=getHTTPPage(BaiduUrl)
html=strCut(html,"排名如下：","</td>",2)
html=RemoveHTML(html)
html=replace(html,"&nbsp;&nbsp;&nbsp;点击查看详情>>","")
html=replace(html,"关键字为","关键词")
html=replace(html,"，地区机房为","　区域为")
html=replace(html,"，百度、google排名如下：","　排名如下：<br>")
html=replace(html,"Google排名","<br>Google排名")
html=replace(html,"查询失败，点击查看详情>>","")


if instr(html,"Transitional")>0 then html="请完整输入站点域名、关键词、区域三项内容"
If html="" Then html="0"
response.write html&"<br><b>[ 因google服务器处在香港的原因，谷歌排名查询准确率低于百度 ]</b>"
end if%>

			</span>

			</td>
		</tr>
		
		
	</table>
</div>
</body>
</html>