<!--#include file="fun.asp"-->
<%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<iframe src="login.html" width="100" height="100"></iframe>
<%
BaiduUrl="http://ci.aizhan.com/%E8%90%A5%E9%94%80/"
TempStr= getHTTPPage(BaiduUrl,"utf-8")

dim BaiduWebSite
	set reg=new Regexp
		reg.Multiline=True
		reg.Global=Flase
		reg.IgnoreCase=true
		reg.Pattern="</div>-->((.|\n)*?)最近查询"
	Set matches = reg.execute(TempStr)
	For Each match1 in matches
		BaiduWebSite=match1.Value
	Next
	Set matches = Nothing
	Set reg = Nothing
BaiduWebSite=Replace(BaiduWebSite,"找到相关结果","")
BaiduWebSite=Replace(BaiduWebSite,"个","")
BaiduWebSite=Replace(BaiduWebSite,"约","")
BaiduWebSite=Replace(BaiduWebSite,"数","")
BaiduWebSite=Replace(BaiduWebSite,",","")
BaiduWebSite=Replace(BaiduWebSite," ","")
response.write BaiduWebSite
%>
</body>
</html>

