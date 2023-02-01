<!--#include file="fun.asp"--><%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="baidus" Then
BaiduUrl="http://www.baidu.com/s?wd=domain%3A"&wd
Else
BaiduUrl="http://www.baidu.com/s?wd=site%3A"&wd
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
baiduWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(BaiduUrl,"gb2312")
dim BaiduWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="找到相关结果((.|\n)*?)个"
Set matches = reg.execute(TempStr)
For Each match1 in matches
BaiduWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing
BaiduWebSite=Replace(BaiduWebSite,"找到相关结果","")
BaiduWebSite=Replace(BaiduWebSite,"个","")
BaiduWebSite=Replace(BaiduWebSite,"约","")
BaiduWebSite=Replace(BaiduWebSite,",","")
BaiduWebSite=Replace(BaiduWebSite," ","")
If baiduWebSite="" Then baiduWebSite="0"

	Response.Cookies(cook)=baiduWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if
%>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=baiduurl%>><%=baiduWebSite%></a>';