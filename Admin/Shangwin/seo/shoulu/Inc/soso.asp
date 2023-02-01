<!--#include file="fun.asp"--><%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="sosos" Then
SosoUrl="http://www.soso.com/q?w="&wd&"&sc=web&ch=w.ptl&lr=chs"
Else
SosoUrl="http://www.soso.com/q?w=site%3A"&wd&"&sc=web&ch=w.ptl&lr=chs"
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
SosoWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(SosoUrl,"gb2312")
dim sosoWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="搜索到约((.|\n)*?)项结果"
Set matches = reg.execute(TempStr)
For Each match1 in matches
sosoWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing

SosoWebSite=Replace(SosoWebSite,"搜索到约","")
SosoWebSite=Replace(SosoWebSite,"项结果","")
SosoWebSite=Replace(SosoWebSite," ","")
SosoWebSite=Replace(SosoWebSite,",","")
SosoWebSite=Replace(SosoWebSite," ","")

If SosoWebSite="" Then SosoWebSite="0"

	Response.Cookies(cook)=SosoWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if %>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=sosourl%>><%=sosoWebSite%></a>';