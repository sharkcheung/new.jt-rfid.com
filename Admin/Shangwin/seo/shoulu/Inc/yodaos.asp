<!--#include file="fun.asp"-->
<%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="yodaos" Then
YoudaoUrl="http://www.youdao.com/search?q=domain%3A"&wd
Else
YoudaoUrl="http://www.youdao.com/search?q=site%3A"&wd
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
YoudaoWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(YoudaoUrl,"utf-8")
dim YoudaoWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="��((.|\n)*?)�����"
Set matches = reg.execute(TempStr)
For Each match1 in matches
YoudaoWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing
YoudaoWebSite=Replace(YoudaoWebSite,"��","")
YoudaoWebSite=Replace(YoudaoWebSite,"�����","")
YoudaoWebSite=Replace(YoudaoWebSite,"Լ","")
YoudaoWebSite=Replace(YoudaoWebSite,",","")
YoudaoWebSite=Replace(YoudaoWebSite,"","")

If YoudaoWebSite="" Then YoudaoWebSite="0"
	Response.Cookies(cook)=YoudaoWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if
%>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=Youdaourl%>><%=YoudaoWebSite%></a>';