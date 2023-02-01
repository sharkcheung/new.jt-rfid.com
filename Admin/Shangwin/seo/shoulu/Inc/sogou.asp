<!--#include file="fun.asp"--><%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="sogous" Then
SoGouUrl="http://www.sogou.com/web?query="&wd
Else
SoGouUrl="http://www.sogou.com/web?query=site%3A"&wd
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
SoGouWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(SoGouUrl,"gb2312")
dim SoGouWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="’“µΩ ((.|\n)*?) <!"
Set matches = reg.execute(TempStr)
For Each match1 in matches
SoGouWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing

SoGouWebSite=Replace(SoGouWebSite,"’“µΩ ","")
SoGouWebSite=Replace(SoGouWebSite,"<!","")
SoGouWebSite=Replace(SoGouWebSite," ","")
SoGouWebSite=Replace(SoGouWebSite,",","")
SoGouWebSite=Replace(SoGouWebSite,"","")

If SoGouWebSite="" Then SoGouWebSite="0"
	Response.Cookies(cook)=SoGouWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if%>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=sogouurl%>><%=sogouWebSite%></a>';