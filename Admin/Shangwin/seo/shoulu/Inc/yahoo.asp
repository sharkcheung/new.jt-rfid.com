<!--#include file="fun.asp"--><%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="yahoos" Then
YahooUrl="http://sitemap.cn.yahoo.com/search?p="&wd&"&bwm=i"
Else
YahooUrl="http://sitemap.cn.yahoo.com/search?p="&wd&"&bwm=p"
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
YahooWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(YahooUrl,"utf-8")
dim YahooWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="�� <strong>((.|\n)*?)</strong> ��"
Set matches = reg.execute(TempStr)
For Each match1 in matches
YahooWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing

YahooWebSite=Replace(YahooWebSite,"�� <strong>","")
YahooWebSite=Replace(YahooWebSite,"</strong> ��","")
YahooWebSite=Replace(YahooWebSite,",","")
YahooWebSite=Replace(YahooWebSite," ","")

If YahooWebSite="" Then YahooWebSite="0" 

	Response.Cookies(cook)=YahooWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if %>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=yahoourl%>><%=yahooWebSite%></a>';