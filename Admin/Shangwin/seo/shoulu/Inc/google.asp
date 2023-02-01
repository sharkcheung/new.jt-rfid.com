<!--#include file="fun.asp"-->
<%
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="googles" Then
  GoogleUrl="http://www.google.com.hk/search?complete=1&hl=zh-CN&q=link%3A"&wd
Else
  GoogleUrl="http://www.google.com.hk/search?complete=1&hl=zh-CN&q=site%3A"&wd
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
GoogleWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(GoogleUrl,"UTF-8")
dim GoogleWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="找到约((.|\n)*?)条结果"
Set matches = reg.execute(TempStr)
For Each match1 in matches
GoogleWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing
GoogleWebSite=Replace(GoogleWebSite,"<b>","")
GoogleWebSite=Replace(GoogleWebSite,"</b>","")
GoogleWebSite=Replace(GoogleWebSite,"条结果","")
GoogleWebSite=Replace(GoogleWebSite,"找到","")
GoogleWebSite=Replace(GoogleWebSite,"约","")
GoogleWebSite=Replace(GoogleWebSite,",","")
GoogleWebSite=Replace(GoogleWebSite," ","")

If GoogleWebSite="" Then GoogleWebSite="0"

	Response.Cookies(cook)=GoogleWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if
%>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=googleurl%>><%=googleWebSite%></a>';