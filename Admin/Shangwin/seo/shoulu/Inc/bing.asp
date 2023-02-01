<!--#include file="fun.asp"-->
<%
function getNum(str) 
 dim re 
 set re=new RegExp 
 re.pattern="\D" 
 re.global=true 
 getNum = re.replace(str, "") 
 end function

response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
wd=Request("d")
If Request("s")="bings" Then
bingUrl="http://cn.bing.com/search?q=link%3A"&wd&"&form=QBLH&filt=all"
Else
bingUrl="http://cn.bing.com/search?q=site%3A"&wd&"&form=QBLH&filt=all"
End If

cook=wd&Request("s")
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
bingWebSite=Request.Cookies(cook)
else

TempStr= getHTTPPage(bingUrl,"UTF-8")
dim bingWebSite
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="<span class=""sb_count"" id=""count"">((.|\n)*?)</span><span class=""sc_bullet"">"
Set matches = reg.execute(TempStr)
For Each match1 in matches
bingWebSite=match1.Value
Next
Set matches = Nothing
Set reg = Nothing
start=Newstring(bingWebSite,"共")
over=Newstring(bingWebSite,"条")
bingWebSite=mid(bingWebSite,start,over)
bingWebSite=Replace(bingWebSite,"共","")
bingWebSite=Replace(bingWebSite,"条","")
bingWebSite=getNum(bingWebSite)

If bingWebSite="" Then bingWebSite="0" 
	Response.Cookies(cook)=bingWebSite
	Response.Cookies(cook).Expires=dateadd("h",24,now())
end if
%>
document.getElementById('<%=Request("s")%>').innerHTML = '<a href=<%=bingurl%>><%=bingWebSite%></a>';