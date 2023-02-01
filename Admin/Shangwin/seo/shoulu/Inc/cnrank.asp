<!--#include file="fun.asp"-->
<%
on error resume next
'远程截取函数开始
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"

Dim Domain,Url,R
Domain=Request("d")

url="http://www.chinarank.org.cn/overview/Info.do?url="&domain

cook=Request("d")&"chinarank"
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
cnrank=Request.Cookies(cook)
else

        TempStr=getHTTPPage(url,"gb2312")
dim cnrank
set reg=new Regexp
reg.Multiline=True
reg.Global=Flase
reg.IgnoreCase=true
reg.Pattern="<span class=""bold"">当前排名：</span><span class=""rank_font_y2"">((.|\n)*?)</span><br />"
Set matches = reg.execute(TempStr)
For Each match1 in matches
cnrank=match1.Value
Next
Set matches = Nothing
Set reg = Nothing
cnrank=Replace(cnRank,"<span class=""bold"">当前排名：</span><span class=""rank_font_y2"">","")
cnrank=Replace(cnRank," ","")
cnrank=Replace(cnRank,"</span><br />","")
cnrank=cint(cnrank)
If cnrank = "" Then
cnrank = "未查到"
End If

	Response.Cookies(cook)=cnrank
	Response.Cookies(cook).Expires=dateadd("h",72,now())
end if
%>document.getElementById("cnrank").innerHTML = "<%=cnrank%>";