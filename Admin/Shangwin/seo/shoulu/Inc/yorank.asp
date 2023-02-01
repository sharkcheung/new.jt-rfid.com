<!--#include file="fun.asp"--><%
on error resume next
'远程截取函数开始
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"

Dim Domain,Url,R
Domain=Request("d")

url="http://toolbarapi-www.youdao.com/api?req=rank&keyfrom=toolbar.2.20.0011.5000&vendor=yodao&p=http%3A%2F%2F"&domain
        wstr=getHTTPPage(url,"gb2312")
        start=Newstring(wstr,"<rank>")
        over=Newstring(wstr,"</rank>")
Rank=mid(wstr,start,over-start)
yorank=Replace(Rank,"<rank>","")
yorank=getNum(yorank)
yorank=Replace(Rank,"</rank>","")
yorank=Replace(Rank,"-1","0")
If yorank = "" Then
yorank = "0"
End If
%>document.getElementById("yorank").innerHTML = "<%=yorank%>";