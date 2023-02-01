<!--#include file="fun.asp"-->
<%
'远程截取函数开始
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"

dim wd
wd=Request("d")
'截取网址
url="http://www.zzsky.cn/tool/sogourank/?q="&wd&""

cook=Request("d")&"sogoudj"
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
gourank=Request.Cookies(cook)
else

        wstr=getHTTPPage(url,"utf-8")
'截取数据
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern=">> 搜狗评级：<font color=red>((.|\n)*?)</font>"
	Set matches = reg.execute(wstr)
		For Each match1 in matches
			gourank=match1.Value
		Next
Set matches = Nothing
Set reg = Nothing
gourank=Replace(gourank,">> 搜狗评级：<font color=red>","")
gourank=Replace(gourank,"</font>","")
gourank=Replace(gourank,"&nbsp;","")
gourank=Replace(gourank,",","")
gourank=Replace(gourank," ","")
If gourank = "" Then
gourank = "无数据"
End If

	Response.Cookies(cook)=gourank
	Response.Cookies(cook).Expires=dateadd("h",72,now())
end if
%>
document.getElementById("gourank").innerHTML = "<%=gourank%>";