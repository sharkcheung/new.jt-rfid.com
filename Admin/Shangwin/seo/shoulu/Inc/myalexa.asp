<!--#include file="fun.asp"-->
<%
on error resume next
'远程截取函数开始
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"
Function del(str)
    str=replace(str,"<REACH RANK=""","")
    str=replace(str,"""/>","")
    str=replace(str," ","")
del=str
End Function

dim wd
wd=Request("d")
'截取网址
url="http://data.alexa.com/data/?cli=10&dat=snba&ver=7.0&url="&wd

cook=Request("d")&"alexa"
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
myalexa1=Request.Cookies(cook)
else

        wstr=getHTTPPage(url,"gb2312")
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern="<REACH RANK=((.|\n)*?)>"
	Set matches = reg.execute(wstr)
	For Each match1 in matches			
	myalexa1=del(match1.Value)
		Next
Set matches = Nothing
Set reg = Nothing

myalexa1=cint(trim(myalexa1))
If myalexa1 = "" Then
myalexa1 = "无数据"
End If

	Response.Cookies(cook)=myalexa1
	Response.Cookies(cook).Expires=dateadd("h",72,now())
end if
%>document.getElementById("myalexa").innerHTML = "<%=myalexa1%>";
