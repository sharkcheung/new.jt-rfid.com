<!--#include file="fun.asp"-->
<%
'Զ�̽�ȡ������ʼ
response.expires = -1
response.addheader "cache-control","no-cache"
Response.AddHeader "Pragma","no-cache"

dim wd
wd=Request("d")
'��ȡ��ַ
url="http://www.zzsky.cn/tool/sogourank/?q="&wd&""

cook=Request("d")&"sogoudj"
cook=replace(cook,".","")
if Request.Cookies(cook)<>"" then
gourank=Request.Cookies(cook)
else

        wstr=getHTTPPage(url,"utf-8")
'��ȡ����
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern=">> �ѹ�������<font color=red>((.|\n)*?)</font>"
	Set matches = reg.execute(wstr)
		For Each match1 in matches
			gourank=match1.Value
		Next
Set matches = Nothing
Set reg = Nothing
gourank=Replace(gourank,">> �ѹ�������<font color=red>","")
gourank=Replace(gourank,"</font>","")
gourank=Replace(gourank,"&nbsp;","")
gourank=Replace(gourank,",","")
gourank=Replace(gourank," ","")
If gourank = "" Then
gourank = "������"
End If

	Response.Cookies(cook)=gourank
	Response.Cookies(cook).Expires=dateadd("h",72,now())
end if
%>
document.getElementById("gourank").innerHTML = "<%=gourank%>";