<%Response.Charset="gb2312"%>
<%
domain=request("domain")
querywords=request.QueryString("querywords")
'url="http://www.google.com.hk/search?hl=zh-CN&source=hp&biw=1276&bih=629&q="&querywords&"&aq=f&aqi=&aql=&oq=&num=100"
url="http://www.google.com.hk/search?hl=zh-CN&source=hp&q="&querywords&"&aq=f&num=100"
'response.write url&"<br>"
html=getHTTPPage(url,"utf-8")
'response.write html
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern="<ol((.|\n)*?)"&domain
	Set matches = reg.execute(html)
		For Each match1 in matches
			html=match1.Value
		Next
	'response.write html
Set matches = Nothing
Set reg = Nothing

if html<>"" then
	html=strCount(html,"��ҳ����")
	html=int(html)+1
else
	html=0
end if
if html>96 then html=0
%>
document.getElementById("paiming2").innerHTML = "<%=html%>";
<%
Set html= Nothing

'=================================������========================================

'ͳ��strA���ַ���,strB�������ַ�����
Function strCount(strA,strB)
 lngA = Len(strA)
 lngB = Len(strB)
 lngC = Len(Replace(strA,strB,""))
 strCount = (lngA - lngC) / lngB
End Function


'��ȡ�ַ���,1.����ǰ���ַ�����2.������ǰ���ַ���
Function strCut(strContent,StartStr,EndStr,CutType)
Dim S1,S2
On Error Resume Next
Select Case CutType
Case 1
  S1 = InStr(strContent,StartStr)
  S2 = InStr(S1,strContent,EndStr)+Len(EndStr)
Case 2
  S1 = InStr(strContent,StartStr)+Len(StartStr)
  S2 = InStr(S1,strContent,EndStr)
End Select
If Err Then
  strCute = "<p align='center' ><font size=-1>��ȡ�ַ�������.</font></p>"
  Err.Clear
  Exit Function
Else
  strCut = Mid(strContent,S1,S2-S1)
End If
End Function


Function getHTTPPage(Path,charset)
        t = GetBody(Path)
        getHTTPPage=BytesToBstr(t,charset)
End function

Function GetBody(url) 
        on error resume next
        'Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
        Set Retrieval = CreateObject("MSXML2.XMLHTTP") 
        With Retrieval 
        .Open "Get", url, False, "", "" 
        .Send 
        if Retrieval.readystate<>4 then 
			GetBody="0"
			exit function
        end if
        GetBody = .ResponseBody
        End With 
        Set Retrieval = Nothing 
End Function

Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function

%>