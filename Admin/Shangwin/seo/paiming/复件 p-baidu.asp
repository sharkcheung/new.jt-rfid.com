<%Response.Charset="gb2312"%>
<%
domain=request("domain")
querywords=request.QueryString("querywords")
url="http://www.baidu.com/s?wd="&querywords&"&rn=100"
html=getHTTPPage(url,"gb2312")
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern="<table cellpadding((.|\n)*?)"&domain
	Set matches = reg.execute(html)
		For Each match1 in matches
			html=match1.Value
		Next
Set matches = Nothing
Set reg = Nothing
html=strCount(html,"<table cellpadding")
if html="101" then html="0"
%>
document.getElementById("paiming").innerHTML = "<%=html%>";
<%
Set html= Nothing

'=================================函数区========================================

'统计strA：字符串,strB：查找字符个数
Function strCount(strA,strB)
 lngA = Len(strA)
 lngB = Len(strB)
 lngC = Len(Replace(strA,strB,""))
 strCount = (lngA - lngC) / lngB
End Function


'截取字符串,1.包括前后字符串，2.不包括前后字符串
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
  strCute = "<p align='center' ><font size=-1>截取字符串出错.</font></p>"
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