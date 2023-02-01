<%
 
'-----------------函数区--------------------------------------------
 
 
'输出分页 
sub pagelist(shuzi)
pageview="<a href='?pp=0'>1</a>"
for i=1 to shuzi
ii=i*15
ppp= "<a href='?pp="&ii&"'>"&i+1&"</a>"
pageview= pageview&" "&ppp
next
response.write pageview
end sub
 
 
 
  '获取页面的HTML
 
function getHTTPPage(url) 
dim Http
set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then 
exit function
end if
getHTTPPage=bytesToBSTR(Http.responseBody,"gb2312")
'getHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")

getHTTPPage=replace(getHTTPPage,"DoNews","企帮")  '批量替换donews
set http=nothing
if err.number<>0 then err.Clear 
end function

Function BytesToBstr(body,Cset)  '转换中文字符乱码
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

%>
