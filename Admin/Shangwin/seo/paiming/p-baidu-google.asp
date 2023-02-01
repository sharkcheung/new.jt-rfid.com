<%Response.Charset="gb2312"%>
<%
domain=request("domain")
querywords=request.QueryString("querywords")
'response.write domain&querywords
call chapaimingjieguo()
%>
<%
sub chapaimingjieguo()
'百度
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
if html<>"" then
	html=strCount(html,"<table cellpadding")
	html=int(html)
else
	html=0
end if
if html>97 then html=0
	if html<>0 then
		rshtml="Baidu:<b>"&html&"</b> "
	else
		rshtml="Baidu:"&html&" "
	end if
Set html= Nothing
'谷歌
url="http://www.google.com.hk/search?hl=zh-CN&source=hp&q="&querywords&"&aq=f&num=100"
html=getHTTPPage(url,"utf-8")
set reg=new Regexp
	reg.Multiline=True
	reg.Global=Flase
	reg.IgnoreCase=true
	reg.Pattern="<ol((.|\n)*?)"&domain
	Set matches = reg.execute(html)
		For Each match1 in matches
			html=match1.Value
		Next
Set matches = Nothing
Set reg = Nothing

if html<>"" then
	html=strCount(html,"网页快照")
	html=int(html)+1
else
	html=0
end if
if html>95 then html=0
	if html<>0 then
		rshtml=rshtml&"Google:<b>"&html&"</b>"
	else
		rshtml=rshtml&"Google:"&html&""
	end if
response.write rshtml
Set html= Nothing
Set rshtml= Nothing
end sub
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
  strCute = "0010"
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