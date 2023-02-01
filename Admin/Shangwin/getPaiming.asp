<%@language="vbscript" codepage="65001"%> 
<!--#include file="easp.asp"-->
<%dim XmlHttpData,posturl,u,k,r,bdrank,ggrank
Server.ScriptTimeout=999999
response.Charset="utf-8"
session.CodePage=65001
u=asp.r("d",0)
'u="qebang.cn"
k=encodeUrl(asp.r("k",0),"65001","936")
Function encodeUrl(paraString,Encoding1,Encoding2)
    '程序使用的编码 utf-8=65001,GB2312=936
    'Encoding1   原编码   Encoding2 为要输出的编码方式
    Session.CodePage=Encoding2
    encodeUrl = server.urlencode(paraString)
    Session.CodePage=Encoding1
End Function
posturl="http://i.linkhelper.cn/searchkeywords.asp"
XmlHttpData="url="&u&"&querywords="&k&"&btnsubmit=%B2%E9+%D1%AF"
r=PostHttpPage(posturl,"gb2312",XmlHttpData,posturl)
'asp.we r
bdrank=trim(GetBody(r,"百度排名：","&nbsp;",false,false))
ggrank=trim(GetBody(r,"Google排名：","&nbsp;",false,false))
'r=GetBody(r,"第","个",false,false)
if bdrank="$False$" Or bdrank="排名在200名之外" then
	asp.w "Baidu:0 "
Else
	If bdrank>100 Then
		asp.w "Baidu:0 "
	else
		asp.w "Baidu:<b>"&bdrank&"</b> "
	End if
end If
if ggrank="$False$" Or ggrank="排名在200名之外" then
	asp.w "Google:0 "
else
	If ggrank>100 Then
		asp.w "Google:0"
	else
		asp.w "Google:<b>"&ggrank&"</b>"
	End if
end If

'远程获取
Function PostHttpPage(PostUrl,PostSet,PostData,PostReferer)
    If InStr(LCase(PostUrl),"http://") = 0 Then
        PostHttpPage = "$Null$":Exit Function
    End If
    On Error Resume Next
    Dim PostHttp
    'Set PostHttp = Server.CreateObject("MSXML2.XMLHttp")
    'Set PostHttp = Server.CreateObject("Microsoft.XMLHTTP")
    Set PostHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    'Set PostHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    'Set PostHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.4.0")
    PostHttp.SetTimeOuts 10000, 10000, 15000, 15000    
    PostHttp.open "POST", PostUrl, False
    PostHttp.setRequestHeader "Content-Length",Len(PostData) 
    PostHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    PostHttp.setRequestHeader "Referer", PostReferer
    PostHttp.Send PostData
    If PostHttp.Readystate <> 4 And PostHttp.status <> 200 Then
        Set PostHttp = Nothing
        PostHttpPage = "$Null$":Exit function
    End If
    PostHttpPage = BytesToBstr(PostHttp.responseBody,PostSet)
    Set PostHttp = Nothing
    If Err.number<>0 Then Err.Clear
    If PostHttpPage = "" Or IsNull(PostHttpPage) Then PostHttpPage = "$Null$"
End Function

Function BytesToBstr(Body,Cset)
    Dim Objstream
    Set Objstream = Server.CreateObject("adodb.stream")
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

'==================================================
'函数名：GetBody
'作 用：截取字符串
'参 数：ConStr ------将要截取的字符串
'参 数：StartStr ------开始字符串
'参 数：OverStr ------结束字符串
'参 数：IncluL ------是否包含StartStr
'参 数：IncluR ------是否包含OverStr
'==================================================
Public Function GetBody(ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or IsNull(ConStr) = True Or StartStr = "" Or IsNull(StartStr) = True Or OverStr = "" Or IsNull(OverStr) = True Then
        GetBody = "$False$"
        Exit Function
    End If
    Dim ConStrTemp
    Dim start, Over
    ConStrTemp = LCase(ConStr)
    StartStr = LCase(StartStr)
    OverStr = LCase(OverStr)
    start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
    If start <= 0 Then
        GetBody = "$False$"
        Exit Function
    Else
        If IncluL = False Then
            start = start + LenB(StartStr)
        End If
    End If
    Over = InStrB(start, ConStrTemp, OverStr, vbBinaryCompare)
    If Over <= 0 Or Over <= start Then
        GetBody = "$False$"
        Exit Function
    Else
        If IncluR = True Then
            Over = Over + LenB(OverStr)
        End If
    End If
    GetBody = MidB(ConStr, start, Over - start)
End Function
%>