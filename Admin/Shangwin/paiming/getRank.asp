<!--#include file="../../easp.asp"-->
<%dim XmlHttpData,strReferer,u,k,r
response.Charset="utf-8"
session.CodePage=65001
u=asp.r("u",0)
'u="qebang.cn"
k=asp.r("k",0)
'strReferer="http://tool.chinaz.com/KeyWords/"
'XmlHttpData="se=0&kw="&k&"&host="&u&"&serverguid=&pn=100&kwsubmit=%E6%9F%A5%E8%AF%A2%E5%85%B3%E9%94%AE%E5%AD%97%E6%8E%92%E5%90%8D&page=0"
'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
'with objXML
'	.open "POST","http://tool.chinaz.com/KeyWords/",false
'	.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'	.setRequestHeader "Accept", "image/gif,image/jpeg,image/pjpeg,image/pjpeg,application/x-shockwave-flash,application/vnd.ms-excel,application/vnd.ms-powerpoint,application/msword,*/*"
'    .setRequestHeader "Referer", strReferer
'    .setRequestHeader "Accept-Language", "zh-cn"
'    .setRequestHeader "User-Agent", "Mozilla/4.0(compatible;MSIE8.0;WindowsNT5.1;Trident/4.0)"
'    .setRequestHeader "Host", "control.blog.sina.com.cn"
'    .setRequestHeader "Content-Length", Len(XmlHttpData)
'    .setRequestHeader "Connection", "Keep-Alive"
'    .setRequestHeader "Cache-Control", "no-cache"
'	.send(XmlHttpData)
'	'asp.we .responseText
'	r=Replace(asp.HtmlFilter(trim(strCut(.responseText,"第<span","个",true,false))),"第 ","")
'	if r="$False$" then
'		asp.w "Baidu:0 "
'	else
'		asp.w "Baidu:<b>"&r&"</b> "
'	end if
'end with
Dim url,bdrank,ggrank,rGg,rBd,rank

r=getHTTPPage("http://tool.aspxhome.com/keywordrank.asp?key="&k&"&domain="&u&"&page=10&baidu=checked&google=checked","gb2312")
r=strCut(r,"您查询的关键词","最近查询",false,false)
rGg=strCut(r,"Google","</span>",false,true)
rBd=strCut(r,"百度","</span>",false,true)
rGg=strCut(rGg,"<span","</span>",false,false)
rBd=strCut(rBd,"<span","</span>",false,false)
If InStr(rGg,"不在10页内")>0 And InStr(rBd,"不在10页内")>0 Then
	rank="Baidu:0 Google:0"
Else
	If InStr(rGg,"$False$")>0 Then 
		ggrank="Google:0"
	else
		ggrank="Google:<b>"&Replace(strCut(rGg,"第","名",false,false)," ","")&"</b>"
	End If
	If InStr(rBd,"$False$")>0 Then
		bdrank="Baidu:0 "
	else
		bdrank="Baidu:<b>"&Replace(strCut(rBd,"第","名",false,false)," ","")&"</b> "
	End if
	rank=bdrank&ggrank
End if
asp.w rank

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

'==================================================
'函数名：GetBody
'作 用：截取字符串
'参 数：ConStr ------将要截取的字符串
'参 数：StartStr ------开始字符串
'参 数：OverStr ------结束字符串
'参 数：IncluL ------是否包含StartStr
'参 数：IncluR ------是否包含OverStr
'==================================================
Public Function strCut(ConStr, StartStr, OverStr, IncluL, IncluR)
    If ConStr = "$False$" Or ConStr = "" Or IsNull(ConStr) = True Or StartStr = "" Or IsNull(StartStr) = True Or OverStr = "" Or IsNull(OverStr) = True Then
        strCut = "$False$"
        Exit Function
    End If
    Dim ConStrTemp
    Dim start, Over
    ConStrTemp = LCase(ConStr)
    StartStr = LCase(StartStr)
    OverStr = LCase(OverStr)
    start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
    If start <= 0 Then
        strCut = "$False$"
        Exit Function
    Else
        If IncluL = False Then
            start = start + LenB(StartStr)
        End If
    End If
    Over = InStrB(start, ConStrTemp, OverStr, vbBinaryCompare)
    If Over <= 0 Or Over <= start Then
        strCut = "$False$"
        Exit Function
    Else
        If IncluR = True Then
            Over = Over + LenB(OverStr)
        End If
    End If
    strCut = MidB(ConStr, start, Over - start)
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


Private Function XMLHttpGet(ByVal XmlHttpURL)
    Dim MyXmlhttp
    On Error Resume next
    Set MyXmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")                  '创建WinHttpRequest对象
    With MyXmlhttp
        .setTimeouts 50000, 50000, 50000, 50000                                 '设置超时时间
        .Open "GET", XmlHttpURL, True
        '无Http头信息
        .send (null)
        .waitForResponse                                                        '异步等待
        If MyXmlhttp.Status = 200 Then                                          '成功获取页面
            XMLHttpRequest = BytesToBstr(.ResponseBody, "gb2312")
        Else
			XMLHttpGet=""
        End If
    End With
    Set MyXmlhttp = Nothing
    Exit Function
	If Err Then
		Err.clear
		XMLHttpGet=""
		Set MyXmlhttp = Nothing
	End if
End Function
%>