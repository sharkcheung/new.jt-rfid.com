<%
'****************************************************
' Software name:KQIQIECMS
' Email: 3955233@qq.com . QQ:3955233
' Web: http://www.kqiqi.com http://www.kqiqi.cn
' Copyright (C) Kqiqi Network All Rights Reserved.
'****************************************************
'SHA256¼ÓÃÜº¯Êý
Dim kqiqi_m_l2Power(30)
Dim kqiqi_m_lOnBits(30)
kqiqi_m_l2Power(0) = CLng(1)
kqiqi_m_l2Power(1) = CLng(2)
kqiqi_m_l2Power(2) = CLng(4)
kqiqi_m_l2Power(3) = CLng(8)
kqiqi_m_l2Power(4) = CLng(16)
kqiqi_m_l2Power(5) = CLng(32)
kqiqi_m_l2Power(6) = CLng(64)
kqiqi_m_l2Power(7) = CLng(128)
kqiqi_m_l2Power(8) = CLng(256)
kqiqi_m_l2Power(9) = CLng(512)
kqiqi_m_l2Power(10) = CLng(1024)
kqiqi_m_l2Power(11) = CLng(2048)
kqiqi_m_l2Power(12) = CLng(4096)
kqiqi_m_l2Power(13) = CLng(8192)
kqiqi_m_l2Power(14) = CLng(16384)
kqiqi_m_l2Power(15) = CLng(32768)
kqiqi_m_l2Power(16) = CLng(65536)
kqiqi_m_l2Power(17) = CLng(131072)
kqiqi_m_l2Power(18) = CLng(262144)
kqiqi_m_l2Power(19) = CLng(524288)
kqiqi_m_l2Power(20) = CLng(1048576)
kqiqi_m_l2Power(21) = CLng(2097152)
kqiqi_m_l2Power(22) = CLng(4194304)
kqiqi_m_l2Power(23) = CLng(8388608)
kqiqi_m_l2Power(24) = CLng(16777216)
kqiqi_m_l2Power(25) = CLng(33554432)
kqiqi_m_l2Power(26) = CLng(67108864)
kqiqi_m_l2Power(27) = CLng(134217728)
kqiqi_m_l2Power(28) = CLng(268435456)
kqiqi_m_l2Power(29) = CLng(536870912)
kqiqi_m_l2Power(30) = CLng(1073741824)
kqiqi_m_lOnBits(0) = CLng(1)
kqiqi_m_lOnBits(1) = CLng(3)
kqiqi_m_lOnBits(2) = CLng(7)
kqiqi_m_lOnBits(3) = CLng(15)
kqiqi_m_lOnBits(4) = CLng(31)
kqiqi_m_lOnBits(5) = CLng(63)
kqiqi_m_lOnBits(6) = CLng(127)
kqiqi_m_lOnBits(7) = CLng(255)
kqiqi_m_lOnBits(8) = CLng(511)
kqiqi_m_lOnBits(9) = CLng(1023)
kqiqi_m_lOnBits(10) = CLng(2047)
kqiqi_m_lOnBits(11) = CLng(4095)
kqiqi_m_lOnBits(12) = CLng(8191)
kqiqi_m_lOnBits(13) = CLng(16383)
kqiqi_m_lOnBits(14) = CLng(32767)
kqiqi_m_lOnBits(15) = CLng(65535)
kqiqi_m_lOnBits(16) = CLng(131071)
kqiqi_m_lOnBits(17) = CLng(262143)
kqiqi_m_lOnBits(18) = CLng(524287)
kqiqi_m_lOnBits(19) = CLng(1048575)
kqiqi_m_lOnBits(20) = CLng(2097151)
kqiqi_m_lOnBits(21) = CLng(4194303)
kqiqi_m_lOnBits(22) = CLng(8388607)
kqiqi_m_lOnBits(23) = CLng(16777215)
kqiqi_m_lOnBits(24) = CLng(33554431)
kqiqi_m_lOnBits(25) = CLng(67108863)
kqiqi_m_lOnBits(26) = CLng(134217727)
kqiqi_m_lOnBits(27) = CLng(268435455)
kqiqi_m_lOnBits(28) = CLng(536870911)
kqiqi_m_lOnBits(29) = CLng(1073741823)
kqiqi_m_lOnBits(30) = CLng(2147483647)

Function HashPR(url)
	Dim lByteCount:lByteCount=0
	Dim seed:seed="Mining PageRank is AGAINST GOOGLE'S TERMS OF SERVICE. Yes, I'm talking to you, scammer."
	Dim result:result=&H01020345
	Dim urlLength:urlLength=Len(url)
	Do Until lByteCount >= urlLength
		result=result  Xor Asc(Mid(seed,(lByteCount Mod Len(seed))+1,1)) Xor Asc(Mid(url,lByteCount+1,1))
		result=(RShift(result,23) And &H1FF)  Or LShift(result,9)
		lByteCount = lByteCount + 1
	Loop
	HashPR="8"&LCase(hex(result))
End function

Private Function WordToHex(lValue)
    Const BITS_TO_A_BYTE = 8
	Dim lByte,lCount
	For lCount = 0 To 3
		lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And kqiqi_m_lOnBits(BITS_TO_A_BYTE - 1)
		WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
	Next
End Function

Private Function LShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		LShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And 1 Then LShift = &H80000000 Else LShift = 0:Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	If (lValue And kqiqi_m_l2Power(31 - iShiftBits)) Then
		LShift = ((lValue And kqiqi_m_lOnBits(31 - (iShiftBits + 1))) * kqiqi_m_l2Power(iShiftBits)) Or &H80000000
	Else
		LShift = ((lValue And kqiqi_m_lOnBits(31 - iShiftBits)) * kqiqi_m_l2Power(iShiftBits))
	End If
End Function

Function RShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		RShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And &H80000000 Then RShift = 1 Else RShift = 0:Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	RShift = (lValue And &H7FFFFFFE) \ kqiqi_m_l2Power(iShiftBits)
	If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ kqiqi_m_l2Power(iShiftBits - 1)))
End Function

Function GetHttpPage(HttpUrl, Coding)
    On Error Resume Next
    If IsNull(HttpUrl) = True Or Len(HttpUrl) < 10 Or HttpUrl = "" Then GetHttpPage = "":Exit Function
    Dim Http:Set Http = Server.CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", HttpUrl, False
    Http.Send
    If Http.Readystate <> 4 Then GetHttpPage = "":Exit Function
    Select Case Coding
     Case 1
        GetHttpPage = BytesToBstr(Http.ResponseBody, "UTF-8")
     Case 2
        GetHttpPage = BytesToBstr(Http.ResponseBody, "Big5")
     Case Else
        GetHttpPage = BytesToBstr(Http.ResponseBody, "GB2312")
    End Select
    Set Http = Nothing
    If Err.Number <> 0 Then Err.Clear
End Function

Function BytesToBstr(Body, Cset)
    Dim Objstream:Set Objstream = Server.CreateObject("adodb.stream")
    With Objstream
     .Type = 1
     .Mode = 3
     .Open
     .Write Body
     .Position = 0
     .Type = 2
     .Charset = Cset
    End With
    BytesToBstr = Objstream.ReadText
    Objstream.Close
    Set Objstream = Nothing
End Function

Function GetGooglePR(url)
	Dim gurl,gcontent
	gurl = "http://www.google.com/search?client=navclient-auto&features=Rank:&q=info:"&url&"&ch="& HashPR(url)
	gcontent=GetHttpPage(gurl,1)
	Select Case Len(gcontent)
	 Case 11
		GetGooglePR=Trim(Mid(gcontent,10,1))
	 Case 12
		GetGooglePR=Trim(Mid(gcontent,10,2))
	 Case Else
		GetGooglePR=0
	End Select
End function

Function CheckUrl(url)
    CheckUrl = false
    Dim re, Match
    Set re = New RegExp
    re.Pattern = "^([A-Z0-9][A-Z0-9_-]*(?:\.[A-Z0-9][A-Z0-9_-]*)+):?(\d+)?\/?$" 
    re.IgnoreCase = True
    Set Match = re.Execute(url)
    if match.count then CheckUrl= True
    Set re=nothing
End Function

wd=Request("d")


If CheckUrl(wd) Then
	cook=wd&"ggpr"
	cook=replace(cook,".","")
	if Request.Cookies(cook)<>"" then
		PR=Request.Cookies(cook)
	else
 		PR=GetGooglePR(wd) 
 		Response.Cookies(cook)=PR
		Response.Cookies(cook).Expires=dateadd("h",72,now())
 	end if
Else
 PR=0
end if
'PR_Images=server.mappath("/icon/pagerank"&PR&".gif")
'PrintPR(PR_Images)
'CloseDb
'ResetClass
%>document.getElementById('mypr').innerHTML = '<%=pr%>/10';