<!--#Include File="include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<%
dim op
op=FKFun.HTMLEncode(Trim(Request("op")))
if op = "Sync_KeyWord" then
	Call sync_keyword() '同步关键词
end if

'转换时间 时间格式化 
Function formatDate(Byval t,Byval ftype) 
	dim y, m, d, h, mi, s 
	formatDate=""
	If IsDate(t)=False Then Exit Function
	y=cstr(year(t)) 
	m=cstr(month(t)) 
	If len(m)=1 Then m="0" & m 
	d=cstr(day(t)) 
	If len(d)=1 Then d="0" & d 
	h = cstr(hour(t)) 
	If len(h)=1 Then h="0" & h 
	mi = cstr(minute(t)) 
	If len(mi)=1 Then mi="0" & mi 
	s = cstr(second(t)) 
	If len(s)=1 Then s="0" & s 
	select case cint(ftype) 
	case 1 
	' yyyy-mm-dd 
	formatDate=y & "-" & m & "-" & d 
	case 2 
	' yy-mm-dd 
	formatDate=right(y,2) & "-" & m & "-" & d 
	case 3 
	' mm-dd 
	formatDate=m & "-" & d 
	case 4 
	' yyyy-mm-dd hh:mm:ss 
	formatDate=y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
	case 5 
	' hh:mm:ss 
	formatDate=h & ":" & mi & ":" & s 
	case 6 
	' yyyy年mm月dd日 
	formatDate=y & "年" & m & "月" & d & "日"
	case 7 
	' yyyymmdd 
	formatDate=y & m & d 
	case 8 
	'yyyymmddhhmmss 
	formatDate=y & m & d & h & mi & s 
	case 9
	' yyyy-mm-dd hh:mm:ss 
	formatDate=y & "-" & m & "-" & d 
	end select 
End Function

sub chkToken(strMobile,strUsertype,strToken)
	dim token,strTime,strWebToken
	Call FKFun.ShowString(strMobile,1,50,0,"非法操作，001","非法操作，001")
	Call FKFun.ShowString(strUsertype,1,50,0,"非法操作，002","非法操作，002")
	Call FKFun.ShowString(strToken,1,50,0,"非法操作，003","非法操作，003")
	token="3PVcDkYEbL8dXuaTM5JUzNjbPCWRuQq5"
    strTime = formatDate(Now, 9)
    strWebToken = MD5(strMobile & token & strUsertype & strTime, 32)
	''response.write strToken&"_"&strWebToken
	if strToken<>strWebToken then
		response.write "非法操作，004"
		response.end
	end if
end sub


Function DecodeURI(ByVal s)
    s = UnEscape(s)
    Dim cs : cs = "GBK"
    With New RegExp
        .Pattern = "^(?:[\x00-\x7f]|[\xfc-\xff][\x80-\xbf]{5}|[\xf8-\xfb][\x80-\xbf]{4}|[\xf0-\xf7][\x80-\xbf]{3}|[\xe0-\xef][\x80-\xbf]{2}|[\xc0-\xdf][\x80-\xbf])+$"
        If .Test(s) Then cs = "UTF-8"
    End With
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Mode = 3
        .Open
        .CharSet = "iso-8859-1"
        .WriteText s
        .Position = 0
        .CharSet = cs
        DecodeURI = .ReadText(-1)
        .Close
    End With
End Function

'========================
'函数名：cxarraynull
'作  用：关键词去重
'参  数：cxstr1:要去重的关键词串;cxstr2:分割符
'========================
function cxarraynull(cxstr1,cxstr2)
dim ss,sss,cxs,cc,m
if isarray(cxstr1) then
cxarraynull = ""
Exit Function
end if
if cxstr1 = "" or isempty(cxstr1) then
cxarraynull = ""
Exit Function
end if
do while instr(cxstr1,",,")>0
cxstr1=replace(cxstr1,",,",",")
loop
if right(cxstr1,1)="," then
cxstr1=left(cxstr1,len(cxstr1)-1)
end if
ss = split(cxstr1,cxstr2)
cxs = cxstr2&ss(0)&cxstr2
sss = cxs
for m = 0 to ubound(ss)
cc = cxstr2&ss(m)&cxstr2
if instr(sss,cc)=0 then
sss = sss&ss(m)&cxstr2
end if
next
cxarraynull = right(sss,len(sss) - len(cxstr2))
cxarraynull = left(cxarraynull,len(cxarraynull) - len(cxstr2))
end function
'==========================================
'函 数 名：sync_keyword()
'作    用：同步关键词库
'参    数：
'==========================================
Sub sync_keyword()
	dim k,KeyWord
	dim strMobile,strUsertype,strToken
	strMobile=Request("mobile")
	strUsertype=Request("usertype")
	strToken=Request("token")
	call chkToken(strMobile,strUsertype,strToken)
	k=Request.form("k")
	if k<>"" then
	KeyWord=ClearSG(cxarraynull(FilterText(k),"|"))
	Call FKFso.CreateFile("KeyWordC.dat",KeyWord)
	Response.Write("关键词库修改成功！")
	else
	Response.Write("关键词为空")
	end if
End Sub

function ClearRightDh(str)
	while(right(str,1)=",")
		str=left(str,len(str)-1)
	wend
	ClearRightDh=str
end function 

function ClearSG(str)
	while(right(str,1)="|")
		str=left(str,len(str)-1)
	wend
	while(left(str,1)="|")
		str=mid(str,2)
	wend
	ClearSG=str
end function 

'===================================== 
'过滤字符 
'===================================== 
Function FilterText(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
t0=Replace(t0,"||","|")'  
FilterText=t0 
End Function 

'===================================== 
'过滤字符 
'===================================== 
Function Filterkwd(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
Filterkwd=t0 
End Function 

Function Easp_Escape(ByVal str)
	Dim i,c,a,s : s = ""
	If isnull(str) Then Easp_Escape = "" : Exit Function
	For i = 1 To Len(str)
		c = Mid(str,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Easp_Escape = s
End Function

Sub Shuffle (ByRef arrInput)
    'declare local variables:
    Dim arrIndices, iSize, x
    Dim arrOriginal

    'calculate size of given array:
    iSize = UBound(arrInput)+1

    'build array of random indices:
    arrIndices = RandomNoDuplicates(0, iSize-1, iSize)

    'copy:
    arrOriginal = CopyArray(arrInput)

    'shuffle:
    For x=0 To UBound(arrIndices)
        arrInput(x) = arrOriginal(arrIndices(x))
    Next
End Sub

Function CopyArray (arr)
    Dim result(), x
    ReDim result(UBound(arr))
    For x=0 To UBound(arr)
        If IsObject(arr(x)) Then
            Set result(x) = arr(x)
        Else
            result(x) = arr(x)
        End If
    Next
    CopyArray = result
End Function

Function RandomNoDuplicates (iMin, iMax, iElements)
    'this function will return array with "iElements" elements, each of them is random
    'integer in the range "iMin"-"iMax", no duplicates.

    'make sure we won't have infinite loop:
    If (iMax-iMin+1)>iElements Then
        Exit Function
    End If

    'declare local variables:
    Dim RndArr(), x, curRand
    Dim iCount, arrValues()

    'build array of values:
    Redim arrValues(iMax-iMin)
    For x=iMin To iMax
        arrValues(x-iMin) = x
    Next

    'initialize array to return:
    Redim RndArr(iElements-1)

    'reset:
    For x=0 To UBound(RndArr)
        RndArr(x) = iMin-1
    Next

    'initialize random numbers generator engine:
    Randomize
    iCount=0

    'loop until the array is full:
    Do Until iCount>=iElements
        'create new random number:
        curRand = arrValues(CLng((Rnd*(iElements-1))+1)-1)

        'check if already has duplicate, put it in array if not
        If Not(InArray(RndArr, curRand)) Then
            RndArr(iCount)=curRand
            iCount=iCount+1
        End If

        'maybe user gave up by now...
        If Not(Response.IsClientConnected) Then
            Exit Function
        End If
    Loop

    'assign the array as return value of the function:
    RandomNoDuplicates = RndArr
End Function

Function InArray(arr, val)
    Dim x
    InArray=True
    For x=0 To UBound(arr)
        If arr(x)=val Then
            Exit Function
        End If
    Next
    InArray=False
End Function


'************************* 
'函数:UBoundStrToArr 
'作用:检测原字符串转换为数组的最大下标值 
'参数:cCheckStr(需要检测的字符串) 
' cUBoundArr(生成数组的最大下标值) 
' cSpaceStr(间隔字符串) 
'返回:数组的最大下标值 
'************************ 
Public Function UBoundStrToArr(ByVal cCheckStr,ByVal cUBoundArr,ByVal cSpaceStr) 
On Error Resume Next

If Instr(cCheckStr,cSpaceStr)=0 Then 
UBoundStrToArr=cUBoundArr 
Exit Function 
End If 
Dim TempSpaceStr,UBoundValue 
TempSpaceStr=Mid(cCheckStr,Len(cCheckStr)-Len(cSpaceStr)+1) '获取字符串右侧间隔字符 
If TempSpaceStr=cSpaceStr Then '如果字符串最右侧存在间隔字符,则下标值需要-1 
UBoundValue=cUBoundArr-1 
Else 
UBoundValue=cUBoundArr 
End If 
UBoundStrToArr=UBoundValue 
End Function 


'********查询关键词在数据库某个表某个字段中出现的次数**********
Function Chakeywordci(Keywordsrt,Keywordslei)
 dim RSC1,RSC2,RSC3,SqlChastr
 RSC1=0:RSC2=0:RSC3=0
 select case Keywordslei
 case 2
	SqlChastr="Select Fk_Article_Title,Fk_Article_Keyword From Fk_Article Where Fk_Article_Title Like '%%"&Keywordsrt&"%%' or Fk_Article_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC1=Rs.RecordCount
	Rs.Close
	
	SqlChastr="Select Fk_Product_Title,Fk_Product_Keyword From Fk_Product Where Fk_Product_Title Like '%%"&Keywordsrt&"%%' or Fk_Product_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC2=Rs.RecordCount
	Rs.Close
case 1 	
	SqlChastr="Select Fk_Module_Keyword From Fk_Module Where Fk_Module_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC3=Rs.RecordCount
	Rs.Close
end select
	Chakeywordci=RSC1+RSC2+RSC3
End function

'****查询关键词是否有做内链**********
Function ChakeywordNLink(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select Fk_Word_Name From Fk_Word Where Fk_Word_Name Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			ChakeywordNLink=1
		else
			ChakeywordNLink=0
		end if
	Rs.Close
End function

'****查询关键词在数据库中的排名记录**********
Function Chanowpaiming(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select SVkeywords,SVpaiming From [keywordSV] Where SVkeywords='"&Keywordsrt&"' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			Chanowpaiming=Rs("SVpaiming")
		else
			Chanowpaiming="查询排名"
		end if
	Rs.Close
End function

'****查询strA中strB出现的次数**********
Function strCount(strA,strB)
lngA = Len(strA)
lngB = Len(strB)
lngC = Len(Replace(strA, strB, ""))
strCount = (lngA - lngC) / lngB
End Function

Function vbsEscape(str)
    dim i,s,c,a
    s=""
    For i=1 to Len(str)
        c=Mid(str,i,1)
        a=ASCW(c)
        If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
            s = s & c
        ElseIf InStr("@*_+-./",c)>0 Then
            s = s & c
        ElseIf a>0 and a<16 Then
            s = s & "%0" & Hex(a)
        ElseIf a>=16 and a<256 Then
            s = s & "%" & Hex(a)
        Else
            s = s & "%u" & Hex(a)
        End If
    Next
    vbsEscape = s
End Function

'页面结束
set rs=nothing
%>
<!--#Include File="../Code.asp"-->
