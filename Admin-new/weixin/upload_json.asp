<!--#Include File="../AdminCheck.asp"-->
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="UpLoad_Class.asp"-->
<!--#include file="JSON_2.0.4.asp"-->
<%

' KindEditor ASP
'
' 本ASP程序是演示程序，建议不要直接在实际项目中使用。
' 如果您确定直接使用本程序，使用之前请仔细确认相关安全设置。
'

Function ByteToStr(vIn)
	Dim strReturn,i,ThisCharCode,innerCode,Hight8,Low8,NextCharCode
	strReturn = "" 
	For i = 1 To LenB(vIn)
	ThisCharCode = AscB(MidB(vIn,i,1))
	If ThisCharCode < &H80 Then
	strReturn = strReturn & Chr(ThisCharCode)
	Else
	NextCharCode = AscB(MidB(vIn,i+1,1))
	strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
	i = i + 1
	End If
	Next
	ByteToStr = strReturn 
End Function

Function DoGet(url)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	'.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send()
	DoGet = .ResponseText
	End With
	Set Http = Nothing
	'DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoGet=""
	end if
End Function

Function DoPost(url,PostStr)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send(PostStr)
	DoPost = .ResponseBody
	End With
	Set Http = Nothing
	DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoPost=""
	end if
End Function

Function strCut(strContent,StartStr,EndStr,CutType)
	Dim strHtml,S1,S2
	strHtml = strContent
	On Error Resume Next
	Select Case CutType
	Case 1
		S1 = InStr(strHtml,StartStr)
		S2 = InStr(S1,strHtml,EndStr)+Len(EndStr)
	Case 2
		S1 = InStr(strHtml,StartStr)+Len(StartStr)
		S2 = InStr(S1,strHtml,EndStr)
	End Select
	If Err Then
		strCute = ""
		Err.Clear
		Exit Function
	Else
		strCut = Mid(strHtml,S1,S2-S1)
	End If
End Function

Dim aspUrl, savePath, saveUrl, maxSize, fileName, fileExt, newFileName, filePath, fileUrl, dirName
Dim extStr, imageExtStr, flashExtStr, mediaExtStr, fileExtStr
Dim upload, file,  ranNum, hash, ymd, mm, dd, result

aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))

'文件保存目录路径
savePath = "../../up/"
'文件保存目录URL
saveUrl = "/up/"
'定义允许上传的文件扩展名
imageExtStr = "gif|jpg|png|bmp"
mediaExtStr = "amr|mp3|wav|wma"


Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.mappath(savePath)) Then
	showError("上传目录不存在。")
End If

dirName = Request.QueryString("dir")
If isEmpty(dirName) Then
	dirName = "image"
End If
If instr(lcase("image,flash,media,file"), dirName) < 1 Then
	dirName = "image"
End If


Select Case dirName
	Case "file" extStr = mediaExtStr:maxSize = 5 * 1024 * 1024 '5M
	Case Else  extStr = imageExtStr:maxSize = 2 * 1024 * 1024 '2M
End Select

set upload = new AnUpLoad
upload.Exe = extStr
upload.MaxSize = maxSize
upload.GetData()
if upload.ErrorID>0 then 
	showError(upload.Description)
end if

dim wx_AppId,wx_AppSecret
set rs=conn.execute("select wx_AppId,wx_AppSecret from weixin_config")
if not rs.eof then
	wx_AppId	 = rs("wx_AppId")
	wx_AppSecret = rs("wx_AppSecret")
end if
rs.close
dim access_token,returnMsg
access_token=DoGet("https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_AppId&"&secret="&wx_AppSecret)
access_token=strCut(access_token,"access_token"":""","""",2)
returnMsg=DoPost("http://file.api.weixin.qq.com/cgi-bin/media/upload?access_token="&access_token&"&type=image",Request.TotalBytes)
Response.Write returnMsg
if returnMsg="{""errcode"":0,""errmsg"":""ok""}" then
	Response.Write("菜单生成成功！请重启微信查看菜单效果")
else
	Response.Write("菜单生成失败！请重试")
end if

'创建文件夹
savePath = savePath & dirName & "/"
saveUrl = saveUrl & dirName & "/"
If Not fso.FolderExists(Server.mappath(savePath)) Then
	fso.CreateFolder(Server.mappath(savePath))
End If
mm = month(now)
If mm < 10 Then
	mm = "0" & mm
End If
dd = day(now)
If dd < 10 Then
	dd = "0" & dd
End If
ymd = year(now) & mm 
savePath = savePath & ymd & "/"
saveUrl = saveUrl & ymd & "/"
If Not fso.FolderExists(Server.mappath(savePath)) Then
	fso.CreateFolder(Server.mappath(savePath))
End If

set file = upload.files("imgFile")
if file is nothing then
	showError("请选择文件。")
end if

set result = file.saveToFile(savePath, 0, true)
if result.error then
	showError(file.Exception)
end if

filePath = Server.mappath(savePath & file.filename)
fileUrl = saveUrl & file.filename

Set upload = nothing
Set file = nothing

If Not fso.FileExists(filePath) Then
	showError("上传文件失败。")
End If

Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
Set hash = jsObject()
hash("error") = 0
hash("url") = fileUrl
hash.Flush
Response.End

Function showError(message)
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	Dim hash
	Set hash = jsObject()
	hash("error") = 1
	hash("message") = message
	hash.Flush
	Response.End
End Function
%>
