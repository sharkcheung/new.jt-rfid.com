<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="UpLoad_Class.asp"-->
<!--#include file="JSON_2.0.4.asp"-->
<!--#Include File="../../inc/md5.asp"-->
<!--#Include File="../config.asp"-->
<%
Function getIP()
	Dim sIPAddress, sHTTP_X_FORWARDED_FOR
	 sHTTP_X_FORWARDED_FOR = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If sHTTP_X_FORWARDED_FOR = "" Or InStr(sHTTP_X_FORWARDED_FOR, "unknown") > 0 Then
		sIPAddress = Request.ServerVariables("REMOTE_ADDR")
	ElseIf InStr(sHTTP_X_FORWARDED_FOR, ",") > 0 Then
		sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ",") -1)
	ElseIf InStr(sHTTP_X_FORWARDED_FOR, ";") > 0 Then
		sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ";") -1)
	Else
		sIPAddress = sHTTP_X_FORWARDED_FOR
	End If
	getIP = Trim(Mid(sIPAddress, 1, 15))
End Function

dim readjson,json,code,n,domainname,curtime,strcode,requestip
requestip = getIP()
if instr(whiteIPList,requestip)=0 then
	showError("no permission")
end if

code = Request.ServerVariables("HTTP_CODE")

if IsEmpty(code) then
	showError("code is null")
end if

domainname = Request.ServerVariables("SERVER_NAME")
curtime = year(date)&right("0"&month(date),2)&right("0"&day(date),2)


strcode = md5(domainname&curtime,32)

if code<>strcode then
	showError("校验失败")
end if

Dim aspUrl, savePath, saveUrl, maxSize, fileName, fileExt, newFileName, filePath, fileUrl, dirName
Dim extStr, imageExtStr, flashExtStr, mediaExtStr, fileExtStr
Dim upload, file, fso, ranNum, hash, ymd, mm, dd, result

aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))

'文件保存目录路径
savePath = "../../up/"
'文件保存目录URL
saveUrl = "/up/"
'定义允许上传的文件扩展名
imageExtStr = "gif|jpg|jpeg|png|bmp"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.mappath(savePath)) Then
	showError("上传目录不存在。")
End If

dirName = "image"
extStr = imageExtStr
maxSize = 2 * 1024 * 1024 '2M

set upload = new AnUpLoad
upload.Exe = extStr
upload.MaxSize = maxSize
upload.GetData()
if upload.ErrorID>0 then 
	showError(upload.Description)
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

set file = upload.files("file")
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

Response.AddHeader "Content-Type", "text/json; charset=UTF-8"
Set hash = jsObject()
hash("success") = true
hash("message") = "图片上传成功"
hash("url") = fileUrl
hash.Flush
Response.End

Function showError(message)
	Response.AddHeader "Content-Type", "text/json; charset=UTF-8"
	Dim hash
	Set hash = jsObject()
	hash("success") = false
	hash("message") = message
	hash.Flush
	Response.End
End Function
%>
