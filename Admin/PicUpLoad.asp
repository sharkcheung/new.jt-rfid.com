<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_Upload.asp"--><%
'==========================================
'文 件 名：PicUpLoad.asp
'文件用途：文件上传
'系统开发：深圳企帮
'==========================================
Dim JpegNow,Fk_Jpeg_Water,Fk_Jpeg_Small,JpegObjs,SmallPic,Pic,TWidth,THeight
JpegNow=FKFun.IsObjInstalled("Persits.Jpeg")
If JpegNow=False Then
	Fk_Jpeg_Water=0
	Fk_Jpeg_Small=0
Else
	Sqlstr="Select * From [Fk_Jpeg]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Jpeg_Water=Rs("Fk_Jpeg_Water")
		Fk_Jpeg_Small=Rs("Fk_Jpeg_Small")
		TempArr=Split(Rs("Fk_Jpeg_Content"),"|-_-|")
	Else
		Fk_Jpeg_Water=0
		Fk_Jpeg_Small=0
	End If
	Rs.Close
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>上传图片</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {
	font-size: 12px;
	color: #000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000;
}
a:hover {
	text-decoration: none;
	color: #333;
}
a:active {
	text-decoration: none;
	color: #000;
}
.Input,.Button {
	height:22px;
	line-height:22px;
	border:1px solid #E0E0E0;
	border-bottom:1px solid #CCC;
	border-right:1px solid #CCC;
	font-size:12px;
}
.Button {
	margin-left:0;
}
-->
</style></head>
<body>
<%
Call UpPic()

'==============================
'函 数 名：UpPic
'作    用：上传图片
'参    数：
'==============================
Sub UpPic()
	If Request.QueryString("Submit")="Pic" Then
		Dim UploadPath,UploadPath2,UploadSize,UploadType,UpRequest,AutoSave,Size,ShowSize,Str
		UploadPath="../Up/"
		UploadPath2="Up/"
		UploadSize="200"
		UploadType="jpg/gif/png/bmp/rar/zip/doc/xls/ppt"
		Set UpRequest=New UpLoadClass
		UpRequest.SavePath=UploadPath
		UpRequest.MaxSize=UploadSize*1024 
		UpRequest.FileType=UploadType
		AutoSave=true
		UpRequest.Open
		on error resume next
		If UpRequest.Form("file_Err")<>0  Then
			Select Case UpRequest.Form("file_Err")
				Case 1:Str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>不成功!超过"&UploadSize&"k [<a href='javascript:history.go(-1)'>重传</a>]</font></div>"
				Case 2:Str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>不成功!格式不对 [<a href='javascript:history.go(-1)']>重传</a>]</font></div>"
				Case 3:Str="<div style=""padding-top:5px;padding-bottom:5px;""> <font color=blue>不成功!太大且格式不对 [<a href='javascript:history.go(-1)'>重传</a>]</font></div>"
			End Select
			Response.Write Str
		Else
			Pic=SiteDir&UploadPath2&UpRequest.Form("File")
			'==========================================
			'Edit   :  Shark
			'Time   :  2013-11-16 1:09
			'Descrip:  判断上传文件是否非法，非法则删除
			'==========================================
			if UpRequest.CheckFileContent(Pic) then
				Set fso = CreateObject("Scripting.FileSystemObject")
      			fso.DeleteFile(Server.mappath(Pic))
      			Set fso = nothing
				Set UpRequest=nothing
				Response.Write "<div style=""padding-top:5px;padding-bottom:5px;""> <font color=""#CC6600"">不成功!您所上传的图片中含有恶意代码 [<a href='javascript:history.go(-1)'>重传</a>]</font></div>"
				response.end
			end if
			
			If Right(Pic,3)="jpg" Then
				If Fk_Jpeg_Small=1 Then
					SmallPic=Replace(Pic,".","_small.")
					Set JpegObjs=Server.CreateObject("Persits.Jpeg")
					JpegObjs.Open Server.MapPath(Pic)
					If Clng(JpegObjs.OriginalWidth)<Clng(TempArr(15)) And Clng(JpegObjs.OriginalHeight)<Clng(TempArr(16)) Then
						SmallPic=Pic
					Else
						TWidth=JpegObjs.OriginalWidth
						THeight=JpegObjs.OriginalHeight
						If Clng(TWidth)>Clng(TempArr(15)) Then
							THeight=TempArr(15)/TWidth*THeight
							TWidth=TempArr(15)
						End If
						If Clng(THeight)>Clng(TempArr(16)) Then
							TWidth=TempArr(16)/THeight*TWidth
							THeight=TempArr(16)
						End If
						Call FKFun.DoSmall(Pic,SmallPic,TWidth,THeight)
					End If
				Else
					SmallPic=Pic
				End If
				If Fk_Jpeg_Water=1 Then
					Call FKFun.DoWater(Pic,Pic,TempArr(0),TempArr(1),TempArr(2),TempArr(3),TempArr(4),TempArr(5))
					If SmallPic<>Pic Then
						'Call FKFun.DoWater(SmallPic,SmallPic,TempArr(0),TempArr(1),TempArr(2),TempArr(3),TempArr(4),TempArr(5))
					End If
				End If
			Else
				SmallPic=Pic
			End If
			Response.Write "<script language=""javascript"">parent."&Request.QueryString("Form")&"."&Request.QueryString("Input")&".value='"&SmallPic&"';" 
			Response.Write "parent."&Request.QueryString("Form")&"."&Request.QueryString("Input")&"Big.value='"&Pic&"';" 
			Response.Write "</script>"
			Size=UpRequest.Form("file_size")
			ShowSize=Size & " Byte"   
			If Size>1024 Then  
				Size=(Size\1024)  
				ShowSize=Size & " KB"  
			End If  
			If Size>1024 Then  
				Size=(Size/1024)  
				ShowSize=FormatNumber(Size,2) & " MB"		  
			End If 
			Response.Write "<div style=""padding-top:5px;padding-bottom:5px;""> <font color=red>上传成功</font> [<a href='javascript:history.go(-1)'>重新上传</a>]</div>"
		End If
		Set UpRequest=nothing
	End If
	Response.Write "<Form name=Form action='?Submit=Pic&Form="&Request.QueryString("Form")&"&Input="&Request.QueryString("Input")&"' method=post enctype=multipart/Form-data>"
	Response.Write "<input type='file' name='file' class='Input' Size='0' style='width:0px;'>&nbsp;"
	Response.Write "<input type='submit' name='submit' value='上传' class=""Button"">"
	Response.Write "</Form>"
End sub
%>
</body>
</html>