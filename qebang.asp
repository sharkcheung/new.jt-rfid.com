
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<head>
<style type="text/css">
* {
	font-size:14px;
}
</style>
</head>

<%
'Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
Response.Expires=-999
Session.Timeout=999
Dim StartTime,EndTime
StartTime=Timer()
%>
<!--#Include File="inc/Site.asp"-->

<%

'获取参数
Types=Clng(Request.QueryString("Type"))
if Request.QueryString("Type")="" then types=1

Select Case Types
	Case 1
		Call SiteSetBox() '读取系统信息
	Case 2
		Call SiteSetDo() '系统设置操作
	Case 3
		Call TestFetion() '测试飞信
End Select

'==========================================
'函 数 名SiteSetBox()
'作    用读取系统信息
'参    数
'==========================================
Sub SiteSetBox()
%>
<form id="SystemSet" name="SystemSet" method="post" action="qebang.asp?Type=2">

<div id="BoxContents" style="width:900px;">
	<table width="100%" border="0" align="center" id="table001"  cellpadding="0" cellspacing="0">
	    <tr>
            <td height="25" align="right" class="MainTableTop">公司名称</td>
            <td>&nbsp;<input name="SiteName" type="text" class="Input" id="SiteName" value="<%=SiteName%>" size="50" /></td>
            <td>域名</td>
            <td><input name="SiteUrl" type="text" class="Input" id="SiteUrl" value="<%=SiteUrl%>" size="32" /></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">电话</td>
            <td>&nbsp;<input name="Tel" type="text" class="Input" id="Tel" value="<%=Tel%>" size="32" /></td>
            <td>400热线</td>
            <td>
			<input name="Tel400" type="text" class="Input" id="Tel400" value="<%=Tel400%>" size="32" /></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">Email</td>
            <td>&nbsp;<input name="Email" type="text" class="Input" id="Email" value="<%=Email%>" size="32" /></td>
            <td>传真</td>
            <td>
			<input name="Fax" type="text" class="Input" id="Fax" value="<%=Fax%>" size="32" /></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">联系人</td>
            <td>&nbsp;<input name="Lianxiren" type="text" class="Input" id="Lianxiren" value="<%=Lianxiren%>" size="32" /></td>
            <td>备案号</td>
            <td>
			<input name="beian" type="text" class="Input" id="beian" value="<%=beian%>" size="32" /></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">地址</td>
            <td colspan="3">&nbsp;<input name="Add" type="text" class="Input" id="Add" value="<%=Add%>" size="110" />&nbsp;</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">SEO关键词</td>
            <td colspan="3">&nbsp;<input name="SiteKeyword" type="text" class="Input" id="SiteKeyword" value="<%=SiteKeyword%>" size="110" /></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">SEO描述</td>
            <td colspan="3">&nbsp;<input name="SiteDescription" type="text" class="Input" id="SiteDescription" value="<%=SiteDescription%>" size="110" /></td>
        </tr>
        </table>
</div>
<div id="BoxBottom" style="width:730px;">
        <input type="submit" class="Button" name="button" id="button" value="保存设置" />&nbsp;
</div>
</form>
<%
End Sub

'==========================================
'函 数 名SiteSetDo()
'作    用系统设置操作
'参    数
'==========================================
Sub SiteSetDo()
	Dim OldSiteTemplate,ObjFile,ObjFiles
	OldSiteTemplate=SiteTemplate
	SiteName=HTMLEncode(Trim(Request.Form("SiteName")))
	SiteKeyword=HTMLEncode(Trim(Request.Form("SiteKeyword")))
	SiteDescription=HTMLEncode(Trim(Request.Form("SiteDescription")))
	SiteUrl=HTMLEncode(Trim(Request.Form("SiteUrl")))
	SiteTest=1
	'自定义增加部分start---------------------------------
	Tel=HTMLEncode(Trim(Request.Form("Tel")))
	Tel400=HTMLEncode(Trim(Request.Form("Tel400")))
	Fax=HTMLEncode(Trim(Request.Form("Fax")))
	Lianxiren=HTMLEncode(Trim(Request.Form("Lianxiren")))
	Email=HTMLEncode(Trim(Request.Form("Email")))
	Beian=HTMLEncode(Trim(Request.Form("Beian")))
	Add=HTMLEncode(Trim(Request.Form("Add")))
	
	Call FsoLineWrite("inc/Site.asp",7,"SiteName="""&SiteName&"""")
	Call FsoLineWrite("inc/Site.asp",8,"SiteUrl="""&SiteUrl&"""")
	Call FsoLineWrite("inc/Site.asp",9,"SiteKeyword="""&SiteKeyword&"""")
	Call FsoLineWrite("inc/Site.asp",10,"SiteDescription="""&SiteDescription&"""")
	Call FsoLineWrite("inc/Site.asp",11,"SiteOpen="&SiteOpen&"")
	Call FsoLineWrite("inc/Site.asp",12,"SiteTemplate="""&SiteTemplate&"""")
	Call FsoLineWrite("inc/Site.asp",13,"SiteHtml="&SiteHtml&"")
	Call FsoLineWrite("inc/Site.asp",14,"PageSizes="&PageSizes&"")
	Call FsoLineWrite("inc/Site.asp",15,"SiteToPinyin="&SiteToPinyin&"")
	Call FsoLineWrite("inc/Site.asp",16,"FetionNum="""&FetionNum&"""")
	Call FsoLineWrite("inc/Site.asp",17,"FetionPass="""&FetionPass&"""")
	Call FsoLineWrite("inc/Site.asp",18,"SiteQQ="&SiteQQ&"")
	Call FsoLineWrite("inc/Site.asp",19,"SiteNoTrash="&SiteNoTrash&"")
	Call FsoLineWrite("inc/Site.asp",20,"SiteMini="&SiteMini&"")
	Call FsoLineWrite("inc/Site.asp",21,"SiteDelWord="&SiteDelWord&"")
	Call FsoLineWrite("inc/Site.asp",22,"SiteTest="&SiteTest&"")
	Call FsoLineWrite("inc/Site.asp",23,"SiteFlash="&SiteFlash&"")
	Call FsoLineWrite("inc/Site.asp",24,"Tel="""&Tel&"""")
	Call FsoLineWrite("inc/Site.asp",25,"Tel400="""&Tel400&"""")
	Call FsoLineWrite("inc/Site.asp",26,"Fax="""&Fax&"""")
	Call FsoLineWrite("inc/Site.asp",27,"Email="""&Email&"""")
	Call FsoLineWrite("inc/Site.asp",28,"Lianxiren="""&Lianxiren&"""")
	Call FsoLineWrite("inc/Site.asp",29,"Beian="""&Beian&"""")
	Call FsoLineWrite("inc/Site.asp",30,"Add="""&Add&"""")
	Call FsoLineWrite("inc/Site.asp",31,"Tjid="""&Tjid&"""")
	Call FsoLineWrite("inc/Site.asp",32,"Kfid="""&Kfid&"""")
	Call FsoLineWrite("inc/Site.asp",33,"SiteLogo="""&SiteLogo&"""")
	Call FsoLineWrite("inc/Site.asp",34,"Sitepic1="""&Sitepic1&"""")
	Call FsoLineWrite("inc/Site.asp",35,"Sitepic2="""&Sitepic2&"""")
	Call FsoLineWrite("inc/Site.asp",36,"Sitepic3="""&Sitepic3&"""")
	Call FsoLineWrite("inc/Site.asp",37,"Sitepic4="""&Sitepic4&"""")
	Call FsoLineWrite("inc/Site.asp",38,"Sitepic5="""&Sitepic5&"""")
	Call FsoLineWrite("inc/Site.asp",39,"Sitepicurl1="""&Sitepicurl1&"""")
	Call FsoLineWrite("inc/Site.asp",40,"Sitepicurl2="""&Sitepicurl2&"""")
	Call FsoLineWrite("inc/Site.asp",41,"Sitepicurl3="""&Sitepicurl3&"""")
	Call FsoLineWrite("inc/Site.asp",42,"Sitepicurl4="""&Sitepicurl4&"""")
	Call FsoLineWrite("inc/Site.asp",43,"Sitepicurl5="""&Sitepicurl5&"""")
	Call FsoLineWrite("inc/Site.asp",44,"Sitepictext1="""&Sitepictext1&"""")
	Call FsoLineWrite("inc/Site.asp",45,"Sitepictext2="""&Sitepictext2&"""")
	Call FsoLineWrite("inc/Site.asp",46,"Sitepictext3="""&Sitepictext3&"""")
	Call FsoLineWrite("inc/Site.asp",47,"Sitepictext4="""&Sitepictext4&"""")
	Call FsoLineWrite("inc/Site.asp",48,"Sitepictext5="""&Sitepictext5&"""")


	Response.Write("修改成功！")
	Response.redirect "/"
End Sub

'==========================================
'函 数 名TestFetion()
'作    用飞信接口测试
'参    数
'==========================================
Sub TestFetion()
	If FetionNum<>"" And FetionPass<>"" Then
		Temp=FKFun.SmsGo("测试飞信接口，如果您收到本短信则飞信接口已经正常运作！")
		Response.Write(Temp)
	Else
		Response.Write("请先设置飞信号！")
	End If
End Sub


	'==============================
	'函 数 名：HTMLEncode
	'作    用：字符转换函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")		
			fString = Replace(fString, CHR(34), "&quot;")
			fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, CHR(9), "&nbsp;")
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "<p></p> ")
			fString = Replace(fString, CHR(10), "<br /> ")
			HTMLEncode = fString
		End If
	End Function
	
	'==============================
	'函 数 名：FsoLineWrite
	'作    用：按行写入文件
	'参    数：文件相对路径FilePath，写入行号LineNum，写入内容LineContent
	'==============================
	Function FsoLineWrite(FilePath,LineNum,LineContent)
		If LineNum<1 Then Exit Function
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If Not Fso.FileExists(Server.MapPath(FilePath)) Then Exit Function
		Temp=FsoFileRead(FilePath)
		TempArr=Split(Temp,Chr(13)&Chr(10))
		TempArr(LineNum-1)=LineContent
		Temp=Join(TempArr,Chr(13)&Chr(10))
		Call CreateFile(FilePath,Temp)
		Set Fso=Nothing
	End Function
	
	'==============================
	'函 数 名：FsoFileRead
	'作    用：读取文件
	'参    数：文件相对路径FilePath
	'==============================
	Function FsoFileRead(FilePath)
		Set objAdoStream = Server.CreateObject("A"&"dod"&"b.St"&"r"&"eam")
		objAdoStream.Type=2
		objAdoStream.mode=3  
		objAdoStream.charset="utf-8"
		objAdoStream.open 
		objAdoStream.LoadFromFile Server.MapPath(FilePath) 
		FsoFileRead=objAdoStream.ReadText 
		objAdoStream.Close
		Set objAdoStream=Nothing
	End Function
	
	'==============================
	'函 数 名：CreateFolder
	'作    用：创建文件夹
	'参    数：文件夹相对路径FolderPath
	'==============================
	Function CreateFolder(FolderPath)
		If FolderPath<>"" Then
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Set F=Fso.CreateFolder(Server.MapPath(FolderPath))
			CreateFolder=F.Path
			Set F=Nothing
			Set Fso=Nothing
		End If
	End Function
	
	'==============================
	'函 数 名：CreateFile
	'作    用：创建文件
	'参    数：文件相对路径FilePath，文件内容FileContent
	'==============================
	Function CreateFile(FilePath,FileContent)
		Dim Temps
		Temps=""
		TempArr=Split(FilePath,"/")
		For i=0 to UBound(TempArr)-1
			If Temps="" Then
				Temps=TempArr(i)
			Else
				Temps=Temps&"/"&TempArr(i)
			End If
			If IsFolder(Temps)=False Then
				Call CreateFolder(Temps)
			End If
		Next
		Set objAdoStream = Server.CreateObject("A"&"dod"&"b.St"&"r"&"eam")
		objAdoStream.Type = 2
		objAdoStream.Charset = "utf-8" 
		objAdoStream.Open
		objAdoStream.WriteText = FileContent
		objAdoStream.SaveToFile Server.MapPath(FilePath),2
		objAdoStream.Close()
		Set objAdoStream = Nothing
	End Function
	
	'==============================
	'函 数 名：DelFolder
	'作    用：删除文件夹
	'参    数：文件夹相对路径FolderPath
	'==============================
	Function DelFolder(FolderPath)
		If IsFolder(FolderPath)=True Then
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Fso.DeleteFolder(Server.MapPath(FolderPath))
			Set Fso=Nothing
		End If 
	End Function 
	
	'==============================
	'函 数 名：DelFile
	'作    用：删除文件
	'参    数：文件相对路径FilePath
	'==============================
	Function DelFile(FilePath)
		If IsFile(FilePath)=True Then 
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			Fso.DeleteFile(Server.MapPath(FilePath))
			Set Fso=Nothing
		End If
	End Function 
	 
	'==============================
	'函 数 名：IsFile
	'作    用：检测文件是否存在
	'参    数：文件相对路径FilePath
	'==============================
	Function IsFile(FilePath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If (Fso.FileExists(Server.MapPath(FilePath))) Then
			IsFile=True
		Else
			IsFile=False
		End If
		Set Fso=Nothing
	End Function
	
	'==============================
	'函 数 名：IsFolder
	'作    用：检测文件夹是否存在
	'参    数：文件相对路径FolderPath
	'==============================
	Function IsFolder(FolderPath)
		If FolderPath<>"" Then
			Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
			If Fso.FolderExists(Server.MapPath(FolderPath)) Then  
				IsFolder=True
			Else
				IsFolder=False
			End If
			Set Fso=Nothing
		End If
	End Function
	
	'==============================
	'函 数 名：CopyFiles
	'作    用：复制文件
	'参    数：文件来源地址SourcePath，文件复制到地址CopyToPath
	'==============================
	Function CopyFiles(SourcePath,CopyToPath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		Fso.CopyFile Server.MapPath(SourcePath),Server.MapPath(CopyToPath)
		Set Fso=nothing
	End Function
	
	'==============================
	'函 数 名：CopyFolder
	'作    用：复制文件夹
	'参    数：源文件夹FolderName，复制到文件夹FolderPath
	'==============================
	Function CopyFolder(FolderName,FolderPath)
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If Fso.Folderexists(Server.MapPath(FolderName)) Then
			If Fso.FolderExists(Server.MapPath(FolderPath)) Then
				Fso.CopyFolder Server.MapPath(FolderName),Server.MapPath(FolderPath)
			Else
				Fso.CreateFolder(Server.MapPath(FolderPath))
				Fso.CopyFolder Server.MapPath(FolderName),Server.MapPath(FolderPath)
			End if 
		End If 
		Set Fso=nothing
	End Function 
	
%><!--Include File="Code.asp"-->