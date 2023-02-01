<%
'==========================================
'文 件 名：Cls_Fso.asp
'文件用途：常规函数类
'==========================================

Class Cls_Fso
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
	'函 数 名：FsoLineWriteVar
	'作    用：按行写入变量
	'参    数：文件相对路径FilePath，写入行号LineNum，写入内容LineContent
	'==============================
	Function FsoLineWriteVer(FilePath,LineNum,LineContent)
		If LineNum<1 Then Exit Function
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		If Not Fso.FileExists(Server.MapPath(FilePath)) Then Exit Function
		Temp=FsoFileRead(FilePath)
		TempArr=Split(Temp,Chr(13)&Chr(10))
		TempArr(LineNum-1)=LineContent
		Temp=Join(TempArr,Chr(13)&Chr(10))
		Temp=Temp&Chr(13)&Chr(10)
		If InStr(Temp,"%"&">")=0 then Temp=Temp&"%"&">"
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
End Class
%>
