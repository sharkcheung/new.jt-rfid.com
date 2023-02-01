<%
'==========================================
'文 件 名：Cls_Upload.asp
'文件用途：上传函数类
'==========================================

Class UpLoadClass

	Private p_MaxSize,p_FileType,p_SavePath,p_AutoSave,p_Error
	Private objForm,binForm,binItem,strDate,lngTime
	Public	FormItem,FileItem

	Public Property Get Version
		Version="Rumor UpLoadClass Version 2.0"
	End Property

	Public Property Get Error
		Error=p_Error
	End Property

	Public Property Get MaxSize
		MaxSize=p_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		if isNumeric(lngSize) Then
			p_MaxSize=clng(lngSize)
		End if
	End Property

	Public Property Get FileType
		FileType=p_FileType
	End Property
	Public Property Let FileType(strType)
		p_FileType=strType
	End Property

	Public Property Get SavePath
		SavePath=p_SavePath
	End Property
	Public Property Let SavePath(strPath)
		p_SavePath=replace(strPath,chr(0),"")
	End Property

	Public Property Get AutoSave
		AutoSave=p_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		select case Flag
			case 0:
			case 1:
			case 2:
			case false:Flag=2
			case else:Flag=0
		End select
		p_AutoSave=Flag
	End Property

	Private Sub Class_Initialize
		p_Error	   = -1
		p_MaxSize  = 153600
		p_FileType = "jpg/gif/png/bmp/rar/zip/doc/xls/ppt"
		p_SavePath = ""
		p_AutoSave = 0
		strDate	   = replace(cstr(Date()),"-","")
		strDate	   = replace(strDate,"/","")
		lngTime	   = clng(timer()*1000)
		Set binForm = Server.CreateObject("ADODB.Stream")
		Set binItem = Server.CreateObject("ADODB.Stream")
		Set objForm = Server.CreateObject("Scripting.Dictionary")
		objForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		objForm.RemoveAll
		Set objForm = nothing
		Set binItem = nothing
		binForm.Close()
		Set binForm = nothing
	End Sub

	Public Sub open()
		if p_Error=-1 Then
			p_Error=0
		else
			Exit Sub
		End if
		Dim lngRequestSize,binRequestData,strFormItem,strFileItem
		Const strSplit="'"">"
		lngRequestSize=Request.TotalBytes
		if lngRequestSize<1 Then
			p_Error=4
			Exit Sub
		End if
		if lngRequestSize>p_MaxSize then
			p_Error=4
			Exit Sub
		end if
		binRequestData=Request.BinaryRead(lngRequestSize)
		binForm.Type = 1
		binForm.open
		binForm.Write binRequestData

		Dim bCrLf,strSeparator,intSeparator
		bCrLf=ChrB(13)&ChrB(10)

		intSeparator=InstrB(1,binRequestData,bCrLf)-1
		strSeparator=LeftB(binRequestData,intSeparator)

		Dim p_start,p_end,strItem,strInam,intTemp,strTemp
		Dim strFtyp,strFnam,strFext,lngFsiz
		p_start=intSeparator+2
		Do
			p_end  =InStrB(p_start,binRequestData,bCrLf&bCrLf)+3
			binItem.Type=1
			binItem.open
			binForm.Position=p_start
			binForm.CopyTo binItem,p_end-p_start
			binItem.Position=0
			binItem.Type=2
			binItem.Charset="utf-8"
			strItem=binItem.ReadText
			binItem.Close()

			p_start=p_end
			p_end  =InStrB(p_start,binRequestData,strSeparator)-1
			binItem.Type=1
			binItem.open
			binForm.Position=p_start
			lngFsiz=p_end-p_start-2
			binForm.CopyTo binItem,lngFsiz

			intTemp=Instr(39,strItem,"""")
			strInam=Mid(strItem,39,intTemp-39)

			if Instr(intTemp,strItem,"filename=""")<>0 Then
			if not objForm.Exists(strInam&"_From") Then
				strFileItem=strFileItem&strSplit&strInam
				if binItem.Size<>0 Then
					intTemp=intTemp+13
					strFtyp=Mid(strItem,Instr(intTemp,strItem,"Content-Type: ")+14)
					strTemp=Mid(strItem,intTemp,Instr(intTemp,strItem,"""")-intTemp)
					intTemp=InstrRev(strTemp,"\")
					strFnam=Mid(strTemp,intTemp+1)
					objForm.Add strInam&"_Type",strFtyp
					objForm.Add strInam&"_Name",strFnam
					objForm.Add strInam&"_Path",Left(strTemp,intTemp)
					objForm.Add strInam&"_Size",lngFsiz
					if Instr(strTemp,".")<>0 Then
						strFext=Mid(strTemp,InstrRev(strTemp,".")+1)
					else
						strFext=""
					End if
					if left(strFtyp,6)="image/" Then
						binItem.Position=0
						binItem.Type=1
						strTemp=binItem.read(10)
						if strcomp(strTemp,chrb(255) & chrb(216) & chrb(255) & chrb(224) & chrb(0) & chrb(16) & chrb(74) & chrb(70) & chrb(73) & chrb(70),0)=0 Then
							if Lcase(strFext)<>"jpg" Then strFext="jpg"
							binItem.Position=3
							do while not binItem.EOS
								do
									intTemp = ascb(binItem.Read(1))
								loop while intTemp = 255 and not binItem.EOS
								if intTemp < 192 or intTemp > 195 Then
									binItem.read(Bin2Val(binItem.Read(2))-2)
								else
									Exit do
								End if
								do
									intTemp = ascb(binItem.Read(1))
								loop while intTemp < 255 and not binItem.EOS
							loop
							binItem.Read(3)
							objForm.Add strInam&"_Height",Bin2Val(binItem.Read(2))
							objForm.Add strInam&"_Width",Bin2Val(binItem.Read(2))
						elseif strcomp(leftB(strTemp,8),chrb(137) & chrb(80) & chrb(78) & chrb(71) & chrb(13) & chrb(10) & chrb(26) & chrb(10),0)=0 Then
							if Lcase(strFext)<>"png" Then strFext="png"
							binItem.Position=18
							objForm.Add strInam&"_Width",Bin2Val(binItem.Read(2))
							binItem.Read(2)
							objForm.Add strInam&"_Height",Bin2Val(binItem.Read(2))
						elseif strcomp(leftB(strTemp,6),chrb(71) & chrb(73) & chrb(70) & chrb(56) & chrb(57) & chrb(97),0)=0 or strcomp(leftB(strTemp,6),chrb(71) & chrb(73) & chrb(70) & chrb(56) & chrb(55) & chrb(97),0)=0 Then
							if Lcase(strFext)<>"gif" Then strFext="gif"
							binItem.Position=6
							objForm.Add strInam&"_Width",BinVal2(binItem.Read(2))
							objForm.Add strInam&"_Height",BinVal2(binItem.Read(2))
						elseif strcomp(leftB(strTemp,2),chrb(66) & chrb(77),0)=0 Then
							if Lcase(strFext)<>"bmp" Then strFext="bmp"
							binItem.Position=18
							objForm.Add strInam&"_Width",BinVal2(binItem.Read(4))
							objForm.Add strInam&"_Height",BinVal2(binItem.Read(4))
						End if
					End if
					objForm.Add strInam&"_Ext",strFext
					objForm.Add strInam&"_From",p_start
					intTemp=GetFerr(lngFsiz,strFext)
					if p_AutoSave<>2 Then
						objForm.Add strInam&"_Err",intTemp
						if intTemp=0 Then
							if p_AutoSave=0 Then
								strFnam=GetTimeStr()
								if strFext<>"" Then strFnam=strFnam&"."&strFext
							End if
							binItem.SaveToFile Server.MapPath(p_SavePath&strFnam),2
							objForm.Add strInam,strFnam
						End if
					End if
				else
					objForm.Add strInam&"_Err",-1
				End if
			End if
			else
				binItem.Position=0
				binItem.Type=2
				binItem.Charset="utf-8"
				strTemp=binItem.ReadText
				if objForm.Exists(strInam) Then
					objForm(strInam) = objForm(strInam)&","&strTemp
				else
					strFormItem=strFormItem&strSplit&strInam
					objForm.Add strInam,strTemp
				End if
			End if

			binItem.Close()
			p_start = p_end+intSeparator+2
		loop Until p_start+3>lngRequestSize
		FormItem=split(strFormItem,strSplit)
		FileItem=split(strFileItem,strSplit)
	End Sub

	Private Function GetTimeStr()
		lngTime=lngTime+1
		GetTimeStr=strDate&lngTime
	End Function

	Private Function GetFerr(lngFsiz,strFext)
		Dim intFerr
		intFerr=0
		if lngFsiz>p_MaxSize and p_MaxSize>0 Then
			if p_Error=0 or p_Error=2 Then p_Error=p_Error+1
			intFerr=intFerr+1
		End if
		if Instr(1,LCase("/"&p_FileType&"/"),LCase("/"&strFext&"/"))=0 and p_FileType<>"" Then
			if p_Error<2 Then p_Error=p_Error+2
			intFerr=intFerr+2
		End if
		GetFerr=intFerr
	End Function

	Public Function Save(Item,strFnam)
		Save=false
		if objForm.Exists(Item&"_From") Then
			Dim intFerr,strFext
			strFext=objForm(Item&"_Ext")
			intFerr=GetFerr(objForm(Item&"_Size"),strFext)
			if objForm.Exists(Item&"_Err") Then
				if intFerr=0 Then
					objForm(Item&"_Err")=0
				End if
			else
				objForm.Add Item&"_Err",intFerr
			End if
			if intFerr<>0 Then Exit Function
			if VarType(strFnam)=2 Then
				select case strFnam
					case 0:strFnam=GetTimeStr()
						if strFext<>"" Then strFnam=strFnam&"."&strFext
					case 1:strFnam=objForm(Item&"_Name")
				End select
			End if
			binItem.Type = 1
			binItem.open
			binForm.Position = objForm(Item&"_From")
			binForm.CopyTo binItem,objForm(Item&"_Size")
			binItem.SaveToFile Server.MapPath(p_SavePath&strFnam),2
			binItem.Close()
			if objForm.Exists(Item) Then
				objForm(Item)=strFnam
			else
				objForm.Add Item,strFnam
			End if
			Save=true
		End if
	End Function

	Public Function GetData(Item)
		GetData=""
		if objForm.Exists(Item&"_From") Then
			if GetFerr(objForm(Item&"_Size"),objForm(Item&"_Ext"))<>0 Then Exit Function
			binForm.Position = objForm(Item&"_From")
			GetData=binFormStream.Read(objForm(Item&"_Size"))
		End if
	End Function

	Public Function Form(Item)
		if objForm.Exists(Item) Then
			Form=objForm(Item)
		else
			Form=""
		End if
	End Function

	Private Function BinVal2(bin)
		Dim lngValue,i
		lngValue = 0
		for i = lenb(bin) to 1 step -1
			lngValue = lngValue *256 + ascb(midb(bin,i,1))
		next
		BinVal2=lngValue
	End Function

	Private Function Bin2Val(bin)
		Dim lngValue,i
		lngValue = 0
		for i = 1 to lenb(bin)
			lngValue = lngValue *256 + ascb(midb(bin,i,1))
		next
		Bin2Val=lngValue
	End Function

	'==========================================
	'Add   :  Shark
	'Time   :  2013-11-16 1:09
	'Descrip:  判断上传文件是否非法
	'==========================================
	Function CheckFileContent(FileName)
		Dim ClientFile, ClientText, ClientContent, DangerString, DSArray, AttackFlag, k
		Set ClientFile = Server.CreateObject("Scripting.FileSystemObject")
		Set ClientText = ClientFile.OpenTextFile(Server.MapPath(FileName), 1)
		ClientContent = LCase(ClientText.ReadAll)
		Set ClientText = Nothing
		Set ClientFile = Nothing
		AttackFlag = False
		DangerString = ".getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language=|include|filesystemobject|shell.application|eval|request"
		DSArray = Split(DangerString, "|")
		For k = 0 To UBound(DSArray)
			If InStr(ClientContent, DSArray(k))>0 Then '判断文件内容中是否包含有危险的操作字符，如有，则必须删除该文件。
				AttackFlag = True
				Exit For
			End If
		Next
		CheckFileContent = AttackFlag
	End Function	
End Class
%>
