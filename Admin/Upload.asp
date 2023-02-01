<!--#Include File="AdminCheck.asp"--><%

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

Response.Write UploadFile("FileData")

function uploadfile(inputname)
	dim immediate,attachdir,dirtype,maxattachsize,upext
	immediate=Request.QueryString("immediate")
	attachdir=SiteDir&"Up"'上传文件保存路径，结尾不要带/
	dirtype=1'1:按天存入目录 2:按月存入目录 3:按扩展名存目录  建议使用按天存
	maxattachsize=2097152'最大上传大小，默认是2M
	upext="txt,rar,zip,jpg,jpeg,gif,png,swf,wmv,avi,wma,mp3,mid,xml"'上传扩展名
	
	dim err,msg,upfile
	err = ""
	msg = ""
	
	set upfile=new upfile_class
	upfile.AllowExt=replace(upext,",",";")+";"
	upfile.GetData(maxattachsize)
	if upfile.isErr then
		select case upfile.isErr
		case 1
			err="无数据提交"
		case 2
			err="文件大小超过 "+cstr(maxattachsize)+"字节"
		case else
			err=upfile.ErrMessage
		end select
	else
		dim attach_dir,attach_subdir,filename,extension,target,tmpfile
		extension=upfile.file(inputname).FileExt
		select case dirtype
			case 1
				attach_subdir="day_"+DateFormat(now,"yymmdd")
			case 2
				attach_subdir="month_"+DateFormat(now,"yymm")
			case 3
				attach_subdir="ext_"+extension
		end select
		attach_dir=attachdir+"/"+attach_subdir+"/"
		'建文件夹
		CreateFolder attach_dir
		tmpfile=upfile.AutoSave(inputname,Server.mappath(attach_dir)+"\")
		if upfile.isErr then
			if upfile.isErr=3 then
				err="上传文件扩展名必需为："+upext
			else
				err=upfile.ErrMessage
			end if
		else
			'生成随机文件名并改名
			Randomize timer
			filename=DateFormat(now,"yyyymmddhhnnss")+cstr(cint(9999*Rnd))+"."+extension
			target=attach_dir+filename
			moveFile attach_dir+tmpfile,target
			msg=target
			Pic=target
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
						If Clng(TWidth)>Clng(TempArr(17)) Then
							THeight=TempArr(17)/TWidth*THeight
							TWidth=TempArr(17)
						End If
						If Clng(THeight)>Clng(TempArr(18)) Then
							TWidth=TempArr(18)/THeight*TWidth
							THeight=TempArr(18)
						End If
						Call FKFun.DoSmall(Pic,SmallPic,TWidth,THeight)
					End If
				Else
					SmallPic=Pic
				End If
				If Fk_Jpeg_Water=1 Then
					Call FKFun.DoWater(Pic,Pic,TempArr(0),TempArr(1),TempArr(2),TempArr(3),TempArr(4),TempArr(5))
					If SmallPic<>Pic Then
						Call FKFun.DoWater(SmallPic,SmallPic,TempArr(0),TempArr(1),TempArr(2),TempArr(3),TempArr(4),TempArr(5))
					End If
				End If
			Else
				SmallPic=Pic
			End If
			if immediate="1" then msg="!"+SmallPic+"||"+Pic
		end if
	end if
	set upfile=nothing
	uploadfile="{err:'"+jsonString(err)+"',msg:'"+jsonString(msg)+"'}"
end function

function jsonString(str)
	str=replace(str,"\","\\")
	str=replace(str,"/","\/")
	str=replace(str,"'","\'")
	jsonString=str
end function

Function Iif(expression,returntrue,returnfalse)
	If expression=true Then
		iif=returntrue
	Else
		iif=returnfalse
	End If
End Function

function DateFormat(strDate,fstr)
	if isdate(strDate) then
		dim i,temp
		temp=replace(fstr,"yyyy",DatePart("yyyy",strDate))
		temp=replace(temp,"yy",mid(DatePart("yyyy",strDate),3))
		temp=replace(temp,"y",DatePart("y",strDate))
		temp=replace(temp,"w",DatePart("w",strDate))
		temp=replace(temp,"ww",DatePart("ww",strDate))
		temp=replace(temp,"q",DatePart("q",strDate))
		temp=replace(temp,"mm",iif(len(DatePart("m",strDate))>1,DatePart("m",strDate),"0"&DatePart("m",strDate)))
		temp=replace(temp,"dd",iif(len(DatePart("d",strDate))>1,DatePart("d",strDate),"0"&DatePart("d",strDate)))
		temp=replace(temp,"hh",iif(len(DatePart("h",strDate))>1,DatePart("h",strDate),"0"&DatePart("h",strDate)))
		temp=replace(temp,"nn",iif(len(DatePart("n",strDate))>1,DatePart("n",strDate),"0"&DatePart("n",strDate)))
		temp=replace(temp,"ss",iif(len(DatePart("s",strDate))>1,DatePart("s",strDate),"0"&DatePart("s",strDate)))
		DateFormat=temp
	else
		DateFormat=false
	end if
end function

Function CreateFolder(FolderPath)
	dim lpath,fs,f
	lpath=Server.MapPath(FolderPath)
	Set fs=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	If not fs.FolderExists(lpath) then
		Set f=fs.CreateFolder(lpath)
		CreateFolder=F.Path
	end if
	Set F=Nothing
	Set fs=Nothing
End Function

Function moveFile(oldfile,newfile)
	dim fs
	Set fs=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	fs.movefile Server.MapPath(oldfile),Server.MapPath(newfile)
	Set fs=Nothing
End Function
 
'----------------------------------------------------------------------
'转发时请保留此声明信息,这段声明不并会影响你的速度!
'*******************	 无惧上传类 V2.2	************************************
'作者:梁无惧
'网站:http://www.25cn.com
'电子邮件:yjlrb@21cn.com
'版权声明:版权所有,源代码公开,各种用途均可免费使用,但是修改后必须把修改后的文件
'发送一份给作者.并且保留作者此版权信息
'**********************************************************************
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'文件上传类
Class UpFile_Class

Dim Form,File
Dim AllowExt_	'允许上传类型(白名单)
Dim NoAllowExt_	'不允许上传类型(黑名单)
Dim IsDebug_ '是否显示出错信息
Private	oUpFileStream	'上传的数据流
Private isErr_		'错误的代码,0或true表示无错
Private ErrMessage_	'错误的字符串信息
Private isGetData_	'指示是否已执行过GETDATA过程

'------------------------------------------------------------------
'类的属性
Public Property Get Version
	Version="无惧上传类 Version V2.0"
End Property

Public Property Get isErr		'错误的代码,0或true表示无错
	isErr=isErr_
End Property

Public Property Get ErrMessage		'错误的字符串信息
	ErrMessage=ErrMessage_
End Property

Public Property Get AllowExt		'允许上传类型(白名单)
	AllowExt=AllowExt_
End Property

Public Property Let AllowExt(Value)	'允许上传类型(白名单)
	AllowExt_=LCase(Value)
End Property

Public Property Get NoAllowExt		'不允许上传类型(黑名单)
	NoAllowExt=NoAllowExt_
End Property

Public Property Let NoAllowExt(Value)	'不允许上传类型(黑名单)
	NoAllowExt_=LCase(Value)
End Property

Public Property Let IsDebug(Value)	'是否设置为调试模式
	IsDebug_=Value
End Property


'----------------------------------------------------------------
'类实现代码

'初始化类
Private Sub Class_Initialize
	isErr_ = 0
	NoAllowExt=""		'黑名单,可以在这里预设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名单
	NoAllowExt=LCase(NoAllowExt)
	AllowExt=""		'白名单,可以在这里预设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
	AllowExt=LCase(AllowExt)
	isGetData_=false
End Sub

'类结束
Private Sub Class_Terminate	
	on error Resume Next
	'清除变量及对像
	Form.RemoveAll
	Set Form = Nothing
	File.RemoveAll
	Set File = Nothing
	oUpFileStream.Close
	Set oUpFileStream = Nothing
	if Err.number<>0 then OutErr("清除类时发生错误!")
End Sub

'分析上传的数据
Public Sub GetData (MaxSize)
	 '定义变量
	on error Resume Next
	if isGetData_=false then 
		Dim RequestBinDate,sSpace,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
		Dim sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		'代码开始
		If Request.TotalBytes < 1 Then	'如果没有数据上传
			isErr_ = 1
			ErrMessage_="没有数据上传,这是因为直接提交网址所产生的错误!"
			OutErr("没有数据上传,这是因为直接提交网址所产生的错误!!")
			Exit Sub
		End If
		If MaxSize > 0 Then '如果限制大小
			If Request.TotalBytes > MaxSize Then
			isErr_ = 2	'如果上传的数据超出限制大小
			ErrMessage_="上传的数据超出限制大小!"
			OutErr("上传的数据超出限制大小!")
			Exit Sub
			End If
		End If
		Set Form = Server.CreateObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = Server.CreateObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = Server.CreateObject ("ADODB.Stream")
		Set oUpFileStream = Server.CreateObject ("ADODB.Stream")
		if Err.number<>0 then OutErr("创建流对象(ADODB.STREAM)时出错,可能系统不支持或没有开通该组件")
		oUpFileStream.Type = 1
		oUpFileStream.Mode = 3
		oUpFileStream.Open 
		oUpFileStream.Write Request.BinaryRead (Request.TotalBytes)
		oUpFileStream.Position = 0
		RequestBinDate = oUpFileStream.Read 
		iFormEnd = oUpFileStream.Size
		bCrLf = ChrB (13) & ChrB (10)
		'取得每个项目之间的分隔符
		sSpace = MidB (RequestBinDate,1, InStrB (1,RequestBinDate,bCrLf)-1)
		iStart = LenB(sSpace)
		iFormStart = iStart+2
		'分解项目
		Do
			iInfoEnd = InStrB (iFormStart,RequestBinDate,bCrLf & bCrLf)+3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			oUpFileStream.Position = iFormStart
			oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = "gb2312"
			sInfo = tStream.ReadText			
			'取得表单项目名称
			iFormStart = InStrB (iInfoEnd,RequestBinDate,sSpace)-1
			iFindStart = InStr (22,sInfo,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sInfo,"""",1)
			sFormName = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
			'如果是文件
			If InStr (45,sInfo,"filename=""",1) > 0 Then
				Set oFileInfo = new FileInfo_Class
				'取得文件属性
				iFindStart = InStr (iFindEnd,sInfo,"filename=""",1)+10
				iFindEnd = InStr (iFindStart,sInfo,""""&vbCrLf,1)
				sFileName = Trim(Mid(sinfo,iFindStart,iFindEnd-iFindStart))
				oFileInfo.FileName = GetFileName(sFileName)
				oFileInfo.FilePath = GetFilePath(sFileName)
				oFileInfo.FileExt = GetFileExt(sFileName)
				iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
				iFindEnd = InStr (iFindStart,sInfo,vbCr)
				oFileInfo.FileMIME = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
				oFileInfo.FileStart = iInfoEnd
				oFileInfo.FileSize = iFormStart -iInfoEnd -2
				oFileInfo.FormName = sFormName
				file.add sFormName,oFileInfo
			else
			'如果是表单项目
				tStream.Close
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open
				oUpFileStream.Position = iInfoEnd 
				oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
				tStream.Position = 0
				tStream.Type = 2
				tStream.CharSet = "gb2312"
				sFormValue = tStream.ReadText
				If Form.Exists (sFormName) Then
					Form (sFormName) = Form (sFormName) & ", " & sFormValue
					else
					Form.Add sFormName,sFormValue
				End If
			End If
			tStream.Close
			iFormStart = iFormStart+iStart+2
			'如果到文件尾了就退出
		Loop Until (iFormStart+2) >= iFormEnd 
		if Err.number<>0 then OutErr("分解上传数据时发生错误,可能客户端的上传数据不正确或不符合上传数据规则")
		RequestBinDate = ""
		Set tStream = Nothing
		isGetData_=true
	end if
End Sub

'保存到文件,自动覆盖已存在的同名文件
Public Function SaveToFile(Item,Path)
	SaveToFile=SaveToFileEx(Item,Path,True)
End Function

'保存到文件,自动设置文件名
Public Function AutoSave(Item,Path)
	AutoSave=SaveToFileEx(Item,Path,false)
End Function

'保存到文件,OVER为真时,自动覆盖已存在的同名文件,否则自动把文件改名保存
Private Function SaveToFileEx(Item,Path,Over)
	On Error Resume Next
	Dim FileExt
	if file.Exists(Item) then
		Dim oFileStream
		Dim tmpPath
		isErr_=0
		Set oFileStream = CreateObject ("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		oUpFileStream.Position = File(Item).FileStart
		oUpFileStream.CopyTo oFileStream,File(Item).FileSize
		tmpPath=Split(Path,".")(0)
		FileExt=GetFileExt(Path)
		if Over then
			if isAllowExt(FileExt) then
				oFileStream.SaveToFile tmpPath & "." & FileExt,2
				if Err.number<>0 then OutErr("保存文件时出错,请检查路径,是否存在该上传目录!该文件保存路径为" & tmpPath & "." & FileExt)
			Else
				isErr_=3
				ErrMessage_="该后缀名的文件不允许上传!"
				OutErr("该后缀名的文件不允许上传")
			End if
		Else
			Path=GetFilePath(Path)
			dim fori
			fori=1
			if isAllowExt(File(Item).FileExt) then
				do
					fori=fori+1
					Err.Clear()
					tmpPath=Path&GetNewFileName()&"."&File(Item).FileExt
					oFileStream.SaveToFile tmpPath
				loop Until ((Err.number=0) or (fori>50))
				if Err.number<>0 then OutErr("自动保存文件出错,已经测试50次不同的文件名来保存,请检查目录是否存在!该文件最后一次保存时全路径为"&Path&GetNewFileName()&"."&File(Item).FileExt)
			Else
				isErr_=3
				ErrMessage_="该后缀名的文件不允许上传!"
				OutErr("该后缀名的文件不允许上传")
			End if
		End if
		oFileStream.Close
		Set oFileStream = Nothing
	else
		ErrMessage_="不存在该对象(如该文件没有上传,文件为空)!"
		OutErr("不存在该对象(如该文件没有上传,文件为空)")
	end if
	if isErr_=3 then SaveToFileEx="" else SaveToFileEx=GetFileName(tmpPath)
End Function

'取得文件数据
Public Function FileData(Item)
	isErr_=0
	if file.Exists(Item) then
		if isAllowExt(File(Item).FileExt) then
			oUpFileStream.Position = File(Item).FileStart
			FileData = oUpFileStream.Read (File(Item).FileSize)
			Else
			isErr_=3
			ErrMessage_="该后缀名的文件不允许上传"
			OutErr("该后缀名的文件不允许上传")
			FileData=""
		End if
	else
		ErrMessage_="不存在该对象(如该文件没有上传,文件为空)!"
		OutErr("不存在该对象(如该文件没有上传,文件为空)")
	end if
End Function


'取得文件路径
Public function GetFilePath(FullPath)
  If FullPath <> "" Then
    GetFilePath = Left(FullPath,InStrRev(FullPath, "\"))
    Else
    GetFilePath = ""
  End If
End function

'取得文件名
Public Function GetFileName(FullPath)
  If FullPath <> "" Then
    GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
    Else
    GetFileName = ""
  End If
End function

'取得文件的后缀名
Public Function GetFileExt(FullPath)
  If FullPath <> "" Then
    GetFileExt = LCase(Mid(FullPath,InStrRev(FullPath, ".")+1))
    Else
    GetFileExt = ""
  End If
End function

'取得一个不重复的序号
Public Function GetNewFileName()
	dim ranNum
	dim dtNow
	dtNow=Now()
	randomize
	ranNum=int(90000*rnd)+10000
	'以下这段由webboy提供
	GetNewFileName=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum
End Function

Public Function isAllowExt(Ext)
	if NoAllowExt="" then
		isAllowExt=cbool(InStr(1,";"&AllowExt&";",LCase(";"&Ext&";")))
	else
		isAllowExt=not CBool(InStr(1,";"&NoAllowExt&";",LCase(";"&Ext&";")))
	end if
End Function
End Class

Public Sub OutErr(ErrMsg)
if IsDebug_=true then
	Response.Write ErrMsg
	Response.End
	End if
End Sub

'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileInfo_Class
Dim FormName,FileName,FilePath,FileSize,FileMIME,FileStart,FileExt
End Class
%>