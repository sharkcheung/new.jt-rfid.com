<%
response.charset="UTF-8"
session.codepage=65001
Rem #####################################################################################
 Rem ## 在线升级类声明
 Class Cls_oUpdate
  Rem #################################################################
  Rem ## 描述: ASP 在线升级类
  Rem ## 版本: 1.0.0
  Rem ## 作者: 萧月痕
  Rem ## MSN:  xiaoyuehen(at)msn.com
  Rem ## 请将(at)以 @ 替换
  Rem ## 版权: 既然共享, 就无所谓版权了. 但必须限于网络传播, 不得用于传统媒体!
  Rem ## 如果您能保留这些说明信息, 本人更加感谢!
  Rem ## 如果您有更好的代码优化, 相关改进, 请记得告诉我, 非常谢谢!
  Rem #################################################################
  Public LocalVersion, LastVersion, FileType
  Public UrlVersion, UrlUpdate, UpdateLocalPath, Info,LogPath
  Public UrlHistory
  Private sstrVersionList, sarrVersionList, sintLocalVersion, sstrLocalVersion
  Private sstrLogContent, sstrHistoryContent, sstrUrlUpdate, sstrUrlLocal
  Rem #################################################################
  Private Sub Class_Initialize()
   Rem ## 版本信息完整URL, 以 http:// 起头
   Rem ## 例: http://localhost/software/Version.htm
   UrlVersion     = ""
   
   Rem ## 升级URL, 以 http:// 起头, /结尾
   Rem ## 例: http://localhost/software/
   UrlUpdate     = ""
   
   Rem ## 本地更新目录, 以 / 起头, /结尾. 以 / 起头是为当前站点更新.防止写到其他目录.
   Rem ## 程序将检测目录是否存在, 不存在则自动创建
   UpdateLocalPath  = "/"
   
   Rem ## 生成的软件历史文件
   UrlHistory     = "admin/LogFile/history.htm"
   
   Rem ## 最后的提示信息
   Info			= ""
   
   '日志目录
   LogPath		= "admin/LogFile/"

   Rem ## 当前版本
   LocalVersion    = "1.0.1"
   
   Rem ## 最新版本
   LastVersion    = "1.0.1"
   
   Rem ## 各版本信息文件后缀名
   FileType      = ".asp"
  End Sub
  Rem #################################################################
  
  Rem #################################################################
  Private Sub Class_Terminate()
  
  End Sub
  Rem #################################################################
  Rem ## 执行升级动作
  Rem #################################################################
  Public function doUpdate()
   doUpdate		= False
   
   LogPath		= Trim(LogPath)
   UrlVersion   = Trim(UrlVersion)
   UrlUpdate    = Trim(UrlUpdate)
   
   Rem ## 升级网址检测
   If (Left(UrlVersion, 7) <> "http://") Or (Left(UrlUpdate, 7) <> "http://") Then
    Info = "版本检测网址为空, 升级网址为空或格式错误(#1)"
    Exit function
   End If
   
   If Right(UrlUpdate, 1) <> "/" Then 
    sstrUrlUpdate = UrlUpdate & "/"
   Else
    sstrUrlUpdate = UrlUpdate
   End If
   
   If Right(UpdateLocalPath, 1) <> "/" Then 
    sstrUrlLocal = UpdateLocalPath & "/"
   Else
    sstrUrlLocal = UpdateLocalPath
   End If   
   
   Rem ## 当前版本信息(数字)
   sstrLocalVersion = LocalVersion
   sintLocalVersion = Replace(sstrLocalVersion, ".", "")
   sintLocalVersion = toNum(sintLocalVersion, 0)
   
   Rem ## 版本检测(初始化版本信息, 并进行比较)
   If IsLastVersion Then Exit function
   
   Rem ## 开始升级
   doUpdate = NowUpdate()
   LastVersion = sstrLocalVersion
  End function
  Rem #################################################################
  
  Rem ## 检测是否为最新版本
  Rem #################################################################
   Private function IsLastVersion()
    Rem ## 初始化版本信息(初始化 sarrVersionList 数组)
    If iniVersionList Then
     Rem ## 若成功, 则比较版本
     Dim i
     IsLastVersion = True
     For i = 0 to UBound(sarrVersionList)
      If sarrVersionList(i) > sintLocalVersion Then
       Rem ## 若有最新版本, 则退出循环
       IsLastVersion = False
       Info = "已经是最新版本!"
       Exit For
      End If
     Next
    Else
     Rem ## 否则返回出错信息
     IsLastVersion = True
     Info = "获取版本信息时出错!(#2)"
    End If   
   End function
  Rem #################################################################
  Rem ## 检测是否为最新版本
  Rem #################################################################
   Private function iniVersionList()
    iniVersionList = False
    
    Dim strVersion
    strVersion = getVersionList()
    
    Rem ## 若返回值为空, 则初始化失败
    If strVersion = "" Then
     Info = "出错......."
     Exit function
    End If
    
    sstrVersionList = Replace(strVersion, " ", "")
    sarrVersionList = Split(sstrVersionList, vbCrLf)
    
    iniVersionList = True
   End function
  Rem #################################################################
  Rem ## 检测是否为最新版本
  Rem #################################################################
   Private function getVersionList()
    getVersionList = GetContent(UrlVersion)
   End function
  Rem #################################################################
  Rem ## 开始更新
  Rem #################################################################
   Private function NowUpdate()
    Dim i
    For i = UBound(sarrVersionList) to 0 step -1
     Call doUpdateVersion(sarrVersionList(i))
    Next
    Info = "升级完成! <a href=""" & sstrUrlLocal & UrlHistory & """>查看</a>"
   End function
  Rem #################################################################
  
  Rem ## 更新版本内容
  Rem #################################################################
   Private function doUpdateVersion(strVer)
    doUpdateVersion = False
    
    Dim intVer
    intVer = toNum(Replace(strVer, ".", ""), 0)
    
    Rem ## 若将更新的版本小于当前版本, 则退出更新
    If intVer <= sintLocalVersion Then
     Exit function
    End If
    
    Dim strFileListContent, arrFileList, strUrlUpdate   
    strUrlUpdate = sstrUrlUpdate & intVer & FileType
    
    strFileListContent = GetContent(strUrlUpdate)
    
    If strFileListContent = "" Then
     Exit function
    End If
    
    Rem ## 更新当前版本号
    sintLocalVersion = intVer
    sstrLocalVersion = strVer
    
    Dim i, arrTmp
    Rem ## 获取更新文件列表
    arrFileList = Split(strFileListContent, vbCrLf)
    
    Rem ## 更新日志
    sstrLogContent = ""
    sstrLogContent = sstrLogContent & strVer & ":" & vbCrLf
    
    Rem ## 开始更新
    For i = 0 to UBound(arrFileList)
     Rem ## 更新格式: 版本号/文件.htm|目的文件
     arrTmp = Split(arrFileList(i), "|")
     sstrLogContent = sstrLogContent & vbTab & arrTmp(1)
     Call doUpdateFile(intVer & "/" & arrTmp(0), arrTmp(1))     
    Next
    
    Rem ## 写入日志文件
    sstrLogContent = sstrLogContent & Now() & vbCrLf
    response.Write("<pre>" & sstrLogContent & "</pre>")
    Call sDoCreateFile(Server.MapPath(sstrUrlLocal & LogPath &"Log" & intVer & ".htm"), _
                                          "<pre>" & sstrLogContent & "</pre>")
    Call sDoAppendFile(Server.MapPath(sstrUrlLocal & UrlHistory), "<pre>" & _
                                          strVer & "_______" & Now() & "</pre>" & vbCrLf)
   End function
  Rem #################################################################
  
  Rem ## 更新文件
  Rem #################################################################
   Private function doUpdateFile(strSourceFile, strTargetFile)
    Dim strContent
    strContent = GetContent(sstrUrlUpdate & strSourceFile)
    
    Rem ## 更新并写入日志
    If sDoCreateFile(Server.MapPath(sstrUrlLocal & strTargetFile), strContent) Then     
     sstrLogContent = sstrLogContent & "  成功" & vbCrLf
    Else
     sstrLogContent = sstrLogContent & "  失败" & vbCrLf
    End If
   End function
  Rem #################################################################
  Rem ## 远程获得内容
  Rem #################################################################
   Private function GetContent(strUrl)
    GetContent = ""
    
    Dim oXhttp, strContent
    Set oXhttp = Server.CreateObject("Microsoft.XMLHTTP")
    'On Error Resume Next 
    With oXhttp
     .Open "GET", strUrl, False, "", ""
     .Send
     If .readystate <> 4 Then Exit function
     strContent = .Responsebody
     
     strContent = sBytesToBstr(strContent)
    End With
    
    Set oXhttp = Nothing
    If Err.Number <> 0 Then
     response.Write(Err.Description)
     Err.Clear
     Exit function
    End If
    
    GetContent = strContent
   End function
  Rem #################################################################
  Rem #################################################################
  Rem ## 编码转换 2进制 => 字符串
   Private function sBytesToBstr(vIn)
    dim objStream
    set objStream = Server.CreateObject("adodb.stream")
    objStream.Type    = 1
    objStream.Mode    = 3
    objStream.Open
    objStream.Write vIn
    
    objStream.Position  = 0
    objStream.Type    = 2
    objStream.Charset  = "UTF-8"
    sBytesToBstr     = objStream.ReadText 
    objStream.Close
    set objStream    = nothing
   End function
  Rem #################################################################
  Rem #################################################################
  Rem ## 编码转换 2进制 => 字符串
   Private function sDoCreateFile1(strFileName, ByRef strContent)
    sDoCreateFile1 = False
    Dim strPath
    strPath = Left(strFileName, InstrRev(strFileName, "\", -1, 1))
    Rem ## 检测路径及文件名有效性
    If Not(CreateMultiFolder(strPath)) Then Exit function
    'If Not(CheckFileName(strFileName)) Then Exit function
    
    'response.Write(strFileName)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(strFileName, ForWriting, True)
    f.Write strContent
    f.Close
    Set fso = nothing
    Set f = nothing
    sDoCreateFile1 = True
   End Function
   
   '================================================
 '函数名：CreatedTextFiles
 '作  用：创建文本文件
 '参  数：filename  ----文件名
 '        body  ----主要内容
 '================================================
 Public Function sDoCreateFile(ByVal FileName, ByVal body)
    sDoCreateFile = False
    Dim strPath
    strPath = Left(FileName, InstrRev(FileName, "\", -1, 1))
    Rem ## 检测路径及文件名有效性
    If Not(CreateMultiFolder(strPath)) Then Exit function
  On Error Resume Next
  Dim oStream
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Type = 2 '设置为可读可写
  oStream.Mode = 3 '设置内容为文本
  oStream.Charset = "UTF-8"
  oStream.Open
  oStream.Position = oStream.Size
  oStream.WriteText body
  oStream.SaveToFile FileName, 2
  oStream.Close
  Set oStream = Nothing
  If Err.Number = 0 Then sDoCreateFile = true
 End Function
  Rem #################################################################
  Rem #################################################################
  Rem ## 编码转换 2进制 => 字符串
   Private function sDoAppendFile(strFileName, ByRef strContent)
    sDoAppendFile = False
    Dim strPath
    strPath = Left(strFileName, InstrRev(strFileName, "\", -1, 1))
    Rem ## 检测路径及文件名有效性
    If Not(CreateMultiFolder(strPath)) Then Exit function
    'If Not(CheckFileName(strFileName)) Then Exit function
    
    'response.Write(strFileName)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(strFileName, ForAppending, True)
    f.Write strContent
    f.Close
    Set fso = nothing
    Set f = nothing
    sDoAppendFile = True
   End Function
   
'创建多级目录，可以创建不存在的根目录
'参数：要创建的目录名称，可以是多级
'返回逻辑值，True成功，False失败
'创建目录的根目录从当前目录开始
'---------------------------------------------------
Function CreateMultiFolder(ByVal CFolder)
Dim objFSO,PhCreateFolder,CreateFolderArray,CreateFolder,rootpath
Dim i,ii,CreateFolderSub,PhCreateFolderSub,BlInfo
BlInfo = False
CreateFolder = CFolder
rootpath=server.mappath("/")
CreateFolder=Replace(CreateFolder,rootpath,"")
CreateFolder=Replace(CreateFolder,"\","/")
'On Error Resume Next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If Err Then
Err.Clear()
Exit Function
End If
CreateFolder = Replace(CreateFolder,"","/")
If Left(CreateFolder,1)="/" Then
'CreateFolder = Right(CreateFolder,Len(CreateFolder)-1)
End If
If Right(CreateFolder,1)="/" Then
CreateFolder = Left(CreateFolder,Len(CreateFolder)-1)
End If
CreateFolderArray = Split(CreateFolder,"/")
For i = 0 to UBound(CreateFolderArray)
CreateFolderSub = ""
For ii = 0 to i
CreateFolderSub = CreateFolderSub & CreateFolderArray(ii) & "/"
Next
PhCreateFolderSub = Server.MapPath(CreateFolderSub)
If Not objFSO.FolderExists(PhCreateFolderSub) Then
objFSO.CreateFolder(PhCreateFolderSub)
End If
Next
If Err Then
Err.Clear()
Else
BlInfo = True
End If
CreateMultiFolder = BlInfo
End Function


  Rem #################################################################
  Rem ## 建立目录的程序，如果有多级目录，则一级一级的创建
  Rem #################################################################
   Private function CreateDir(ByVal strLocalPath)
    Dim i, strPath, objFolder, tmpPath, tmptPath
    Dim arrPathList, intLevel
    
    'On Error Resume Next
    strPath     = Replace(strLocalPath, "", "\")
    Set objFolder  = server.CreateObject("Scripting.FileSystemObject")
    arrPathList   = Split(strPath, "\")
    intLevel     = UBound(arrPathList)
    For I = 0 To intLevel
     If I = 0 Then
      tmptPath = arrPathList(0) & "\"
     Else
      tmptPath = tmptPath & arrPathList(I) & "\"
     End If
     tmpPath = Left(tmptPath, Len(tmptPath) - 1)
     If Not objFolder.FolderExists(tmpPath) Then objFolder.CreateFolder tmpPath
    Next
    
    Set objFolder = Nothing
    If Err.Number <> 0 Then
     CreateDir = False
     Err.Clear
    Else
     CreateDir = True
    End If
   End function
  Rem #################################################################
  Rem ## 长整数转换
  Rem #################################################################
   Private function toNum(s, default)
    If IsNumeric(s) and s <> "" then
     toNum = CLng(s)
    Else
     toNum = default
    End If
   End function
  Rem #################################################################
 End Class
 Rem #####################################################################################
%> 