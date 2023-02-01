<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Fiel.asp
'文件用途：文件管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

'定义页面变量
Dim NowFile,NowFloder,DirFloder,ObjFiles,ObjFile,ObjFloders,ObjFloder
Dim Fk_Template_Name,Fk_Template_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FileList() '上传文件列表
	Case 2
		Call FileDelDo() '删除上传文件执行
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FileList()
'作    用：上传文件列表
'参    数：
'==========================================
Sub FileList()
	Session("NowPage")=FkFun.GetNowUrl()
	NowFloder=FKFun.HTMLEncode(Trim(Request.QueryString("NowFloder")))
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    上传文件管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">文件/文件夹名</td>
            <td align="center" class="ListTdTop">类型</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	TempArr=Split(NowFloder,"/")
	For i=0 To UBound(TempArr)-1
		If DirFloder="" Then
			DirFloder=TempArr(i)
		Else
			DirFloder=DirFloder&"/"&TempArr(i)
		End If
	Next
	If NowFloder<>"" Then
%>
        <tr>
            <td height="20" colspan="3">&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:void(0);" title="返回上层" onclick="SetRContent('MainRight','File.asp?Type=1&NowFloder=<%=DirFloder%>')">../</a></td>
        </tr>
<%
	End If
	If NowFloder="" Then
		Temp=Server.MapPath("../Up/")
	Else
		NowFloder=NowFloder&"/"
		Temp=Server.MapPath("../Up/"&NowFloder)
	End If
	Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	Set F=Fso.GetFolder(Temp)
	Set ObjFloders=F.Subfolders
	For Each ObjFloder In ObjFloders
%>
        <tr>
            <td height="20">&nbsp;&nbsp;<a href="javascript:void(0);" onclick="SetRContent('MainRight','File.asp?Type=1&NowFloder=<%=NowFloder&ObjFloder.Name%>')"><%=ObjFloder.Name%></a></td>
            <td align="center">文件夹</td>
            <td align="center"><a href="javascript:void(0);" onclick="SetRContent('MainRight','File.asp?Type=1&NowFloder=<%=NowFloder&ObjFloder.Name%>')">进入</a></td>
        </tr>
<%
	Next
	Set ObjFloders=Nothing
	Set ObjFiles=F.Files
	For Each ObjFile In ObjFiles
%>
        <tr>
            <td height="20">&nbsp;&nbsp;<a href="../Up/<%=NowFloder&ObjFile.Name%>" target="_blank"><%=ObjFile.Name%></a></td>
            <td align="center">.<%=UCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,"."))))%></td>
            <td align="center"><a href="javascript:void(0);" onclick="DelIt('是否删除“<%=ObjFile.Name%>”？','File.asp?Type=2&File=<%=Server.URLEncode(NowFloder&ObjFile.Name)%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
	Next
	Set ObjFiles=Nothing
	Set F=Nothing
	Set Fso=Nothing
%>
        <tr>
            <td height="30" colspan="3">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FileDelDo()
'作    用：删除上传文件执行
'参    数：
'==========================================
Sub FileDelDo()
	Temp=Request.QueryString("File")
	Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
	Fso.DeleteFile(Server.MapPath("../Up/"&Temp))
	Set Fso = nothing
	Response.Write("文件删除成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->