<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Template.asp
'文件用途：模板管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System3") Then
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
		Call TemplateList() '模板列表
	Case 2
		Call TemplateAddForm() '添加新文件表单
	Case 3
		Call TemplateAddDo() '添加新文件执行
	Case 4
		Call TemplateEditForm() '修改文件表单
	Case 5
		Call TemplateEditDo() '修改文件执行
	Case 6
		Call TemplateReLoad() '模板重新导入
	Case 7
		Call TemplateAd() '去除广告
	Case 8
		Call TemplateTempList() '模板缓存列表
	Case 9
		Call TemplateTempDel() '模板缓存删除执行
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：TemplateList()
'作    用：模板列表
'参    数：
'==========================================
Sub TemplateList()
	Session("NowPage")=FkFun.GetNowUrl()
	NowFloder=FKFun.HTMLEncode(Trim(Request.QueryString("NowFloder")))
%>
<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>模板管理</h3>
        <a class="no1" href="javascript:void(0);" onclick="ShowBox('Template.asp?Type=2&NowFloder=<%=NowFloder%>','添加新模板文件','1000px','500px');">新建文件</a>
        <a class="no3" href="javascript:void(0);" onclick="DelIt('需要更新模板缓存么？','Template.asp?Type=6','MainRight','<%=Session("NowPage")%>');">重载模板</a>
        <a class="lb" href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=8');return false">模板缓存列表</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop">文件/文件夹名</th>
            <th align="center" class="ListTdTop">类型</th>
            <th align="center" class="ListTdTop" width="140">操作</th>
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
            <td height="20" colspan="3"  style="padding-left:20px; height:31px;"><a style="width:auto; line-height:21px;" href="javascript:void(0);" title="返回上层" onclick="SetRContent('MainRight','Template.asp?Type=1&NowFloder=<%=DirFloder%>');return false">../</a></td>
        </tr>
<%
	End If
	If NowFloder="" Then
		Temp=Server.MapPath("../Skin/")
	Else
		NowFloder=NowFloder&"/"
		Temp=Server.MapPath("../Skin/"&NowFloder)
	End If
	Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	Set F=Fso.GetFolder(Temp)
	Set ObjFloders=F.Subfolders
	For Each ObjFloder In ObjFloders
%>
        <tr>
            <td height="20" style="padding-left:20px;"><a style="width:auto; line-height:21px;" href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=1&NowFloder=<%=NowFloder&ObjFloder.Name%>');return false;"><%=ObjFloder.Name%></a></td>
            <td align="center">文件夹</td>
            <td align="center"><a style="width:30px; text-align:center; line-height:21px;" href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=1&NowFloder=<%=NowFloder&ObjFloder.Name%>');return false;">进入</a></td>
        </tr>
<%
	Next
	Set ObjFloders=Nothing
	Set ObjFiles=F.Files
	For Each ObjFile In ObjFiles
		If Instr(",html,css,",LCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,".")))))>0 Then
%>
        <tr>
            <td height="20"  style="padding-left:20px;"><a style=" width:auto; line-height:21px;" href="javascript:void(0);" onclick="ShowBox('Template.asp?Type=4&File=<%=ObjFile.Name%>&NowFloder=<%=NowFloder%>');return false"><%=ObjFile.Name%></a></td>
            <td align="center">.<%=UCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,"."))))%></td>
            <td align="center"><a  style="width:30px; text-align:center; line-height:21px;" href="javascript:void(0);" onclick="ShowBox('Template.asp?Type=4&File=<%=ObjFile.Name%>&NowFloder=<%=NowFloder%>','修改模板','1000px','500px');return false">修改</a></td>
        </tr>
<%
		End If
	Next
	Set ObjFiles=Nothing
	Set F=Nothing
	Set Fso=Nothing
%>
        <tr>
            <td height="30" colspan="3">&nbsp;</td>
        </tr>
    </table>
<%
If Not FKFso.IsFile("Images/gaitu.txt") Then
%>
<%
End If
%>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：TemplateAddForm()
'作    用：添加新文件表单
'参    数：
'==========================================
Sub TemplateAddForm()
	NowFloder=FKFun.HTMLEncode(Trim(Request.QueryString("NowFloder")))
%>
<form id="TemplateAdd" name="TemplateAdd" method="post" action="Template.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right" width="70">文件名：</td>
            <td>&nbsp;<span style="font-size:14px;">Skin/<%=NowFloder%>/</span>&nbsp;<input name="Fk_Template_Name" type="text" class="Input" id="Fk_Template_Name" /><span style="font-size:14px;">.html</span></td>
        </tr>
        <tr>
            <td height="30" align="right">文件内容：</td>
            <td style="line-height:normal;">&nbsp;<textarea style="border: 1px solid #ccc; height: 500px; margin:10px 0; font-size:12px; color:#555; padding: 5px; width: 790px; line-height:22px;" name="Fk_Template_Content" class="TextArea" id="Fk_Template_Content"></textarea></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:93%;" class="tcbtm">
        <input type="hidden" name="NowFloder" value="<%=NowFloder%>" />
        <input type="submit" onclick="Sends('TemplateAdd','Template.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：TemplateAddDo
'作    用：添加新文件执行
'参    数：
'==============================
Sub TemplateAddDo()
	NowFloder=FKFun.HTMLEncode(Trim(Request.Form("NowFloder")))
	Fk_Template_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Template_Name")))
	Fk_Template_Content=Trim(Request.Form("Fk_Template_Content"))
	Call FKFun.ShowString(NowFloder,1,50,0,"模板文件夹内方可创建文件！","模板文件夹不能大于50个字符！")
	Call FKFun.ShowString(Fk_Template_Name,1,50,0,"请输入文件名！","文件名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Template_Content,1,50000,0,"请输入模板内容！","模板内容不能大于50000个字符！")
	If FKFso.IsFile("../Skin/"&NowFloder&"/"&Fk_Template_Name&".html") Then
		Response.Write("文件已经存在！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If NowFloder=SiteTemplate Then
		Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Fk_Template_Name&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Rs.Eof Then
			Application.Lock()
			Rs.AddNew()
			Rs("Fk_Template_Name")=Fk_Template_Name
			Rs("Fk_Template_Content")=Fk_Template_Content
			Rs.Update()
			Application.UnLock()
		Else
			Application.Lock()
			Rs("Fk_Template_Content")=Fk_Template_Content
			Rs.Update()
			Application.UnLock()
		End If
		Rs.Close
	End If
	Call FKFso.CreateFile("../Skin/"&NowFloder&"/"&Fk_Template_Name&".html",Fk_Template_Content)
	Response.Write("新文件添加成功！")
End Sub

'==========================================
'函 数 名：TemplateEditForm()
'作    用：修改模板表单
'参    数：
'==========================================
Sub TemplateEditForm()
	NowFloder=FKFun.HTMLEncode(Trim(Request.QueryString("NowFloder")))
	NowFile=FKFun.HTMLEncode(Trim(Request.QueryString("File")))
	Fk_Template_Content=Server.HTMLEncode(FKFso.FsoFileRead("../Skin/"&NowFloder&NowFile))
%>
<form id="TemplateEdit" name="TemplateEdit" method="post" action="Template.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right">文件名：</td>
            <td>&nbsp;<span style="font-size:14px;">Skin/<%=NowFloder%><%=NowFile%></span></td>
        </tr>
        <tr>
            <td height="30" align="right">文件内容：</td>
            <td style="line-height:normal;">&nbsp;<textarea style="border: 1px solid #ccc; height: 500px; margin:10px 0; font-size:12px; color:#555; padding: 5px; width: 790px; line-height:22px;" name="Fk_Template_Content" cols="65" rows="20" class="TextArea" id="Fk_Template_Content"><%=Fk_Template_Content%></textarea></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:93%;" class="tcbtm">
		<input type="hidden" name="NowFloder" value="<%=NowFloder%>" />
		<input type="hidden" name="NowFile" value="<%=NowFile%>" />
        <input style="margin-left:20px;" type="submit" onclick="Sends('TemplateEdit','Template.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：TemplateEditDo
'作    用：执行修改模板
'参    数：
'==============================
Sub TemplateEditDo()
	NowFloder=FKFun.HTMLEncode(Trim(Request.Form("NowFloder")))
	NowFile=FKFun.HTMLEncode(Trim(Request.Form("NowFile")))
	Fk_Template_Content=Trim(Request.Form("Fk_Template_Content"))
	Call FKFun.ShowString(Fk_Template_Content,1,50000,0,"请输入模板内容！","模板内容不能大于50000个字符！")
	If Replace(NowFloder,"/","")=SiteTemplate And Instr(NowFile,".html")>0 Then
		Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Replace(NowFile,".html","")&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Application.Lock()
			Rs("Fk_Template_Content")=Fk_Template_Content
			Rs.Update()
			Application.UnLock()
		End If
		Rs.Close
	End If
	Call FKFso.CreateFile("../Skin/"&NowFloder&NowFile,Fk_Template_Content)
	Response.Write("文件修改成功！")
End Sub

'==============================
'函 数 名：TemplateReLoad
'作    用：模板重新导入
'参    数：
'==============================
Sub TemplateReLoad()
	Dim ObjFile,ObjFiles
	Application.Lock()
	Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	Set F=Fso.GetFolder(Server.MapPath("../Skin/"&SiteTemplate))
	Set ObjFiles=F.Files
	For Each ObjFile In ObjFiles
		If LCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,"."))))="html" Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Replace(LCase(ObjFile.Name),".html","")&"'"
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				Rs("Fk_Template_Name")=Replace(LCase(ObjFile.Name),".html","")
				Rs("Fk_Template_Content")=FKFso.FsoFileRead("../Skin/"&SiteTemplate&"/"&ObjFile.Name)
				Rs.Update()
			Else
				Rs.AddNew()
				Rs("Fk_Template_Name")=Replace(LCase(ObjFile.Name),".html","")
				Rs("Fk_Template_Content")=FKFso.FsoFileRead("../Skin/"&SiteTemplate&"/"&ObjFile.Name)
				Rs.Update()
			End If
			Rs.Close
		End If
	Next
	Set ObjFiles=Nothing
	Set F=Nothing
	Set Fso=Nothing
	Application.UnLock()
	Response.Write("模板缓存更新成功！")
End Sub

'==============================
'函 数 名：TemplateAd
'作    用：去除广告
'参    数：
'==============================
Sub TemplateAd()
	Call FKFso.CreateFile("Images/gaitu.txt","不再显示该图网图片")
	Response.Write("处理成功，本图片不会再显示了！")
End Sub

'==========================================
'函 数 名：TemplateTempList()
'作    用：模板缓存列表
'参    数：
'==========================================
Sub TemplateTempList()
	Session("NowPage")=FkFun.GetNowUrl()
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>模板缓存管理</h3>
        <a class="lb" href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=1');return false">模板文件列表</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th width="140" align="center" class="ListTdTop">编号</th>
            <th align="center" class="ListTdTop">名称</th>
            <th width="140" align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Sqlstr="Select * From [Fk_Template] Order By Fk_Template_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Template_Id")%></td>
            <td>&nbsp;<%=Rs("Fk_Template_Name")%></td>
            <td align="center"><a style="width:auto; line-height:21px;" href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Template_Name")%>”，此操作不可逆！','Template.asp?Type=9&Id=<%=Rs("Fk_Template_Id")%>','MainRight','<%=Session("NowPage")%>');return false">删除</a></td>
        </tr>
<%
			Rs.MoveNext
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="5" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="5">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==============================
'函 数 名：TemplateTempDel
'作    用：模板缓存删除执行
'参    数：
'==============================
Sub TemplateTempDel()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("模板缓存删除成功！")
	Else
		Response.Write("广告不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->