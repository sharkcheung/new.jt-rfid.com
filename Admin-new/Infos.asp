<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Infos.asp
'文件用途：独立信息管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
'If Not FkFun.CheckLimit("System13") Then
'	Response.Write("无权限！")
'	Call FKDB.DB_Close()
'	Session.CodePage=936
'	Response.End()
'End If

'定义页面变量
Dim Fk_Info_Name,Fk_Info_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call InfoList() '独立信息列表
	Case 2
		Call InfoInfodForm() '添加独立信息表单
	Case 3
		Call InfoInfodDo() '执行添加独立信息
	Case 4
		Call InfoEditForm() '修改独立信息表单
	Case 5
		Call InfoEditDo() '执行修改独立信息
	Case 6
		Call InfoDelDo() '执行删除独立信息
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：InfoList()
'作    用：独立信息列表
'参    数：
'==========================================
Sub InfoList()
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>模块设置</h3>
        <a class="gg" href="javascript:void(0);" onclick="ShowBox('siteset.asp?Type=1&Snr=0','广告橱窗','1000px','500px')">广告橱窗</a>
         <%If Request.Cookies("FkAdminLimitId")=0 Then%><a style="width:73px" href="javascript:void(0);" onclick="ShowBox('Infos.asp?Type=2','添加新模块');">添加模块</a><%end if%>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop" width="120">编号</th>
            <th align="center" class="ListTdTop">模块标题</th>
            <th align="center" class="ListTdTop">模块内容</th>
            <th align="center" class="ListTdTop">修改</th>
        </tr>
<%
	Sqlstr="Select * From [Fk_Info] Order By Fk_Info_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Info_Id")%></td>
            <td align="center">［<%=Rs("Fk_Info_Name")%>］&nbsp;&nbsp;&nbsp;&nbsp;标签：{$InfoTit(<%=Rs("Fk_Info_Id")%>)$}</td>
            <td align="center">标签：{$Info(<%=Rs("Fk_Info_Id")%>)$}</td>
            <td align="center" class="no6"><a class="no2" href="javascript:void(0);" title="修改 " onclick="ShowBox('Infos.asp?Type=4&Id=<%=Rs("Fk_Info_Id")%>','修改模块');"></a> <%If Request.Cookies("FkAdminLimitId")=0 Then%><a class="no4" title="删除 " href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Info_Name")%>”，此操作不可逆！','Infos.asp?Type=6&Id=<%=Rs("Fk_Info_Id")%>','MainRight','Infos.asp?Type=1');"></a><%end if%></td>
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

'==========================================
'函 数 名：InfoInfodForm()
'作    用：添加独立信息表单
'参    数：
'==========================================
Sub InfoInfodForm()
%>
<form id="InfoInfod" name="InfoInfod" method="post" action="Infos.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Info_Name" type="text" class="Input" id="Fk_Info_Name" size="40" /></td>
	        </tr>
        <tr>
            <td height="30" align="right" width="100">内容：</td>
            <td style="padding:10px 0 10px 10px;"><textarea name="Fk_Info_Content" class="<%=bianjiqi%>" style="width:100%;" rows="20" class="TextArea" id="Fk_Info_Content"></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:center;" class="tcbtm">
        <input type="submit" onclick="Sends('InfoInfod','Infos.asp?Type=3',0,'',0,1,'MainRight','Infos.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：InfoInfodDo
'作    用：执行添加独立信息
'参    数：
'==============================
Sub InfoInfodDo()
	Fk_Info_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Info_Name")))
	Fk_Info_Content=Request.Form("Fk_Info_Content")
	Call FKFun.ShowString(Fk_Info_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Info_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Sqlstr="Select * From [Fk_Info] Where Fk_Info_Name='"&Fk_Info_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Info_Name")=Fk_Info_Name
		Rs("Fk_Info_Content")=Fk_Info_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新独立信息添加成功！")
	Else
		Response.Write("该独立信息已经被存在，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：InfoEditForm()
'作    用：修改独立信息表单
'参    数：
'==========================================
Sub InfoEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Info_Name=FKFun.HTMLDncode(Rs("Fk_Info_Name"))
		Fk_Info_Content=Rs("Fk_Info_Content")
	End If
	Rs.Close
%>
<form id="InfoEdit" name="InfoEdit" method="post" action="Infos.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Info_Name" type="text" class="Input" id="Fk_Info_Name" value="<%=Fk_Info_Name%>" size="40" /></td>
	        </tr>
        <tr>
            <td height="30" width="100" align="right">内容：</td>
            <td style="padding:10px 0 10px 10px"><textarea name="Fk_Info_Content" class="<%=bianjiqi%>" style="width:100%;" rows="20" class="TextArea" id="Fk_Info_Content"><%=Fk_Info_Content%></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:center;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('InfoEdit','Infos.asp?Type=5',0,'',0,1,'MainRight','Infos.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：InfoEditDo
'作    用：执行修改独立信息
'参    数：
'==============================
Sub InfoEditDo()
	Fk_Info_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Info_Name")))
	Fk_Info_Content=Request.Form("Fk_Info_Content")
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Info_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Info_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Info_Name")=Fk_Info_Name
		Rs("Fk_Info_Content")=Fk_Info_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("独立信息修改成功！")
	Else
		Response.Write("独立信息不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：InfoDelDo
'作    用：执行删除独立信息
'参    数：
'==============================
Sub InfoDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("独立信息删除成功！")
	Else
		Response.Write("独立信息不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->