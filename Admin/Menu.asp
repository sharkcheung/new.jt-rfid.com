<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Menu.asp
'文件用途：菜单管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System6") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Menu_Name

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call MenuList() '菜单列表
	Case 2
		Call MenuAddForm() '添加菜单表单
	Case 3
		Call MenuAddDo() '执行添加菜单
	Case 4
		Call MenuEditForm() '修改菜单表单
	Case 5
		Call MenuEditDo() '执行修改菜单
	Case 6
		Call MenuDelDo() '执行删除菜单
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：MenuList()
'作    用：菜单列表
'参    数：
'==========================================
Sub MenuList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Menu.asp?Type=2');">添加新菜单</a></li>
    </ul>
</div>
<div id="ListTop">
    菜单管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">菜单名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Menu_Id")%></td>
            <td align="center"><%=Rs("Fk_Menu_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Menu.asp?Type=4&Id=<%=Rs("Fk_Menu_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Menu_Name")%>”，此操作不可逆！','Menu.asp?Type=6&Id=<%=Rs("Fk_Menu_Id")%>','Nav|MainLeft|MainRight','Get.asp?Type=1|Get.asp?Type=8|Menu.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="3" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
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
'函 数 名：MenuAddForm()
'作    用：添加菜单表单
'参    数：
'==========================================
Sub MenuAddForm()
%>
<form id="MenuAdd" name="MenuAdd" method="post" action="Menu.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">添加新菜单[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">菜单名称：</td>
	        <td>&nbsp;<input name="Fk_Menu_Name" type="text" class="Input" id="Fk_Menu_Name" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
        <input type="submit" onclick="Sends('MenuAdd','Menu.asp?Type=3',0,'',0,1,'Nav|MainLeft|MainRight','Get.asp?Type=1|Get.asp?Type=8|Menu.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：MenuAddDo
'作    用：执行添加菜单
'参    数：
'==============================
Sub MenuAddDo()
	Fk_Menu_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Menu_Name")))
	Call FKFun.ShowString(Fk_Menu_Name,1,50,0,"请输入菜单名称！","菜单名称不能大于50个字符！")
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Name='"&Fk_Menu_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Menu_Name")=Fk_Menu_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("新菜单添加成功！")
	Else
		Response.Write("该菜单名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：MenuEditForm()
'作    用：修改菜单表单
'参    数：
'==========================================
Sub MenuEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	End If
	Rs.Close
%>
<form id="MenuEdit" name="MenuEdit" method="post" action="Menu.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:400px;">修改菜单[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:400px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">菜单名称：</td>
	        <td>&nbsp;<input name="Fk_Menu_Name" value="<%=Fk_Menu_Name%>" type="text" class="Input" id="Fk_Menu_Name" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:380px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('MenuEdit','Menu.asp?Type=5',0,'',0,1,'Nav|MainLeft|MainRight','Get.asp?Type=1|Get.asp?Type=8|Menu.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：MenuEditDo
'作    用：执行修改菜单
'参    数：
'==============================
Sub MenuEditDo()
	Fk_Menu_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Menu_Name")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Menu_Name,1,50,0,"请输入菜单名称！","菜单名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Menu_Name")=Fk_Menu_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("菜单修改成功！")
	Else
		Response.Write("菜单不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：MenuDelDo
'作    用：执行删除菜单
'参    数：
'==============================
Sub MenuDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write("该菜单尚在使用中，无法删除！")
		Response.End()
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("菜单删除成功！")
	Else
		Response.Write("菜单不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->