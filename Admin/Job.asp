<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Job.asp
'文件用途：招聘管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System4") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Job_Name,Fk_Job_Count,Fk_Job_About,Fk_Job_Area,Fk_Job_Date

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call JobList() '招聘列表
	Case 2
		Call JobAddForm() '添加招聘表单
	Case 3
		Call JobAddDo() '执行添加招聘
	Case 4
		Call JobEditForm() '修改招聘表单
	Case 5
		Call JobEditDo() '执行修改招聘
	Case 6
		Call JobDelDo() '执行删除招聘
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：JobList()
'作    用：招聘列表
'参    数：
'==========================================
Sub JobList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Job.asp?Type=2');">添加新招聘</a></li>
    </ul>
</div>
<div id="ListTop">
    招聘管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">职位</td>
            <td align="center" class="ListTdTop">人数</td>
            <td align="center" class="ListTdTop">工作地点</td>
            <td align="center" class="ListTdTop">有效期</td>
            <td align="center" class="ListTdTop">添加时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_Job] Order By Fk_Job_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Job_Id")%></td>
            <td align="center"><%=Rs("Fk_Job_Name")%></td>
            <td align="center"><%=Rs("Fk_Job_Count")%></td>
            <td align="center"><%=Rs("Fk_Job_Area")%></td>
            <td align="center"><%If Rs("Fk_Job_Date")=0 Then%>长期<%Else%><%=Rs("Fk_Job_Date")%><%End If%></td>
            <td align="center"><%=Rs("Fk_Job_Time")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Job.asp?Type=4&Id=<%=Rs("Fk_Job_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Job_Name")%>”，此操作不可逆！','Job.asp?Type=6&Id=<%=Rs("Fk_Job_Id")%>','MainRight','Job.asp?Type=1');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="7" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="7">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：JobAddForm()
'作    用：添加招聘表单
'参    数：
'==========================================
Sub JobAddForm()
%>
<form id="JobAdd" name="JobAdd" method="post" action="Job.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新招聘[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">职位：</td>
	        <td>&nbsp;<input name="Fk_Job_Name" type="text" class="Input" id="Fk_Job_Name" /></td>
	        </tr>
        <tr>
            <td height="30" align="right">招聘人数：</td>
            <td>&nbsp;<input name="Fk_Job_Count" type="text" class="Input" id="Fk_Job_Count" /></td>
        </tr>
        <tr>
            <td height="30" align="right">工作地点：</td>
            <td>&nbsp;<input name="Fk_Job_Area" type="text" class="Input" id="Fk_Job_Area" /></td>
        </tr>
        <tr>
            <td height="30" align="right">有效期：</td>
            <td>&nbsp;<input name="Fk_Job_Date" type="text" class="Input" id="Fk_Job_Date" />&nbsp;天（请输入数字，如果长期有效请填0）</td>
        </tr>
        <tr>
            <td height="30" align="right">招聘要求：</td>
            <td>&nbsp;<textarea name="Fk_Job_About" cols="60" rows="10" class="TextArea" id="Fk_Job_About"></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('JobAdd','Job.asp?Type=3',0,'',0,1,'MainRight','Job.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：JobAddDo
'作    用：执行添加招聘
'参    数：
'==============================
Sub JobAddDo()
	Fk_Job_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Name")))
	Fk_Job_Count=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Count")))
	Fk_Job_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_About")))
	Fk_Job_Area=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Area")))
	Fk_Job_Date=Trim(Request.Form("Fk_Job_Date"))
	Call FKFun.ShowString(Fk_Job_Name,1,255,0,"请输入招聘职位！","招聘职位不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_Count,1,50,0,"请输入招聘数量！","招聘数量不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_About,1,5000,0,"请输入招聘要求！","招聘要求不能大于5000个字符！")
	Call FKFun.ShowString(Fk_Job_Area,1,50,0,"请输入工作地点！","工作地点不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Job_Date,"请输入有效期，单位为天，必须是数字，长期有效请填0！")
	Sqlstr="Select * From [Fk_Job] Where Fk_Job_Name='"&Fk_Job_Name&"' And Fk_Job_About='"&Fk_Job_Area&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Job_Name")=Fk_Job_Name
		Rs("Fk_Job_Count")=Fk_Job_Count
		Rs("Fk_Job_About")=Fk_Job_About
		Rs("Fk_Job_Area")=Fk_Job_Area
		Rs("Fk_Job_Date")=Fk_Job_Date
		Rs.Update()
		Application.UnLock()
		Response.Write("新招聘添加成功！")
	Else
		Response.Write("该招聘职位已经被发布，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：JobEditForm()
'作    用：修改招聘表单
'参    数：
'==========================================
Sub JobEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Job_Name=FKFun.HTMLDncode(Rs("Fk_Job_Name"))
		Fk_Job_Count=FKFun.HTMLDncode(Rs("Fk_Job_Count"))
		Fk_Job_About=FKFun.HTMLDncode(Rs("Fk_Job_About"))
		Fk_Job_Area=FKFun.HTMLDncode(Rs("Fk_Job_Area"))
		Fk_Job_Date=Rs("Fk_Job_Date")
	End If
	Rs.Close
%>
<form id="JobEdit" name="JobEdit" method="post" action="Job.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改招聘[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">职位：</td>
	        <td>&nbsp;<input name="Fk_Job_Name" value="<%=Fk_Job_Name%>" type="text" class="Input" id="Fk_Job_Name" /></td>
	        </tr>
        <tr>
            <td height="30" align="right">招聘人数：</td>
            <td>&nbsp;<input name="Fk_Job_Count" value="<%=Fk_Job_Count%>" type="text" class="Input" id="Fk_Job_Count" /></td>
        </tr>
        <tr>
            <td height="30" align="right">工作地点：</td>
            <td>&nbsp;<input name="Fk_Job_Area" value="<%=Fk_Job_Area%>" type="text" class="Input" id="Fk_Job_Area" /></td>
        </tr>
        <tr>
            <td height="30" align="right">有效期：</td>
            <td>&nbsp;<input name="Fk_Job_Date" value="<%=Fk_Job_Date%>" type="text" class="Input" id="Fk_Job_Date" />&nbsp;天（请输入数字，如果长期有效请填0）</td>
        </tr>
        <tr>
            <td height="30" align="right">招聘要求：</td>
            <td>&nbsp;<textarea name="Fk_Job_About" cols="60" rows="10" class="TextArea" id="Fk_Job_About"><%=Fk_Job_About%></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('JobEdit','Job.asp?Type=5',0,'',0,1,'MainRight','Job.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：JobEditDo
'作    用：执行修改招聘
'参    数：
'==============================
Sub JobEditDo()
	Fk_Job_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Name")))
	Fk_Job_Count=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Count")))
	Fk_Job_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_About")))
	Fk_Job_Area=FKFun.HTMLEncode(Trim(Request.Form("Fk_Job_Area")))
	Fk_Job_Date=Trim(Request.Form("Fk_Job_Date"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Job_Name,1,255,0,"请输入招聘职位！","招聘职位不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_Count,1,50,0,"请输入招聘数量！","招聘数量不能大于50个字符！")
	Call FKFun.ShowString(Fk_Job_About,1,5000,0,"请输入招聘要求！","招聘要求不能大于5000个字符！")
	Call FKFun.ShowString(Fk_Job_Area,1,50,0,"请输入工作地点！","工作地点不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Job_Date,"请输入有效期，单位为天，必须是数字，长期有效请填0！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Job_Name")=Fk_Job_Name
		Rs("Fk_Job_Count")=Fk_Job_Count
		Rs("Fk_Job_About")=Fk_Job_About
		Rs("Fk_Job_Area")=Fk_Job_Area
		Rs("Fk_Job_Date")=Fk_Job_Date
		Rs.Update()
		Application.UnLock()
		Response.Write("招聘修改成功！")
	Else
		Response.Write("招聘不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：JobDelDo
'作    用：执行删除招聘
'参    数：
'==============================
Sub JobDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Job] Where Fk_Job_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("招聘删除成功！")
	Else
		Response.Write("招聘不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->