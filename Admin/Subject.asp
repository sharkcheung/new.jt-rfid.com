<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Subject.asp
'文件用途：专题管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System7") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Subject_Name,Fk_Subject_Template,Fk_Subject_Pic

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call SubjectList() '专题列表
	Case 2
		Call SubjectAddForm() '添加专题表单
	Case 3
		Call SubjectAddDo() '执行添加专题
	Case 4
		Call SubjectEditForm() '修改专题表单
	Case 5
		Call SubjectEditDo() '执行修改专题
	Case 6
		Call SubjectDelDo() '执行删除专题
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：SubjectList()
'作    用：专题列表
'参    数：
'==========================================
Sub SubjectList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Subject.asp?Type=2');">添加新专题</a></li>
    </ul>
</div>
<div id="ListTop">
    专题管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">专题名称</td>
            <td align="center" class="ListTdTop">图</td>
            <td align="center" class="ListTdTop">模板</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Subject] Order By Fk_Subject_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
			If Rs("Fk_Subject_Template")>0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Rs("Fk_Subject_Template")
				Rs2.Open Sqlstr,Conn,1,3
				If Not Rs2.Eof Then
					Fk_Subject_Template=Rs2("Fk_Template_Name")
				Else
					Fk_Subject_Template="未知模板"
				End If
				Rs2.Close
			Else
				Fk_Subject_Template="默认模板"
			End If
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Subject_Id")%></td>
            <td align="center"><%=Rs("Fk_Subject_Name")%></td>
            <td align="center"><%If Rs("Fk_Subject_Pic")<>"" Then%><a href="<%=Rs("Fk_Subject_Pic")%>" target="_blank">点击查看</a><%Else%>无图<%End If%></td>
            <td align="center"><%=Fk_Subject_Template%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Subject.asp?Type=4&Id=<%=Rs("Fk_Subject_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Subject_Name")%>”，此操作不可逆！','Subject.asp?Type=6&Id=<%=Rs("Fk_Subject_Id")%>','MainRight','Subject.asp?Type=1');">删除</a></td>
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
	Set Rs2=Nothing
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
'函 数 名：SubjectAddForm()
'作    用：添加专题表单
'参    数：
'==========================================
Sub SubjectAddForm()
%>
<form id="SubjectAdd" name="SubjectAdd" method="post" action="Subject.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">添加新专题[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right">专题名称：</td>
            <td>&nbsp;<input name="Fk_Subject_Name" type="text" class="Input" id="Fk_Subject_Name" /></td>
        </tr>
        <tr>
            <td height="28" align="right">专题图片：</td>
            <td>&nbsp;<input name="Fk_Subject_Pic" type="text" class="Input" id="Fk_Subject_Pic" size="60" /><br />
            &nbsp;<iframe frameborder="0" width="290" height="25" scrolling="No" id="Fk_Subject_Pics" name="Fk_Subject_Pics" src="PicUpLoad.asp?Form=SubjectAdd&Input=Fk_Subject_Pic"></iframe></td>
        </tr>
        <tr>
            <td height="30" align="right">专题模板：</td>
            <td>&nbsp;<select name="Fk_Subject_Template" class="Input" id="Fk_Subject_Template">
                <option value="0">默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:580px;">
        <input type="submit" onclick="Sends('SubjectAdd','Subject.asp?Type=3',0,'',0,1,'MainRight','Subject.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：SubjectAddDo
'作    用：执行添加专题
'参    数：
'==============================
Sub SubjectAddDo()
	Fk_Subject_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Subject_Name")))
	Fk_Subject_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Subject_Pic")))
	Fk_Subject_Template=Trim(Request.Form("Fk_Subject_Template"))
	Call FKFun.ShowString(Fk_Subject_Name,1,50,0,"请输入专题名称！","专题名称不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Subject_Template,"请选择模板！")
	Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Name='"&Fk_Subject_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Subject_Name")=Fk_Subject_Name
		Rs("Fk_Subject_Pic")=Fk_Subject_Pic
		Rs("Fk_Subject_Template")=Fk_Subject_Template
		Rs.Update()
		Application.UnLock()
		Response.Write("新专题添加成功！")
	Else
		Response.Write("该专题名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：SubjectEditForm()
'作    用：修改专题表单
'参    数：
'==========================================
Sub SubjectEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Subject_Name=Rs("Fk_Subject_Name")
		Fk_Subject_Template=Rs("Fk_Subject_Template")
		Fk_Subject_Pic=Rs("Fk_Subject_Pic")
	End If
	Rs.Close
%>
<form id="SubjectEdit" name="SubjectEdit" method="post" action="Subject.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">修改专题[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right">专题名称：</td>
            <td>&nbsp;<input name="Fk_Subject_Name" value="<%=Fk_Subject_Name%>" type="text" class="Input" id="Fk_Subject_Name" /></td>
        </tr>
        <tr>
            <td height="28" align="right">专题图片：</td>
            <td>&nbsp;<input name="Fk_Subject_Pic" value="<%=Fk_Subject_Pic%>" type="text" class="Input" id="Fk_Subject_Pic" size="60" /><br />
            &nbsp;<iframe frameborder="0" width="290" height="25" scrolling="No" id="Fk_Subject_Pics" name="Fk_Subject_Pics" src="PicUpLoad.asp?Form=SubjectEdit&Input=Fk_Subject_Pic"></iframe></td>
        </tr>
        <tr>
            <td height="30" align="right">专题模板：</td>
            <td>&nbsp;<select name="Fk_Subject_Template" class="Input" id="Fk_Subject_Template">
                <option value="0">默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Subject_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:580px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('SubjectEdit','Subject.asp?Type=5',0,'',0,1,'MainRight','Subject.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：SubjectEditDo
'作    用：执行修改专题
'参    数：
'==============================
Sub SubjectEditDo()
	Fk_Subject_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Subject_Name")))
	Fk_Subject_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Subject_Pic")))
	Fk_Subject_Template=Trim(Request.Form("Fk_Subject_Template"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Subject_Name,1,50,0,"请输入专题名称！","专题名称不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Subject_Template,"请选择模板！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Subject_Name")=Fk_Subject_Name
		Rs("Fk_Subject_Pic")=Fk_Subject_Pic
		Rs("Fk_Subject_Template")=Fk_Subject_Template
		Rs.Update()
		Application.UnLock()
		Response.Write("专题修改成功！")
	Else
		Response.Write("专题不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：SubjectDelDo
'作    用：执行删除专题
'参    数：
'==============================
Sub SubjectDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("专题删除成功！")
	Else
		Response.Write("专题不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->