<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Field.asp
'文件用途：自定义字段管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Field_Name,Fk_Field_Tag,Fk_Field_Type,Fk_Field_Type1,Fk_Field_Type2

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FieldList() '自定义字段列表
	Case 2
		Call FieldFielddForm() '添加自定义字段表单
	Case 3
		Call FieldFielddDo() '执行添加自定义字段
	Case 4
		Call FieldEditForm() '修改自定义字段表单
	Case 5
		Call FieldEditDo() '执行修改自定义字段
	Case 6
		Call FieldDelDo() '执行删除自定义字段
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FieldList()
'作    用：自定义字段列表
'参    数：
'==========================================
Sub FieldList()
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>自定义字段管理</h3>
        <a class="tjxtjlx" href="javascript:void(0);" onclick="ShowBox('Field.asp?Type=2','添加自定义字段','450px');">添加自定义字段</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th width="70" align="center" class="ListTdTop">编号</th>
            <th width="200" align="center" class="ListTdTop">名称</th>
            <th align="center" class="ListTdTop">类型</th>
            <th align="center" class="ListTdTop">列表标签</th>
            <th align="center" class="ListTdTop">内容标签</th>
            <th width="120" align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Sqlstr="Select * From [Fk_Field] Order By Fk_Field_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
			Select Case Rs("Fk_Field_Type")
				Case 0
					Fk_Field_Type="文章模块"
					Fk_Field_Type1="{$ArticleList_"&Rs("Fk_Field_Tag")&"$}"
					Fk_Field_Type2="{$Article_"&Rs("Fk_Field_Tag")&"$}"
				Case 1
					Fk_Field_Type="产品模块"
					Fk_Field_Type1="{$ProductList_"&Rs("Fk_Field_Tag")&"$}"
					Fk_Field_Type2="{$Product_"&Rs("Fk_Field_Tag")&"$}"
				Case 2
					Fk_Field_Type="下载模块"
					Fk_Field_Type1="{$DownList_"&Rs("Fk_Field_Tag")&"$}"
					Fk_Field_Type2="{$Down_"&Rs("Fk_Field_Tag")&"$}"
			End Select
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Field_Id")%></td>
            <td>&nbsp;<%=Rs("Fk_Field_Name")%></td>
            <td align="center"><%=Fk_Field_Type%></td>
            <td align="center"><%=Fk_Field_Type1%></td>
            <td align="center"><%=Fk_Field_Type2%></td>
            <td align="center" class="no6"><a class="no2" href="javascript:void(0);" onclick="ShowBox('Field.asp?Type=4&Id=<%=Rs("Fk_Field_Id")%>','修改自定义字段','450px');"></a> <a class="no4" href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Field_Name")%>”，此操作不可逆！','Field.asp?Type=6&Id=<%=Rs("Fk_Field_Id")%>','MainRight','Field.asp?Type=1');"></a></td>
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
'函 数 名：FieldFielddForm()
'作    用：添加自定义字段表单
'参    数：
'==========================================
Sub FieldFielddForm()
%>
<form id="FieldFieldd" name="FieldFieldd" method="post" action="Field.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">名称：</td>
	        <td><input name="Fk_Field_Name" type="text" class="Input" id="Fk_Field_Name" size="40" /></td>
	        </tr>
        <tr>
            <td height="30" align="right">标签：</td>
            <td style="padding-left:10px;"><input style="margin-left:0; margin-top:5px" name="Fk_Field_Tag" type="text" class="Input" id="Fk_Field_Tag" size="20" /><br />*输入单词即可，添加后不可修改</td>
        </tr>
        <tr>
            <td height="30" align="right">类型：</td>
            <td><select name="Fk_Field_Type" class="Input" id="Fk_Field_Type">
                    <option value="0">文章模块</option>
                    <option value="1">产品模块</option>
                    <option value="2">下载模块</option>
                    </select></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('FieldFieldd','Field.asp?Type=3',0,'',0,1,'MainRight','Field.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FieldFielddDo
'作    用：执行添加自定义字段
'参    数：
'==============================
Sub FieldFielddDo()
	Fk_Field_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Name")))
	Fk_Field_Tag=FKFun.HTMLEncode(Request.Form("Fk_Field_Tag"))
	Fk_Field_Type=Trim(Request.Form("Fk_Field_Type"))
	Call FKFun.ShowString(Fk_Field_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Field_Tag,1,50,0,"请输入标签！","标签不能大于5000个字符！")
	Call FKFun.ShowNum(Fk_Field_Type,"请选择自定义字段类型！")
	Sqlstr="Select * From [Fk_Field] Where (Fk_Field_Name='"&Fk_Field_Name&"' Or Fk_Field_Tag='"&Fk_Field_Tag&"') And Fk_Field_Type="&Fk_Field_Type&""
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Field_Name")=Fk_Field_Name
		Rs("Fk_Field_Tag")=Fk_Field_Tag
		Rs("Fk_Field_Type")=Fk_Field_Type
		Rs.Update()
		Application.UnLock()
		Response.Write("新自定义字段添加成功！")
	Else
		Response.Write("该自定义字段已经被存在，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：FieldEditForm()
'作    用：修改自定义字段表单
'参    数：
'==========================================
Sub FieldEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Field_Name=FKFun.HTMLDncode(Rs("Fk_Field_Name"))
		Fk_Field_Tag=FKFun.HTMLDncode(Rs("Fk_Field_Tag"))
		Fk_Field_Type=Rs("Fk_Field_Type")
	End If
	Rs.Close
%>
<form id="FieldEdit" name="FieldEdit" method="post" action="Field.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Field_Name" value="<%=Fk_Field_Name%>" type="text" class="Input" id="Fk_Field_Name" size="40" /></td>
	        </tr>
        <tr>
            <td height="30" align="right">标签：</td>
            <td>&nbsp;<input name="Fk_Field_Tag" value="<%=Fk_Field_Tag%>" type="text" class="Input" id="Fk_Field_Tag" size="20" disabled="disabled" /></td>
        </tr>
        <tr>
            <td height="30" align="right">类型：</td>
            <td>&nbsp;<select name="Fk_Field_Type" class="Input" id="Fk_Field_Type">
                    <option value="0"<%=FKFun.BeSelect(Fk_Field_Type,0)%>>文章模块</option>
                    <option value="1"<%=FKFun.BeSelect(Fk_Field_Type,1)%>>产品模块</option>
                    <option value="2"<%=FKFun.BeSelect(Fk_Field_Type,2)%>>下载模块</option>
                    </select></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; text-align:left; margin:0 auto;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input style="margin-left:113px;" type="submit" onclick="Sends('FieldEdit','Field.asp?Type=5',0,'',0,1,'MainRight','Field.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FieldEditDo
'作    用：执行修改自定义字段
'参    数：
'==============================
Sub FieldEditDo()
	Fk_Field_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Field_Name")))
	Fk_Field_Type=Trim(Request.Form("Fk_Field_Type"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Field_Name,1,50,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Field_Type,"请选择自定义字段类型！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Field_Name")=Fk_Field_Name
		Rs("Fk_Field_Type")=Fk_Field_Type
		Rs.Update()
		Application.UnLock()
		Response.Write("自定义字段修改成功！")
	Else
		Response.Write("自定义字段不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：FieldDelDo
'作    用：执行删除自定义字段
'参    数：
'==============================
Sub FieldDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("自定义字段删除成功！")
	Else
		Response.Write("自定义字段不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->