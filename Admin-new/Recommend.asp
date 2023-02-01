<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Recommend.asp
'文件用途：推荐类型管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System8") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Recommend_Name

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call RecommendList() '推荐类型列表
	Case 2
		Call RecommendAddForm() '添加推荐类型表单
	Case 3
		Call RecommendAddDo() '执行添加推荐类型
	Case 4
		Call RecommendEditForm() '修改推荐类型表单
	Case 5
		Call RecommendEditDo() '执行修改推荐类型
	Case 6
		Call RecommendDelDo() '执行删除推荐类型
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：RecommendList()
'作    用：推荐类型列表
'参    数：
'==========================================
Sub RecommendList()
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>推荐类型管理</h3>
        <a class="tjxtjlx" href="javascript:void(0);" onclick="ShowBox('Recommend.asp?Type=2','添加新推荐类型','459px');">添加新推荐类型</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop">编号</th>
            <th align="center" class="ListTdTop">类型名称</th>
            <th width="120" align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Sqlstr="Select * From [Fk_Recommend] Order By Fk_Recommend_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Recommend_Id")%></td>
            <td align="center"><%=Rs("Fk_Recommend_Name")%></td>
            <td align="center" class="no6"><a class="no2" href="javascript:void(0);" onclick="ShowBox('Recommend.asp?Type=4&Id=<%=Rs("Fk_Recommend_Id")%>','修改推荐类型','450px');"></a> <a class="no4" href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Recommend_Name")%>”，此操作不可逆！','Recommend.asp?Type=6&Id=<%=Rs("Fk_Recommend_Id")%>','MainRight','Recommend.asp?Type=1');"></a></td>
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
'函 数 名：RecommendAddForm()
'作    用：添加推荐类型表单
'参    数：
'==========================================
Sub RecommendAddForm()
%>
<form id="RecommendAdd" name="RecommendAdd" method="post" action="Recommend.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">类型名称：</td>
	        <td><input name="Fk_Recommend_Name" type="text" class="Input" id="Fk_Recommend_Name" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
        <input style="margin-left:113px" type="submit" onclick="Sends('RecommendAdd','Recommend.asp?Type=3',0,'',0,1,'MainRight','Recommend.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：RecommendAddDo
'作    用：执行添加推荐类型
'参    数：
'==============================
Sub RecommendAddDo()
	Fk_Recommend_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Recommend_Name")))
	Call FKFun.ShowString(Fk_Recommend_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Name='"&Fk_Recommend_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Recommend_Name")=Fk_Recommend_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("新推荐类型添加成功！")
	Else
		Response.Write("该类型名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：RecommendEditForm()
'作    用：修改推荐类型表单
'参    数：
'==========================================
Sub RecommendEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Recommend_Name=Rs("Fk_Recommend_Name")
	End If
	Rs.Close
%>
<form id="RecommendEdit" name="RecommendEdit" method="post" action="Recommend.asp?Type=5" onsubmit="return false;">

<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">类型名称：</td>
	        <td><input name="Fk_Recommend_Name" value="<%=Fk_Recommend_Name%>" type="text" class="Input" id="Fk_Recommend_Name" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input style="margin-left:113px;" type="submit" onclick="Sends('RecommendEdit','Recommend.asp?Type=5',0,'',0,1,'MainRight','Recommend.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：RecommendEditDo
'作    用：执行修改推荐类型
'参    数：
'==============================
Sub RecommendEditDo()
	Fk_Recommend_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Recommend_Name")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Recommend_Name,1,50,0,"请输入类型名称！","类型名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Recommend_Name")=Fk_Recommend_Name
		Rs.Update()
		Application.UnLock()
		Response.Write("推荐类型修改成功！")
	Else
		Response.Write("推荐类型不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：RecommendDelDo
'作    用：执行删除推荐类型
'参    数：
'==============================
Sub RecommendDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Recommend Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write("该推荐类型尚在使用中，无法删除！")
		Response.End()
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Recommend Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write("该推荐类型尚在使用中，无法删除！")
		Response.End()
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("推荐类型删除成功！")
	Else
		Response.Write("推荐类型不存在！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->