<!--#Include File="AdminCheck.asp"--><head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
</head>

<%
'==========================================
'文 件 名：Vote.asp
'文件用途：市场调查管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System12") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Vote_Name,Fk_Vote_Content,Fk_Vote_Ticket,Fk_Vote_Count

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call VoteList() '市场调查列表
	Case 2
		Call VoteAddForm() '添加市场调查表单
	Case 3
		Call VoteAddDo() '执行添加市场调查
	Case 4
		Call VoteEditForm() '修改市场调查表单
	Case 5
		Call VoteEditDo() '执行修改市场调查
	Case 6
		Call VoteDelDo() '执行删除市场调查
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：VoteList()
'作    用：市场调查列表
'参    数：
'==========================================
Sub VoteList()
%>


<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>市场调查管理</h3>
        <a class="no1" href="javascript:void(0);" onclick="ShowBox('Vote.asp?Type=2','添加新市场调查');">添加调查</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop">编号</th>
            <th align="center" class="ListTdTop">名称</th>
            <th align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Sqlstr="Select * From [Fk_Vote] Order By Fk_Vote_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Vote_Id")%></td>
            <td>&nbsp;<%=Rs("Fk_Vote_Name")%>&nbsp;&nbsp;<a style="width:auto; background:none; height:auto; line-height:auto;" href="javascript:void(0);" onclick="window.clipboardData.setData('Text','<script type=text/javascript src=<%=SiteDir%>Vote.asp?Id=<%=Rs("Fk_Vote_Id")%>></script>');layer.msg('市场调查代码复制成功');">[复制代码]</a></td>
            <td align="center" class="no6"><a class="no2" href="javascript:void(0);" title="修改 " onclick="ShowBox('Vote.asp?Type=4&Id=<%=Rs("Fk_Vote_Id")%>');"></a> <a class="no4" href="javascript:void(0);" title="删除 " onclick="DelIt('您确认要删除“<%=Rs("Fk_Vote_Name")%>”，此操作不可逆！','Vote.asp?Type=6&Id=<%=Rs("Fk_Vote_Id")%>','MainRight','Vote.asp?Type=1');"></a></td>
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
'函 数 名：VoteAddForm()
'作    用：添加市场调查表单
'参    数：
'==========================================
Sub VoteAddForm()
%>
<form id="VoteVoted" name="VoteVoted" method="post" action="Vote.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Vote_Name" type="text" class="Input" id="Fk_Vote_Name" /></td>
	        </tr>
        <tr>
            <td align="right">条目：</td>
            <td style="line-height:normal">&nbsp;<textarea style="margin:10px 0 10px 10px; border:1px solid #ccc; color:#999; padding:5px;" name="Fk_Vote_Content" cols="60" rows="10" class="TextArea" id="Fk_Vote_Content"></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('VoteVoted','Vote.asp?Type=3',0,'',0,1,'MainRight','Vote.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：VoteAddDo
'作    用：执行添加市场调查
'参    数：
'==============================
Sub VoteAddDo()
	Fk_Vote_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Name")))
	Fk_Vote_Content=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Content")))
	Call FKFun.ShowString(Fk_Vote_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Vote_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Name='"&Fk_Vote_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Vote_Name")=Fk_Vote_Name
		Rs("Fk_Vote_Content")=Fk_Vote_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新市场调查添加成功！")
	Else
		Response.Write("该市场调查已经被发布，请查看后重新添加！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：VoteEditForm()
'作    用：修改市场调查表单
'参    数：
'==========================================
Sub VoteEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Vote_Name=FKFun.HTMLDncode(Rs("Fk_Vote_Name"))
		Fk_Vote_Content=FKFun.HTMLDncode(Rs("Fk_Vote_Content"))
	End If
	Rs.Close
%>
<form id="VoteEdit" name="VoteEdit" method="post" action="Vote.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">名称：</td>
	        <td>&nbsp;<input name="Fk_Vote_Name" value="<%=Fk_Vote_Name%>" type="text" class="Input" id="Fk_Vote_Name" /></td>
	        </tr>
        <tr>
            <td height="30" align="right">条目：</td>
            <td>&nbsp;<textarea style="border:1px solid #ccc; margin:10px; font-size:12px; padding:5px; color:#999;" name="Fk_Vote_Content" cols="60" rows="10" class="TextArea" id="Fk_Vote_Content"><%=Fk_Vote_Content%></textarea></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input style="margin-left:113px;" type="submit" onclick="Sends('VoteEdit','Vote.asp?Type=5',0,'',0,1,'MainRight','Vote.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：VoteEditDo
'作    用：执行修改市场调查
'参    数：
'==============================
Sub VoteEditDo()
	Fk_Vote_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Name")))
	Fk_Vote_Content=FKFun.HTMLEncode(Trim(Request.Form("Fk_Vote_Content")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Vote_Name,1,255,0,"请输入名称！","名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Vote_Content,1,5000,0,"请输入内容！","内容不能大于5000个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Temp=Rs("Fk_Vote_Ticket")
		If UBound(Split(Fk_Vote_Content,"<br />"))>UBound(Split(Temp,"|")) Then
			For i=1 To UBound(Split(Fk_Vote_Content,"<br />"))-UBound(Split(Temp,"|"))
				Temp=Temp&"|0"
			Next
		End If
		Application.Lock()
		Rs("Fk_Vote_Name")=Fk_Vote_Name
		Rs("Fk_Vote_Content")=Fk_Vote_Content
		Rs("Fk_Vote_Ticket")=Temp
		Rs.Update()
		Application.UnLock()
		Response.Write("市场调查修改成功！")
	Else
		Response.Write("市场调查不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：VoteDelDo
'作    用：执行删除市场调查
'参    数：
'==============================
Sub VoteDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Vote] Where Fk_Vote_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("市场调查删除成功！")
	Else
		Response.Write("市场调查不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->