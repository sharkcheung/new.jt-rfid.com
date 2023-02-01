<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Word.asp
'文件用途：关键词链接管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System11") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Word_Name,Fk_Word_Url,Fk_Word_level

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WordList() '关键词链接列表
	Case 2
		Call WordAddForm() '添加关键词链接表单
	Case 3
		Call WordAddDo() '执行添加关键词链接
	Case 4
		Call WordEditForm() '修改关键词链接表单
	Case 5
		Call WordEditDo() '执行修改关键词链接
	Case 6
		Call WordDelDo() '执行删除关键词链接
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WordList()
'作    用：关键词链接列表
'参    数：
'==========================================
Sub WordList()
	Session("NowPage")=FkFun.GetNowUrl()
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	on error resume next 
%>
<div id="BoxTop" style="width:98%;"><span>关键词内链</span></div>
<div id="ListNav" style="width:98%; margin-left:8px;border-left:1px solid #7998B7;border-right:1px solid #7998B7;">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Word.asp?Type=2');">添加</a></li>
    </ul>
</div>
<div id="ListTop" style="display:none">
    页面关键词内链管理
</div>
<div id="ListContent" style="width:98%;margin-left:8px;background-color:#FFF;border-left:1px solid #7998B7;border-right:1px solid #7998B7;border-bottom:1px solid #D7E0E7;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">关键词</td>
            <td align="center" class="ListTdTop">优先级别</td>
            <td align="center" class="ListTdTop">链接</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_Word] Order By Fk_Word_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Word_Id")%></td>
            <td align="center"><%=Rs("Fk_Word_Name")%></td>
            <td align="center"><%if isnull(Rs("Fk_Word_level")) then response.Write 0 else response.Write Rs("Fk_Word_level")%></td>
            <td align="center"><%=Rs("Fk_Word_Url")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Word.asp?Type=4&Id=<%=Rs("Fk_Word_Id")%>');" title="修改 "><img src="images/edit.png" /></a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Word_Name")%>”，此操作不可逆！','Word.asp?Type=6&Id=<%=Rs("Fk_Word_Id")%>','MainRight','<%=Session("NowPage")%>');" title="删除 "><img src="images/del.png" /></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" align="center" colspan="5">&nbsp;<%Call FKFun.ShowPageCode("Word.asp?Type=1&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="5" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
	if err then 
		if err.description="在对应所需名称或序数的集合中，未找到项目。"  then
			conn.execute("alter table [Fk_Word] add column [Fk_Word_level] int")
			response.write  "<script language=javascript>window.location.href='/admin/file-shangwin.asp?viewstyle=0&filename=word';</script>"
		end if
	end if
%>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：WordAddForm()
'作    用：添加关键词链接表单
'参    数：
'==========================================
Sub WordAddForm()
%>
<form id="WordAdd" name="WordAdd" method="post" action="Word.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:500px;"><span>添加关键词内链</span></div>
<div id="BoxContents" style="width:500px;">
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">关键词：</td>
	        <td>&nbsp;<input name="Fk_Word_Name" type="text" class="Input" id="Fk_Word_Name" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">优先级别：</td>
	        <td>&nbsp;<input name="Fk_Word_level" type="text" class="Input" id="Fk_Word_level" value="0"  onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"/> (数字越大，越优先显示，只允许数字)</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接：</td>
	        <td>&nbsp;<input name="Fk_Word_Url" type="text" class="Input" id="Fk_Word_Url" value="<%=SiteUrl%>" size="60" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
        <input type="submit" onclick="Sends('WordAdd','Word.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WordAddDo
'作    用：执行添加关键词链接
'参    数：
'==============================
Sub WordAddDo()
	on error resume next
	Fk_Word_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Name")))
	Fk_Word_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Url")))
	Fk_Word_level=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_level")))
	Call FKFun.ShowString(Fk_Word_Name,1,50,0,"请输入关键词！","关键词不能大于50个字符！")
	Call FKFun.ShowString(Fk_Word_Url,1,255,0,"请输入链接！","链接不能大于255个字符！")
	Call FKFun.ShowString(Fk_Word_level,1,5,0,"请输入优先级别！","优先级别不能多于5个字符！")
	Sqlstr="Select * From [Fk_Word] Where Fk_Word_Name='"&Fk_Word_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Word_Name")=Fk_Word_Name
		Rs("Fk_Word_Url")=Fk_Word_Url
		Rs("Fk_Word_level")=Fk_Word_level
		Rs.Update()
		Application.UnLock()
		Response.Write("新关键词链接添加成功！")
	Else
		Response.Write("该类型名称已经被占用，请重新选择！")
	End If
	Rs.Close
	if err then
		'response.Write err.description
		'response.end
		if err.description="在对应所需名称或序数的集合中，未找到项目。"  then
			conn.execute("alter table [Fk_Word] add column [Fk_Word_level] int")
			conn.execute("update [Fk_Word] set [Fk_Word_level] ="&Fk_Word_level)
			response.write  "<script language=javascript>window.location.href='/admin/file-shangwin.asp?viewstyle=0&filename=word';</script>"
		end if
	end if
End Sub

'==========================================
'函 数 名：WordEditForm()
'作    用：修改关键词链接表单
'参    数：
'==========================================
Sub WordEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select Fk_Word_Name,Fk_Word_Url,Fk_Word_level From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Word_Name=FKFun.HTMLDncode(Rs("Fk_Word_Name"))
		Fk_Word_Url=FKFun.HTMLDncode(Rs("Fk_Word_Url"))
		if isnull(Rs("Fk_Word_level")) then 
			Fk_Word_level= 0
		else 
			Fk_Word_level= Rs("Fk_Word_level")
		end if
	End If
	Rs.Close
%>
<form id="WordEdit" name="WordEdit" method="post" action="Word.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:500px;"><span>修改关键词内链</span></div>
<div id="BoxContents" style="width:500px;">
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">关键词：</td>
	        <td>&nbsp;<input name="Fk_Word_Name" value="<%=Fk_Word_Name%>" type="text" class="Input" id="Fk_Word_Name" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">优先级别：</td>
	        <td>&nbsp;<input name="Fk_Word_level" value="<%=Fk_Word_level%>" type="text" class="Input" id="Fk_Word_level"  onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"/> (数字越大，越优先显示，只允许数字)</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接：</td>
	        <td>&nbsp;<input name="Fk_Word_Url" value="<%=Fk_Word_Url%>" type="text" class="Input" id="Fk_Word_Url" size="60" /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('WordEdit','Word.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WordEditDo
'作    用：执行修改关键词链接
'参    数：
'==============================
Sub WordEditDo()
	on error resume next
	Fk_Word_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Name")))
	Fk_Word_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_Url")))
	Fk_Word_level=FKFun.HTMLEncode(Trim(Request.Form("Fk_Word_level")))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Word_Name,1,50,0,"请输入关键词！","关键词不能大于50个字符！")
	Call FKFun.ShowString(Fk_Word_Url,1,255,0,"请输入链接！","链接不能大于50个字符！")
	Call FKFun.ShowString(Fk_Word_level,1,5,0,"请输入优先级别！","优先级别不能多于5个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Word_Name")=Fk_Word_Name
		Rs("Fk_Word_Url")=Fk_Word_Url
		Rs("Fk_Word_level")=Fk_Word_level
		Rs.Update()
		Application.UnLock()
		Response.Write("关键词链接修改成功！")
	Else
		Response.Write("关键词链接不存在！")
	End If
	Rs.Close
	if err then
		if err.description="在对应所需名称或序数的集合中，未找到项目。"  then
			conn.execute("alter table [Fk_Word] add column [Fk_Word_level] int")
			conn.execute("update [Fk_Word] set [Fk_Word_level] ="&Fk_Word_level)
			response.write  "<script language=javascript>window.location.href='/admin/file-shangwin.asp?viewstyle=0&filename=word';</script>"
		end if
	end if
End Sub

'==============================
'函 数 名：WordDelDo
'作    用：执行删除关键词链接
'参    数：
'==============================
Sub WordDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Word] Where Fk_Word_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("关键词链接删除成功！")
	Else
		Response.Write("关键词链接不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->