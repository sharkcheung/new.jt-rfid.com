<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：GBook.asp
'文件用途：互动管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_GBook_Title,Fk_GBook_Content,Fk_GBook_Name,Fk_GBook_Contact,Fk_GBook_Ip,Fk_GBook_Time,Fk_GBook_ReContent,Fk_GBook_ReAdmin,Fk_GBook_ReIp,Fk_GBook_ReTime
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GBookList() '互动列表
	Case 2
		Call GBookReForm() '回复互动表单
	Case 3
		Call GBookReDo() '执行回复互动
	Case 4
		Call GBookDelDo() '执行删除互动
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：GBookList()
'作    用：互动列表
'参    数：
'==========================================
Sub GBookList()
	Session("NowPage")=FkFun.GetNowUrl()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		PageErr=1
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','GBook.asp?Type=1&ModuleId=<%=Fk_Module_Id%>');return false">刷新内容</a></li>
    </ul>
</div>
<div id="ListTop">
    <%=Fk_Module_Name%>模块&nbsp;&nbsp;请选择模块：
<select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);">
      <option value="alert('请选择模块');">请选择模块</option>
<%
Call ModuleSelectUrl(Fk_Module_Menu,0,Fk_Module_Id)
%>
</select>
</div>
<div id="ListContent">
    <table width="99%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="left" class="ListTdTop">&nbsp;&nbsp;&nbsp; 标题</td>
            <td align="center" class="ListTdTop">联系人</td>
            <td align="center" class="ListTdTop">联系方式</td>
            <td align="center" class="ListTdTop">来源IP</td>
            <td align="center" class="ListTdTop">时间</td>
            <td align="center" class="ListTdTop">回复时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Module="&Fk_Module_Id&" Order By Fk_GBook_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Dim GBookTemplate
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
            <td height="20" align="center"><%=Rs("Fk_GBook_Id")%></td>
            <td align="left" class="lm2" >&nbsp;&nbsp;<%=Rs("Fk_GBook_Title")%></td>
            <td align="center"><%=Rs("Fk_GBook_Name")%></td>
            <td align="center"><%=Rs("Fk_GBook_Contact")%></td>
            <td align="center"><%=Rs("Fk_GBook_Ip")%></td>
            <td align="center"><%=Rs("Fk_GBook_Time")%></td>
            <td align="center">&nbsp;<%=Rs("Fk_GBook_ReTime")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('GBook.asp?Type=2&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_GBook_Id")%>');">回复</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_GBook_Name")%>”？此操作不可逆！','GBook.asp?Type=4&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_GBook_Id")%>','MainRight','<%=Session("NowPage")%>');">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="8">&nbsp;<%Call FKFun.ShowPageCode("GBook.asp?Type=1&ModuleId="&Fk_Module_Id&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="8" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：GBookReForm()
'作    用：回复互动表单
'参    数：
'==========================================
Sub GBookReForm()
	Id=Clng(Request.QueryString("Id"))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Id=Rs("Fk_Module_Id")
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_GBook_Title=Rs("Fk_GBook_Title")
		Fk_GBook_Content=Rs("Fk_GBook_Content")
		Fk_GBook_Name=Rs("Fk_GBook_Name")
		Fk_GBook_Contact=Rs("Fk_GBook_Contact")
		Fk_GBook_Ip=Rs("Fk_GBook_Ip")
		Fk_GBook_Time=Rs("Fk_GBook_Time")
		Fk_GBook_ReContent=Rs("Fk_GBook_ReContent")
	End If
	Rs.Close
%>
<form id="GBookRe" name="GBookRe" method="post" action="GBook.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>互动回复</span></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">标题：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_Title%></td>
    </tr>
     <tr>
        <td height="28" align="right">互动内容：</td>
        <td  class="lm3">&nbsp;<%=Fk_GBook_Content%></td>
    </tr>
    <tr>
        <td height="28" align="right">联系人：</td>
        <td>&nbsp;<%=Fk_GBook_Name%></td>
    </tr>
    <tr>
        <td height="28" align="right">联系方式：</td>
        <td>&nbsp;<%=Fk_GBook_Contact%></td>
    </tr>
    <tr>
        <td height="28" align="right">来源IP：</td>
        <td>&nbsp;<%=Fk_GBook_Ip%></td>
    </tr>
    <tr>
        <td height="28" align="right">时间：</td>
        <td>&nbsp;<%=Fk_GBook_Time%></td>
    </tr>
   
    <tr>
        <td height="28" align="right">回复内容：</td>
        <td>&nbsp;<textarea name="Fk_GBook_ReContent" cols="70" rows="5" class="TextArea" id="Fk_GBook_ReContent"><%=Fk_GBook_ReContent%></textarea></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('GBookRe','GBook.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="回 复" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：GBookReDo
'作    用：执行回复互动
'参    数：
'==============================
Sub GBookReDo()
	Fk_GBook_ReContent=FKFun.HTMLEncode(Request.Form("Fk_GBook_ReContent"))
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_GBook_ReContent,1,1,1,"请输入回复内容，不少于1个字符！","互动内容不能大于1个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Fk_GBook_Title=Rs("Fk_GBook_Title")
		Rs("Fk_GBook_ReContent")=Fk_GBook_ReContent
		Rs("Fk_GBook_ReAdmin")=Request.Cookies("FkAdminId")
		Rs("Fk_GBook_ReIp")=Request.ServerVariables("REMOTE_ADDR")
		Rs("Fk_GBook_ReTime")=Now()
		Rs.Update()
		Application.UnLock()
		Response.Write(Fk_GBook_Title&"回复成功！")
	Else
		Response.Write("互动不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：GBookDelDo
'作    用：执行删除互动
'参    数：
'==============================
Sub GBookDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		'判断权限
		'If Not FkFun.CheckLimit("Module"&Rs("Fk_GBook_Module")) Then
		'	Response.Write("无权限！")
		'	Call FKDB.DB_Close()
		'	Session.CodePage=936
		'	Response.End()
		'End If
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("互动删除成功！")
	Else
		Response.Write("互动不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->