<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：GBook.asp
'文件用途：互动管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_GBook_Name,Fk_GBook_mobile,Fk_GBook_company,Fk_GBook_position,Fk_GBook_address,Fk_GBook_age,Fk_GBook_email,Fk_GBook_cardid,Fk_GBook_area,Fk_GBook_country,Fk_GBook_birth,Fk_GBook_edu,Fk_GBook_joinedbefore,Fk_GBook_why,Fk_GBook_familyjoined,Fk_GBook_wholejoin,Fk_GBook_health,Fk_GBook_ip,Fk_GBook_time,Fk_GBook_ispass,Fk_isWhole,Fk_Bad_Reason,Fk_GBook_sex,Fk_GBook_tj

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
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onClick="SetRContent('MainRight','<%=Session("NowPage")%>')">刷新内容</a></li>
    </ul>
</div>

<div id="ListContent">
    <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
            <td align="center" class="ListTdTop">选择</td>
            <td align="left" class="ListTdTop">&nbsp;&nbsp;&nbsp; 姓名</td>
            <td align="center" class="ListTdTop">来源</td>
            <td align="center" class="ListTdTop">手机</td>
            <td align="center" class="ListTdTop">公司</td>
            <td align="center" class="ListTdTop">邮箱</td>
            <td align="center" class="ListTdTop">职位</td>
            <td align="center" class="ListTdTop">身份证</td>
            <td align="center" class="ListTdTop">申请时间</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_AD_Apply] Order by ID Desc"
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
            <td height="20" align="center"><input name="id" type="checkbox" value="<%=rs("ID")%>"></td>
            <td align="left" class="lm2" >&nbsp;&nbsp;<%=Rs("Fk_Apply_name")%></td>
            <td align="center"><%=Rs("Fk_Apply_from")%></td>
            <td align="center"><%=Rs("Fk_Apply_tel")%></td>
            <td align="center"><%=Rs("Fk_Apply_company")%></td>
            <td align="center"><%=Rs("Fk_Apply_email")%></td>
            <td align="center"><%=Rs("Fk_Apply_postion")%></td>
            <td align="center">&nbsp;<%=Rs("Fk_Apply_idcard")%></td>
            <td align="center">&nbsp;<%=Rs("Fk_Apply_Time")%></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
	<tr > 
<td height="30" colspan="9" align="right">全选 
<input type="checkbox" name="checkbox" value="Check All" onClick="SelectAll('id')">
<input onClick="var str='';$('input[name=id]').each(function(){if(this.checked){if(str==''){str=$(this).val();}else{str+=','+$(this).val()}}});DelIt('您确认要删除，此操作不可逆！','api_M_Online.asp?Type=4&id='+str,'MainRight','<%=Session("NowPage")%>');" type="button" class="Button" name="Submit" value="删 除" >
&nbsp;</td>
</tr>

        <tr>
            <td height="30" colspan="9" style="text-align:center;">&nbsp;<%Call FKFun.ShowPageCode("api_M_Online.asp?Type=1&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="9" align="center">暂无记录</td>
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
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	Sqlstr="Select * From [API_Online_Registration] Where ID=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_GBook_Name=Rs("G_name")
		Fk_GBook_mobile=Rs("G_mobile")
		Fk_GBook_company=Rs("G_company")
		Fk_GBook_position=Rs("G_position")
		Fk_GBook_address=Rs("G_address")
		Fk_GBook_age=Rs("G_age")
		Fk_GBook_email=Rs("G_email")
		
		Fk_GBook_joinedbefore=Rs("G_joinedbefore")
		Fk_GBook_why=Rs("G_why")
		Fk_GBook_cardid=Rs("G_cardid")
		Fk_GBook_area=Rs("G_area")
		Fk_GBook_country=Rs("G_country")
		Fk_GBook_birth=Rs("G_birth")
		Fk_GBook_edu=Rs("G_edu")
		
		Fk_GBook_familyjoined=Rs("G_familyjoined")
		Fk_GBook_wholejoin=Rs("G_wholejoin")
		Fk_GBook_health=Rs("G_health")
		Fk_GBook_ip=Rs("G_ip")
		Fk_GBook_time=Rs("G_time")
		Fk_GBook_ispass=Rs("G_ispass")
		Fk_Bad_Reason=Rs("G_Bad_Reason")
		Fk_IsWhole=Rs("G_IsWhole")
		Fk_GBook_tj=Rs("G_tj")
		Fk_GBook_sex=Rs("G_sex")
	End If
	Rs.Close
%>
<form id="GBookRe" name="GBookRe" method="post" action="?Type=3">
<div id="BoxTop" style="width:98%;"><span>报名审核</span></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="23%" height="28" align="right">姓名：</td>
        <td width="27%" class="lm3">&nbsp;<%=Fk_GBook_Name%></td>
        <td width="25%" height="28" align="right">手机：</td>
        <td width="25%" class="lm3">&nbsp;<%=Fk_GBook_mobile%></td>
    </tr>
     <tr>
        <td height="28" align="right">性别：</td>
        <td  class="lm3">&nbsp;<%=Fk_GBook_sex%></td>
        <td height="28" align="right">推荐人：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_tj%></td>
    </tr>
     <tr>
        <td height="28" align="right">公司：</td>
        <td  class="lm3">&nbsp;<%=Fk_GBook_company%></td>
        <td height="28" align="right">职位：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_position%></td>
    </tr>
    <tr>
        <td height="28" align="right">地址：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_address%></td>
        <td height="28" align="right">年龄：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_age%></td>
    </tr>
    <tr>
        <td height="28" align="right">电子邮箱：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_email%></td>
        <td height="28" align="right">身份证/护照号码：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_cardid%></td>
    </tr>
    <tr>
        <td height="28" align="right">来源地区：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_area%></td>
        <td height="28" align="right">国籍：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_country%></td>
    </tr>
    <tr>
        <td height="28" align="right">出生年月日：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_birth%></td>
        <td height="28" align="right">学历：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_edu%></td>
    </tr>
    <tr>
        <td height="28" align="right">是否接受过传统文化培训：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_joinedbefore%></td>
        <td height="28" align="right">为什么想参加：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_why%></td>
    </tr>
    <tr>
        <td height="28" align="right">是否有家庭成员或朋友一起参加：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_familyjoined%></td>
        <td height="28" align="right">能否确保全程参加完课程培训：</td>
        <td class="lm3">&nbsp;<%=Fk_IsWhole%><%if Fk_IsWhole="有问题" then response.write "("&Fk_GBook_wholejoin&")"%></td>
    </tr>
    <tr>
        <td height="28" align="right">身体是否有健康问题或疾病：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_health%><%if Fk_GBook_health="有" then response.write "("&Fk_Bad_Reason&")"%></td>
        <td height="28" align="right">IP：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_ip%></td>
    </tr>
    <tr>
        <td height="28" align="right">申请时间：</td>
        <td class="lm3">&nbsp;<%=Fk_GBook_time%></td>
        <td height="28" align="right">审核状态：</td>
        <td class="lm3">&nbsp;<%if Fk_GBook_ispass=1 then response.Write "已通过审核" else response.Write "未审核"%></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="hidden" name="Id" value="<%=Id%>" />
        <input type="button" class="Button" name="btnprint" id="btnprint" value="打 印" onClick="window.print();"/>
        <input type="submit" class="Button" name="button" id="button" value="审 核" />
        <input type="button" onClick="window.location.href='<%=Session("NowPage")%>';" class="Button" name="btnclode" id="btnclode" value="关 闭" />
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
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	Id=Trim(Request.Form("Id"))
	Sqlstr="Select top 1 G_ispass From [API_Online_Registration] Where ID=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("G_ispass")=1
		Rs.Update()
		Application.UnLock()
	    response.Write "<script language=javascript>alert(""审核成功！"");window.location.href='"&Session("NowPage")&"'</script>"
	Else
	    response.Write "<script language=javascript>alert(""信息不存在！"");window.location.href='"&Session("NowPage")&"'</script>"
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：GBookDelDo
'作    用：执行删除互动
'参    数：
'==============================
Sub GBookDelDo()
	Id=Request("id")
	if id<>"" then
		on error resume next
		conn.execute("delete from [Fk_AD_Apply] where ID in (" & Id &")")
		if err then
			response.Write "信息删除失败！"
		else
			response.Write "信息删除成功！"
		end if
	Else
		response.Write "未选择要删除的信息！"
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->