<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Friends.asp
'文件用途：友情链接管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System2") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Friends_Name,Fk_Friends_About,Fk_Friends_Url,Fk_Friends_Logo,Fk_Friends_ShowType,Fk_Friends_FriendsType

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FriendsList() '友情链接列表
	Case 2
		Call FriendsAddForm() '添加友情链接表单
	Case 3
		Call FriendsAddDo() '执行添加友情链接
	Case 4
		Call FriendsEditForm() '修改友情链接表单
	Case 5
		Call FriendsEditDo() '执行修改友情链接
	Case 6
		Call FriendsDelDo() '执行删除友情链接
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FriendsList()
'作    用：友情链接列表
'参    数：
'==========================================
Sub FriendsList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Friends.asp?Type=2');">添加</a></li>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','FriendsType.asp?Type=1');return false">类型设置</a></li>
    </ul>
</div>
<div id="ListTop">
    友情链接管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">站点名称</td>
            <td align="center" class="ListTdTop">站点LOGO</td>
            <td align="center" class="ListTdTop">显示模式</td>
            <td align="center" class="ListTdTop">链接类型</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_FriendsList] Order By Fk_Friends_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Friends_Id")%></td>
            <td align="center"><%=Rs("Fk_Friends_Name")%></td>
            <td align="center"><%If Rs("Fk_Friends_Logo")<>"" Then%><img src="<%=Rs("Fk_Friends_Logo")%>" width="88" height="31" /><%Else%>无LOGO<%End If%></td>
            <td align="center"><%If Rs("Fk_Friends_ShowType")=1 Then%>LOGO<%Else%>文字<%End If%></td>
            <td align="center"><%=Rs("Fk_FriendsType_Name")%></td>
            <td align="center"><a href="javascript:void(0);" title="修改 " onclick="ShowBox('Friends.asp?Type=4&Id=<%=Rs("Fk_Friends_Id")%>');"><img src="images/edit.png"></a> <a title="删除 " href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Friends_Name")%>”，此操作不可逆！','Friends.asp?Type=6&Id=<%=Rs("Fk_Friends_Id")%>','MainRight','Friends.asp?Type=1');"><img src="images/del.png"></a></td>
        </tr>
<%
			Rs.MoveNext
		Wend
	Else
%>
        <tr>
            <td height="25" colspan="6" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
        <tr>
            <td height="30" colspan="6">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FriendsAddForm()
'作    用：添加友情链接表单
'参    数：
'==========================================
Sub FriendsAddForm()
%>
<form id="FriendsAdd" name="FriendsAdd" method="post" action="Friends.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加友情链接</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">站点名称：</td>
	        <td>&nbsp;<input name="Fk_Friends_Name" type="text" class="Input" id="Fk_Friends_Name" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点地址：</td>
	        <td>&nbsp;<input name="Fk_Friends_Url" type="text" class="Input" id="Fk_Friends_Url" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点介绍：</td>
	        <td>&nbsp;<input name="Fk_Friends_About" type="text" class="Input" id="Fk_Friends_About" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点LOGO：</td>
	        <td>&nbsp;<input name="Fk_Friends_Logo" type="text" class="Input" id="Fk_Friends_Logo" size="60" /><br />
        &nbsp;<iframe frameborder="0" width="290" height="25" scrolling="No" id="Fk_Friends_Logos" name="Fk_Friends_Logos" src="PicUpLoad.asp?Form=FriendsAdd&Input=Fk_Friends_Logo"></iframe></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接类型：</td>
	        <td>&nbsp;<select name="Fk_Friends_FriendsType" class="Input" id="Fk_Friends_FriendsType">
<%
	Sqlstr="Select * From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_FriendsType_Id")%>"><%=Rs("Fk_FriendsType_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">显示模式：</td>
	        <td>&nbsp;<input name="Fk_Friends_ShowType" class="Input" type="radio" id="Fk_Friends_ShowType" value="1" checked="checked" />LOGO
            <input type="radio" name="Fk_Friends_ShowType" class="Input" id="Fk_Friends_ShowType" value="2" />文字</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('FriendsAdd','Friends.asp?Type=3',0,'',0,1,'MainRight','Friends.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsAddDo
'作    用：执行添加友情链接
'参    数：
'==============================
Sub FriendsAddDo()
	Fk_Friends_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Name")))
	Fk_Friends_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_About")))
	Fk_Friends_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Url")))
	Fk_Friends_Logo=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Logo")))
	Fk_Friends_ShowType=Trim(Request.Form("Fk_Friends_ShowType"))
	Fk_Friends_FriendsType=Trim(Request.Form("Fk_Friends_FriendsType"))
	Call FKFun.ShowString(Fk_Friends_Name,1,255,0,"请输入友情链接名称！","友情链接名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_About,1,255,2,"请输入友情链接介绍！","友情链接介绍不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Url,1,255,0,"请输入友情链接地址！","友情链接地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Logo,1,255,2,"请输入友情链接LOGO！","友情链接LOGO不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Friends_ShowType,"请选择友情链接显示类型！")
	Call FKFun.ShowNum(Fk_Friends_FriendsType,"请选择友情链接类型！")
	Sqlstr="Select * From [Fk_Friends] Where Fk_Friends_Name='"&Fk_Friends_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Friends_Name")=Fk_Friends_Name
		Rs("Fk_Friends_About")=Fk_Friends_About
		Rs("Fk_Friends_Url")=Fk_Friends_Url
		Rs("Fk_Friends_Logo")=Fk_Friends_Logo
		Rs("Fk_Friends_ShowType")=Fk_Friends_ShowType
		Rs("Fk_Friends_FriendsType")=Fk_Friends_FriendsType
		Rs.Update()
		Application.UnLock()
		Response.Write("新友情链接添加成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="添加友情链接：【"&Fk_Friends_Name&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("该名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：FriendsEditForm()
'作    用：修改友情链接表单
'参    数：
'==========================================
Sub FriendsEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Friends_Name=Rs("Fk_Friends_Name")
		Fk_Friends_About=Rs("Fk_Friends_About")
		Fk_Friends_Url=Rs("Fk_Friends_Url")
		Fk_Friends_Logo=Rs("Fk_Friends_Logo")
		Fk_Friends_ShowType=Rs("Fk_Friends_ShowType")
		Fk_Friends_FriendsType=Rs("Fk_Friends_FriendsType")
	End If
	Rs.Close
%>
<form id="FriendsEdit" name="FriendsEdit" method="post" action="Friends.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改友情链接</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">站点名称：</td>
	        <td>&nbsp;<input name="Fk_Friends_Name" value="<%=Fk_Friends_Name%>" type="text" class="Input" id="Fk_Friends_Name" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点地址：</td>
	        <td>&nbsp;<input name="Fk_Friends_Url" value="<%=Fk_Friends_Url%>" type="text" class="Input" id="Fk_Friends_Url" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点介绍：</td>
	        <td>&nbsp;<input name="Fk_Friends_About" value="<%=Fk_Friends_About%>" type="text" class="Input" id="Fk_Friends_About" size="35" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">站点LOGO：</td>
	        <td>&nbsp;<input name="Fk_Friends_Logo" value="<%=Fk_Friends_Logo%>" type="text" class="Input" id="Fk_Friends_Logo" size="60" /><br />
        &nbsp;<iframe frameborder="0" width="290" height="25" scrolling="No" id="Fk_Friends_Logos" name="Fk_Friends_Logos" src="PicUpLoad.asp?Form=FriendsEdit&Input=Fk_Friends_Logo"></iframe></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">链接类型：</td>
	        <td>&nbsp;<select name="Fk_Friends_FriendsType" class="Input" id="Fk_Friends_FriendsType">
<%
	Sqlstr="Select * From [Fk_FriendsType] Order By Fk_FriendsType_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_FriendsType_Id")%>"<%=FKFun.BeSelect(Rs("Fk_FriendsType_Id"),Fk_Friends_FriendsType)%>><%=Rs("Fk_FriendsType_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
                </select></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">显示模式：</td>
	        <td>&nbsp;<input name="Fk_Friends_ShowType" class="Input" type="radio" id="Fk_Friends_ShowType" value="1"<%=FKFun.BeCheck(Fk_Friends_ShowType,1)%> />LOGO
            <input type="radio" name="Fk_Friends_ShowType" class="Input" id="Fk_Friends_ShowType" value="2"<%=FKFun.BeCheck(Fk_Friends_ShowType,2)%> />文字</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('FriendsEdit','Friends.asp?Type=5',0,'',0,1,'MainRight','Friends.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FriendsEditDo
'作    用：执行修改友情链接
'参    数：
'==============================
Sub FriendsEditDo()
	Fk_Friends_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Name")))
	Fk_Friends_About=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_About")))
	Fk_Friends_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Url")))
	Fk_Friends_Logo=FKFun.HTMLEncode(Trim(Request.Form("Fk_Friends_Logo")))
	Fk_Friends_ShowType=Trim(Request.Form("Fk_Friends_ShowType"))
	Fk_Friends_FriendsType=Trim(Request.Form("Fk_Friends_FriendsType"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Friends_Name,1,255,0,"请输入友情链接名称！","友情链接名称不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_About,1,255,2,"请输入友情链接介绍！","友情链接介绍不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Url,1,255,0,"请输入友情链接地址！","友情链接地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Friends_Logo,1,255,2,"请输入友情链接LOGO！","友情链接LOGO不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Friends_ShowType,"请选择友情链接显示类型！")
	Call FKFun.ShowNum(Fk_Friends_FriendsType,"请选择友情链接类型！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Friends_Name")=Fk_Friends_Name
		Rs("Fk_Friends_About")=Fk_Friends_About
		Rs("Fk_Friends_Url")=Fk_Friends_Url
		Rs("Fk_Friends_Logo")=Fk_Friends_Logo
		Rs("Fk_Friends_ShowType")=Fk_Friends_ShowType
		Rs("Fk_Friends_FriendsType")=Fk_Friends_FriendsType
		Rs.Update()
		Application.UnLock()
		Response.Write("友情链接修改成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="修改友情链接：【"&Fk_Friends_Name&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("友情链接不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：FriendsDelDo
'作    用：执行删除友情链接
'参    数：
'==============================
Sub FriendsDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Friends] Where Fk_Friends_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		log_content="删除友情链接：【"&rs("Fk_Friends_Name")&"】"
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("友情链接删除成功！")		
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("友情链接不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->