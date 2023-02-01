<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Limit.asp
'文件用途：用户权限管理拉取页面
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
Dim Fk_Limit_Name,Fk_Limit_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call LimitList() '权限列表
	Case 2
		Call LimitAddForm() '添加权限表单
	Case 3
		Call LimitAddDo() '执行添加权限
	Case 4
		Call LimitEditForm() '修改权限表单
	Case 5
		Call LimitEditDo() '执行修改权限
	Case 6
		Call LimitDelDo() '执行删除权限
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：LimitList()
'作    用：权限列表
'参    数：
'==========================================
Sub LimitList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Limit.asp?Type=2');">添加新权限</a></li>
    </ul>
</div>
<div id="ListTop">
    权限管理
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">权限名称</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Sqlstr="Select * From [Fk_Limit] Order By Fk_Limit_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
        <tr>
            <td height="20" align="center"><%=Rs("Fk_Limit_Id")%></td>
            <td align="center"><%=Rs("Fk_Limit_Name")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Limit.asp?Type=4&Id=<%=Rs("Fk_Limit_Id")%>');">修改</a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Limit_Name")%>”，此操作不可逆！','Limit.asp?Type=6&Id=<%=Rs("Fk_Limit_Id")%>','MainRight','Limit.asp?Type=1');">删除</a></td>
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
'函 数 名：LimitAddForm()
'作    用：添加权限表单
'参    数：
'==========================================
Sub LimitAddForm()
%>
<form id="LimitAdd" name="LimitAdd" method="post" action="Limit.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">添加新权限[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">权限名称：</td>
	        <td>&nbsp;<input name="Fk_Limit_Name" type="text" class="Input" id="Fk_Limit_Name" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">权限详情：</td>
	        <td>
        	<p>&nbsp;1.系统权限</p>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System1" />(1).系统设置
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System2" />(2).友情连接管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System3" />(3).模板管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System4" />(4).招聘管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System5" />(5).缓存管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System6" />(6).菜单管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System7" />(7).专题管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System8" />(8).推荐管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System9" />(9).生成管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System10" />(10).广告管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System11" />(11).站内关键字管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System12" />(12).投票管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System13" />(13).独立信息管理
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System20" />(20).订单管理、会员功能
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System21" />(21).SEO模块功能
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System22" />(22).SEO模块管理功能
        	<p>&nbsp;2.栏目权限</p>
<%
	Dim MenuList
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof 
		If MenuList="" Then
			MenuList=Rs("Fk_Menu_Id")
		Else
			MenuList=MenuList&","&Rs("Fk_Menu_Id")
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	TempArr=Split(MenuList,",")
	For Each Temp In TempArr
		Call ModuleLimit(Temp)
		Response.Write("<p>&nbsp;&nbsp;</p>")
	Next
%>            
            </td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('LimitAdd','Limit.asp?Type=3',0,'',0,1,'MainRight','Limit.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：LimitAddDo
'作    用：执行添加权限
'参    数：
'==============================
Sub LimitAddDo()
	Fk_Limit_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Limit_Name")))
	Fk_Limit_Content=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Content"))," ",""))&","
	Call FKFun.ShowString(Fk_Limit_Name,1,50,0,"请输入权限名称！","权限名称不能大于50个字符！")
	Sqlstr="Select * From [Fk_Limit] Where Fk_Limit_Name='"&Fk_Limit_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Limit_Name")=Fk_Limit_Name
		Rs("Fk_Limit_Content")=Fk_Limit_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("新权限添加成功！")
	Else
		Response.Write("该权限名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：LimitEditForm()
'作    用：修改权限表单
'参    数：
'==========================================
Sub LimitEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Limit_Name=Rs("Fk_Limit_Name")
		Fk_Limit_Content=Rs("Fk_Limit_Content")
	End If
	Rs.Close
%>
<form id="LimitEdit" name="LimitEdit" method="post" action="Limit.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:700px;">修改权限[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">权限名称：</td>
	        <td>&nbsp;<input name="Fk_Limit_Name" value="<%=Fk_Limit_Name%>" type="text" class="Input" id="Fk_Limit_Name" /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">权限详情：</td>
	        <td>
        	<p>&nbsp;1.系统权限</p>
         <li> <input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System1"<%If Instr(Fk_Limit_Content,"System1")>0 Then%> checked="checked"<%End If%> />(1).系统设置</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System17"<%If Instr(Fk_Limit_Content,"System17")>0 Then%> checked="checked"<%End If%> />(17).关键词库管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System11"<%If Instr(Fk_Limit_Content,"System11")>0 Then%> checked="checked"<%End If%> />(11).关键词链接管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System2"<%If Instr(Fk_Limit_Content,"System2")>0 Then%> checked="checked"<%End If%> />(2).友情连接管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System5"<%If Instr(Fk_Limit_Content,"System5")>0 Then%> checked="checked"<%End If%> />(5).缓存管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System6"<%If Instr(Fk_Limit_Content,"System6")>0 Then%> checked="checked"<%End If%> />(6).菜单管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System8"<%If Instr(Fk_Limit_Content,"System8")>0 Then%> checked="checked"<%End If%> />(8).推荐管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System10"<%If Instr(Fk_Limit_Content,"System10")>0 Then%> checked="checked"<%End If%> />(10).JS广告调用管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System12"<%If Instr(Fk_Limit_Content,"System12")>0 Then%> checked="checked"<%End If%> />(12).投票管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System13"<%If Instr(Fk_Limit_Content,"System13")>0 Then%> checked="checked"<%End If%> />(13).独立信息管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System15"<%If Instr(Fk_Limit_Content,"System15")>0 Then%> checked="checked"<%End If%> />(15).SEO索引地图</li>        
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System14"<%If Instr(Fk_Limit_Content,"System14")>0 Then%> checked="checked"<%End If%> />(14).上传文件管理</li>
<li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System20"<%If Instr(Fk_Limit_Content,"System20")>0 Then%> checked="checked"<%End If%> />(20).订单管理、会员管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System21"<%If Instr(Fk_Limit_Content,"System21")>0 Then%> checked="checked"<%End If%> />(21).SEO模块功能
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System22"<%If Instr(Fk_Limit_Content,"System22")>0 Then%> checked="checked"<%End If%> />(22).SEO模块管理功能
           
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System18"<%If Instr(Fk_Limit_Content,"System18")>0 Then%> checked="checked"<%End If%> />(18).缩略水印设置</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System16"<%If Instr(Fk_Limit_Content,"System16")>0 Then%> checked="checked"<%End If%> />(16).过滤字符管理</li> 
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System3"<%If Instr(Fk_Limit_Content,"System3")>0 Then%> checked="checked"<%End If%> />(3).模板管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System4"<%If Instr(Fk_Limit_Content,"System4")>0 Then%> checked="checked"<%End If%> />(4).招聘管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System7"<%If Instr(Fk_Limit_Content,"System7")>0 Then%> checked="checked"<%End If%> />(7).专题管理</li>
           <li><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="System9"<%If Instr(Fk_Limit_Content,"System9")>0 Then%> checked="checked"<%End If%> />(9).生成管理</li>

        	<p>&nbsp;2.栏目权限</p>
<%
	Dim MenuList
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof 
		If MenuList="" Then
			MenuList=Rs("Fk_Menu_Id")
		Else
			MenuList=MenuList&","&Rs("Fk_Menu_Id")
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	TempArr=Split(MenuList,",")
	For Each Temp In TempArr
		Call ModuleLimit(Temp)
		Response.Write("<p>&nbsp;&nbsp;</p>")
	Next
%>            
            </td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:680px;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="submit" onclick="Sends('LimitEdit','Limit.asp?Type=5',0,'',0,1,'MainRight','Limit.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：LimitEditDo
'作    用：执行修改权限
'参    数：
'==============================
Sub LimitEditDo()
	Fk_Limit_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Limit_Name")))
	Fk_Limit_Content=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Limit_Content"))," ",""))&","
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Limit_Name,1,50,0,"请输入权限名称！","权限名称不能大于50个字符！")
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Limit_Name")=Fk_Limit_Name
		Rs("Fk_Limit_Content")=Fk_Limit_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("权限修改成功！")
	Else
		Response.Write("权限不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：LimitDelDo
'作    用：执行删除权限
'参    数：
'==============================
Sub LimitDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_Limit=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write("该权限尚在使用中，无法删除！")
		Response.End()
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Limit] Where Fk_Limit_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("权限删除成功！")
	Else
		Response.Write("权限不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ModuleLimit
'作    用：输出ModuleLimit列表
'参    数：要输出的菜单MenuIds
'==============================
Function ModuleLimit(MenuIds)
	Call ModuleLimitM(MenuIds,0,"")
End Function
Function ModuleLimitM(MenuIds,LevelId,TitleBack)
	Dim Rs2,TitleBacks,ii
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	ii=1
	While Not Rs2.Eof
	%>
		<p>&nbsp;&nbsp;&nbsp;&nbsp;<%=TitleBack%><input type="checkbox" class="Input" name="Fk_Limit_Content" id="Fk_Limit_Content" value="Module<%=Rs2("Fk_Module_Id")%>"<%If Instr(Fk_Limit_Content,"Module"&Rs2("Fk_Module_Id"))>0 Then%> checked="checked"<%End If%> />(<%=ii%>).<%=Rs2("Fk_Module_Name")%></p>
	<%
		If LevelId=0 Then
			TitleBacks="&nbsp;&nbsp;&nbsp;├"
		Else
			TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
		End If
		Call ModuleLimitM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks)
		Rs2.MoveNext
		ii=ii+1
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function
%><!--#Include File="../Code.asp"-->