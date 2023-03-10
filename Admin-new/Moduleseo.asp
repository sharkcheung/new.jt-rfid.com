<!--#Include File="CheckToken.asp"-->
<!--#Include File="../Class/Cls_Template.asp"-->
<%
'==========================================
'文 件 名：ModuleSEO.asp
'文件用途：栏目管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================


'定义页面变量
Dim MenuId
Dim Mkid
Dim Lmlx
Dim Fk_Menu_Name
Dim Fk_Module_Name,Fk_Module_Type,Fk_Module_Dir,Fk_Module_FileName,Fk_Module_Pic,Fk_Module_Menu,Fk_Module_Level,Fk_Module_LevelList,Fk_Module_Click,Fk_Module_Template,Fk_Module_LowTemplate,Fk_Module_Show,Fk_Module_Url,Fk_Module_Order,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_PageCount,Fk_Module_PageCode,Fk_Module_Seotitle
Dim FKTemplate
Set FKTemplate=New Cls_Template


'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ModuleList() '栏目列表
	Case 2
		Call ModuleAddForm() '添加栏目表单
	Case 3
		Call ModuleAddDo() '执行添加栏目
	Case 4
		Call ModuleEditForm() '修改栏目表单
	Case 5
		Call ModuleEditDo() '执行修改栏目
	Case 6
		Call ModuleDelDo() '执行删除栏目
	Case 7
		Call ModuleOrderForm() '栏目SEO表单
	Case 8
		Call ModuleOrderDo() '执行栏目SEO保存
	Case Else
		Call ModuleList() '栏目列表
End Select

'==========================================
'函 数 名：ModuleList()
'作    用：栏目列表
'参    数：
'==========================================
Sub ModuleList()
	'新功能，追加SEO title字段
	'2017年5月22日
	'middy241@163.com
	if CheckFields("Fk_Module_Seotitle","Fk_Down")=false then
		conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(500) null")
	end if
	Session("NowPage")=FkFun.GetNowUrl()
	MenuId=Clng(Request.QueryString("MenuId"))
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	End If
	Rs.Close
%>
<div id="BoxTop" style="width:98%;margin-left:8px; margin-top:8px;">
    <span>栏目SEO设置</span>
</div>
<div id="ListContent" style="width:98%;margin-left:8px;background-color:#FFF;border-left:1px solid #7998B7;border-right:1px solid #7998B7;border-bottom:1px solid #D7E0E7;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="left" class="ListTdTop">栏目名称</td>
            <td align="center" class="ListTdTop">是否显示</td>
            <td align="center" class="ListTdTop">SEO标题</td>
            <td align="center" class="ListTdTop">关键词(KeyWords)</td>
            <td align="left" class="ListTdTop">描述(Description)</td>
            <td align="center" style="display:none" class="ListTdTop">模板</td>
            <td align="left" class="ListTdTop">操作按钮</td>
        </tr>
<%
	Call ShowModuleListSEO(MenuId)
%>
        <tr>
            <td height="30" colspan="8">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom" style="display:none">

</div>
<%
End Sub

'==========================================
'函 数 名：ModuleAddForm()
'作    用：添加栏目表单
'参    数：
'==========================================
Sub ModuleAddForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Mkid=Clng(Request.QueryString("Mkid"))
	Lmlx=Request.QueryString("Lmlx")
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	End If
	Rs.Close
%>
<form id="ModuleAdd" name="ModuleAdd" method="post" action="ModuleSEO.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:600px;"><span>添加栏目</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:600px;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">栏目类型：</td>
        <td>&nbsp;        <select name="Fk_Module_Type" class="Input" id="Fk_Module_Type" onchange="ModuleTypeChange(this.options[this.options.selectedIndex].value);">
        <%if Lmlx<>"" then%>
           <option value="<%=Lmlx%>">同父目录</option>
            <%else%>
            <%	For i=0 To UBound(FKModuleId)%>
             <option value="<%=FKModuleId(i)%>"<%=FKFun.BeSelect(Fk_Module_Type,Clng(FKModuleId(i)))%>><%=FKModuleName(i)%></option>  
           <%	Next %>

            <%end if%> 
             </select>

        </td>
    </tr>
    <tr>
        <td height="28" align="right">栏目名称：</td>
        <td>&nbsp;<input name="Fk_Module_Name" type="text" class="Input" id="Fk_Module_Name"　<%If SiteToPinyin=1 Then%> onmousemove="GetPinyin('Fk_Module_FileName','ToPinyin.asp?Str='+this.value);" onmouseout="GetPinyin('Fk_Module_Dir','ToPinyin.asp?Str='+this.value);"<%End If%> size="30" /></td>
    </tr>
    <tr id="Fk_Module_Keywords">
        <td height="28" align="right">SEO标题：</td>
        <td>&nbsp;<input name="Fk_Module_Seotitle" type="text" class="Input" id="Fk_Module_Seotitle" size="40" /></td>
    </tr>
    <tr id="Fk_Module_Keywords">
        <td height="28" align="right">SEO关键词：</td>
        <td>&nbsp;<input name="Fk_Module_Keyword" type="text" class="Input" id="Fk_Module_Keyword" size="40" /></td>
    </tr>
    <tr id="Fk_Module_Descriptions">
        <td height="28" align="right">SEO描述：</td>
        <td>&nbsp;<input name="Fk_Module_Description" type="text" class="Input" id="Fk_Module_Description" size="60" /></td>
    </tr>
    <tr id="Fk_Module_Dirs" style="display:<% if Lmlx<>1 and Lmlx<>2 and Lmlx<>7 then response.write "none"%>;">
        <td height="28" align="right">生成目录：</td>
        <td>&nbsp;<input name="Fk_Module_Dir" type="text" class="Input" id="Fk_Module_Dir" size="40" />*不可修改</td>
    </tr>
    <tr id="Fk_Module_FileNames" style="display:<% if Lmlx=1 or Lmlx=2 or Lmlx=7 then response.write "none"%>">
        <td height="28" align="right">生成页面：</td>
        <td>&nbsp;<input name="Fk_Module_FileName" type="text" class="Input" id="Fk_Module_FileName" size="40" />
            *不可修改</td>
    </tr>
    <tr id="Fk_Module_Urls" style="display:none;">
        <td align="right" style="height: 28px">转向链接：</td>
        <td style="height: 28px">&nbsp;<input name="Fk_Module_Url" type="text" class="Input" id="Fk_Module_Url" size="60" /></td>
    </tr>
    <tr>
        <td height="28" align="right">栏目分级：</td>
        <td>&nbsp;<select name="Fk_Module_Level" class="Input" id="Fk_Module_Level">
            <option value="0">顶级栏目</option>

<%
	Call ShowModuleSelect(MenuId,Mkid)
	'Call ShowModuleSelect(MenuId,0)
%>
            </select>
        </td>
    </tr>
    <tr id="Fk_Module_PageCounts" style="display:none;">
        <td height="28" align="right">每页条数：</td>
        <td>&nbsp;<select name="Fk_Module_PageCount" class="Input" id="Fk_Module_PageCount">
            <option value="0">系统默认</option>
<%
	For i=1 To 50
%>
            <option value="<%=i%>"><%=i%>条</option>
<%
	Next
%>
            </select>
        </td>
       </tr>
    <tr id="Fk_Module_PageCodes" style="display:none;">
        <td height="28" align="right" <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>页码字符：</td>
        <td <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>&nbsp;<input name="Fk_Module_PageCode" value="第一页|--|上一页|--|下一页|--|尾页|--|条/页|--|共|--|页/|--|条|--|当前第|--|页|--|第|--|页" type="text" class="Input" id="Fk_Module_PageCode" size="60" />*请按格式修改</td>
    </tr>
    <tr id="Fk_Module_Templates">
        <td height="28" align="right" <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>显示模板：</td>
        <td <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
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
            </select>
        </td>
    </tr>
    <tr id="Fk_Module_LowTemplates" style="display:none;">
        <td height="28" align="right" <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>子内容模板：</td>
        <td <%If Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>&nbsp;<select name="Fk_Module_LowTemplate" class="Input" id="Fk_Module_LowTemplate">
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
            </select>
        </td>
    </tr>
    <tr>
        <td height="28" align="right">是否显示：</td>
        <td>&nbsp;<input name="Fk_Module_Show" type="radio" class="Input" id="Fk_Module_Show" value="1" checked="checked" />显示
        <input type="radio" name="Fk_Module_Show" class="Input" id="Fk_Module_Show" value="0" />不显示</td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:580px;">
		<input type="hidden" name="MenuId" value="<%=MenuId%>" />
        <input type="submit" onclick="Sends('ModuleAdd','ModuleSEO.asp?Type=3',0,'',0,1,'MainRight','ModuleSEO.asp?Type=1&MenuId=<%=MenuId%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ModuleAddDo
'作    用：执行添加栏目
'参    数：
'==============================
Sub ModuleAddDo()
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Seotitle")))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Fk_Module_Type=Trim(Request.Form("Fk_Module_Type"))
	Fk_Module_Level=Trim(Request.Form("Fk_Module_Level"))
	Fk_Module_Show=Trim(Request.Form("Fk_Module_Show"))
	Fk_Module_Menu=Trim(Request.Form("MenuId"))
	Call FKFun.ShowString(Fk_Module_Name,1,50,0,"请输入栏目名称！","栏目名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Seotitle,0,500,2,"请输入SEO标题！","SEO标题不能大于500个字符！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入SEO关键词！","SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入描述！","描述不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Type,"请选择栏目类型！")
	Call FKFun.ShowNum(Fk_Module_Level,"请选择栏目分级！")
	Call FKFun.ShowNum(Fk_Module_Show,"请选择栏目是否菜单显示！")
	Call FKFun.ShowNum(Fk_Module_Menu,"系统参数错误，请刷新页面！")
	Select Case Fk_Module_Type
		Case 0
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
		Case 1
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 2
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 3
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
		Case 4
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 5
			Fk_Module_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Url")))
			Call FKFun.ShowString(Fk_Module_Url,1,255,0,"请输入转向链接！","转向链接不能大于255个字符！")
		Case 7
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
	End Select
	If Fk_Module_Level>0 Then
		Fk_Module_LevelList=GetModuleLevelList(Fk_Module_Level)
	Else
		Fk_Module_LevelList=""
	End If
	If Instr(",0,3,",Fk_Module_Type)>0 And Fk_Module_FileName<>"" Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_FileName='"&Fk_Module_FileName&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Response.Write("生成页面已经被占用，请重新选择一个！")
			Rs.Close
			Call FKDB.DB_Close()
			Response.End()
		End If
		Rs.Close
	End If
	If Instr(",1,2,7,",Fk_Module_Type)>0 And Fk_Module_Dir<>"" Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Response.Write("生成目录已经被占用，请重新选择一个！")
			Rs.Close
			Call FKDB.DB_Close()
			Response.End()
		End If
		Rs.Close
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Name='"&Fk_Module_Name&"' And Fk_Module_Level="&Fk_Module_Level&""
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Seotitle")=Fk_Module_Seotitle
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		Rs("Fk_Module_Type")=Fk_Module_Type
		Rs("Fk_Module_Level")=Fk_Module_Level
		Rs("Fk_Module_Show")=Fk_Module_Show
		Rs("Fk_Module_LevelList")=Fk_Module_LevelList
		Rs("Fk_Module_Menu")=Fk_Module_Menu
		Select Case Fk_Module_Type
			Case 0
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_FileName")=Fk_Module_FileName
			Case 1
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 2
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 3
				Rs("Fk_Module_FileName")=Fk_Module_FileName
				Rs("Fk_Module_Template")=Fk_Module_Template
			Case 4
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_FileName")=Fk_Module_FileName
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 5
				Rs("Fk_Module_Url")=Fk_Module_Url
			Case 7
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
		End Select
		Rs("Fk_Module_Admin")=Request.Cookies("FkAdminId")
		Rs("Fk_Module_Ip")=Request.ServerVariables("REMOTE_ADDR")
		Rs.Update()
		Application.UnLock()
		Response.Write("新栏目添加成功！")
	Else
		Response.Write("该名称已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：ModuleEditForm()
'作    用：修改栏目表单
'参    数：
'==========================================
Sub ModuleEditForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Seotitle=Rs("Fk_Module_Seotitle")
		Fk_Module_Keyword=Rs("Fk_Module_Keyword")
		Fk_Module_Description=Rs("Fk_Module_Description")
		Fk_Module_Type=Rs("Fk_Module_Type")
		Fk_Module_Level=Rs("Fk_Module_Level")
		Fk_Module_Show=Rs("Fk_Module_Show")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_PageCount=Rs("Fk_Module_PageCount")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Module_Template=Rs("Fk_Module_Template")
		Fk_Module_LowTemplate=Rs("Fk_Module_LowTemplate")
		Fk_Module_FileName=Rs("Fk_Module_FileName")
		Fk_Module_Url=Rs("Fk_Module_Url")
		Fk_Module_PageCode=Rs("Fk_Module_PageCode")
	End If
	Rs.Close
%>
<form id="ModuleEdit" name="ModuleEdit" method="post" action="ModuleSEO.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:600px;"><span>修改栏目</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:600px;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">栏目类型：</td>
        <td>&nbsp;<select name="Fk_Module_Type" class="Input" id="Fk_Module_Type" onchange="ModuleTypeChange(this.options[this.options.selectedIndex].value);">

<%
	For i=0 To UBound(FKModuleId)
%>
                <option value="<%=FKModuleId(i)%>"<%=FKFun.BeSelect(Fk_Module_Type,Clng(FKModuleId(i)))%>><%=FKModuleName(i)%></option>
<%
	Next
%>
                </select>
        </td>
    </tr>
    <tr>
        <td height="28" align="right">栏目名称：</td>
        <td>&nbsp;<input name="Fk_Module_Name" type="text" class="Input" id="Fk_Module_Name" value="<%=Fk_Module_Name%>" size="30"　<%If SiteToPinyin=1 Then%> onmousemove="GetPinyin('Fk_Module_FileName','ToPinyin.asp?Str='+this.value);" onmouseout="GetPinyin('Fk_Module_Dir','ToPinyin.asp?Str='+this.value);" <%End If%> /></td>
    </tr>
    <tr id="Fk_Module_Keywords"<%If Instr(",5,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">SEO标题：</td>
        <td>&nbsp;<input name="Fk_Module_Seotitle" value="<%=Fk_Module_Seotitle%>" type="text" class="Input" id="Fk_Module_Seotitle" size="40" /></td>
    </tr>
    <tr id="Fk_Module_Keywords"<%If Instr(",5,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">SEO关键词：</td>
        <td>&nbsp;<input name="Fk_Module_Keyword" value="<%=Fk_Module_Keyword%>" type="text" class="Input" id="Fk_Module_Keyword" size="40" /></td>
    </tr>
    <tr id="Fk_Module_Descriptions"<%If Instr(",5,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">SEO描述：</td>
        <td>&nbsp;<input name="Fk_Module_Description" value="<%=Fk_Module_Description%>" type="text" class="Input" id="Fk_Module_Description" size="60" /></td>
    </tr>
    
    <tr id="Fk_Module_Dirs"<%If Instr(",0,3,4,5,6,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">生成目录：</td>
        <td>&nbsp;<input name="Fk_Module_Dir" type="text" class="Input" id="Fk_Module_Dir" value="<%=Fk_Module_Dir%>"<%If Fk_Module_Dir<>"" Then%> readonly="readonly"<%End If%> size="40"  />*一旦确立不可修改</td>
    </tr>
    <tr id="Fk_Module_FileNames"<%If Instr(",1,2,5,6,7,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">生成页面：</td>
        <td>&nbsp;<input name="Fk_Module_FileName" type="text" class="Input" id="Fk_Module_FileName" value="<%=Fk_Module_FileName%>"<%If Fk_Module_FileName<>"" Then%> readonly="readonly"<%End If%> size="40" />
            *一旦确立不可修改</td>
    </tr>
    <tr id="Fk_Module_Urls"<%If Instr(",0,1,2,4,6,7,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">栏目转向链接：</td>
        <td>&nbsp;<input name="Fk_Module_Url" type="text" class="Input" id="Fk_Module_Url" value="<%=Fk_Module_Url%>" size="60" /></td>
    </tr>
    <tr id="Fk_Module_PageCounts"<%If Instr(",0,3,5,6,",Fk_Module_Type)>0 Then%> style="display:none;"<%End If%>>
        <td align="right" style="height: 28px">每页条数：</td>
        <td style="height: 28px">&nbsp;<select name="Fk_Module_PageCount" class="Input" id="Fk_Module_PageCount">
            <option value="0">系统默认</option>
<%
	For i=1 To 50
%>
            <option value="<%=i%>"<%=FKFun.BeSelect(Fk_Module_PageCount,i)%>><%=i%>条</option>
<%
	Next
%>
            </select>
        </td>
        </tr>
    <tr id="Fk_Module_PageCodes"<%If Instr(",0,3,5,6,",Fk_Module_Type)>0 or Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">页码字符：</td>
        <td>&nbsp;<input name="Fk_Module_PageCode" type="text" class="Input" id="Fk_Module_PageCode" value="<%=Fk_Module_PageCode%>" size="60" />*请按格式修改</td>
    </tr>
    <tr>
        <td height="28" align="right">栏目分级：</td>
        <td>&nbsp;<select name="Fk_Module_Level" class="Input" id="Fk_Module_Level">
            <option value="0">顶级栏目</option>
<%
	Call ShowModuleSelect(MenuId,Fk_Module_Level)
%>
            </select>
        </td>
    </tr>
    <tr id="Fk_Module_Templates"<%If Instr(",5,",Fk_Module_Type)>0 or Request.Cookies("FkAdminLimitId")<>0 Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">显示模板：</td>
        <td>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
            <option value="0"<%=FKFun.BeSelect(Fk_Module_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>
        </td>
    </tr>
    <tr id="Fk_Module_LowTemplates"<%If Instr(",0,3,4,5,6,",Fk_Module_Type)>0 or Request.Cookies("FkAdminLimitId")<>0  Then%> style="display:none;"<%End If%>>
        <td height="28" align="right">子内容模板：</td>
        <td>&nbsp;<select name="Fk_Module_LowTemplate" class="Input" id="Fk_Module_LowTemplate">
            <option value="0"<%=FKFun.BeSelect(Fk_Module_LowTemplate,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_LowTemplate,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>
        </td>
    </tr>
    <tr>
        <td height="28" align="right">是否显示：</td>
        <td>&nbsp;<input name="Fk_Module_Show" type="radio" class="Input" id="Fk_Module_Show" value="1"<%=FKFun.BeCheck(Fk_Module_Show,1)%> />显示
        <input type="radio" name="Fk_Module_Show" class="Input" id="Fk_Module_Show" value="0"<%=FKFun.BeCheck(Fk_Module_Show,0)%> />不显示</td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:580px;text-align:right;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="MenuId" value="<%=MenuId%>" />
        <input type="submit" onclick="Sends('ModuleEdit','ModuleSEO.asp?Type=5',0,'',0,1,'MainRight','ModuleSEO.asp?Type=1&MenuId=<%=MenuId%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ModuleEditDo
'作    用：执行修改栏目
'参    数：
'==============================
Sub ModuleEditDo()
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Seotitle")))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Fk_Module_Type=Trim(Request.Form("Fk_Module_Type"))
	Fk_Module_Level=Trim(Request.Form("Fk_Module_Level"))
	Fk_Module_Show=Trim(Request.Form("Fk_Module_Show"))
	Fk_Module_Menu=Trim(Request.Form("MenuId"))
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Module_Name,1,50,0,"请输入栏目名称！","栏目名称不能大于50个字符！")
	Call FKFun.ShowString(Fk_Module_Seotitle,0,500,2,"请输入SEO标题！","SEO标题不能大于500个字符！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入SEO关键词！","SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入描述！","描述不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Type,"请选择栏目类型！")
	Call FKFun.ShowNum(Fk_Module_Level,"请选择栏目分级！")
	Call FKFun.ShowNum(Fk_Module_Show,"请选择栏目是否菜单显示！")
	Call FKFun.ShowNum(Fk_Module_Menu,"MenuId系统参数错误，请刷新页面！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Select Case Fk_Module_Type
		Case 0
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
		Case 1
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 2
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 3
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_Url=Trim(Request.Form("Fk_Module_Url"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
		Case 4
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入生成页面！","生成页面不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
		Case 5
			Fk_Module_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Url")))
			Call FKFun.ShowString(Fk_Module_Url,1,255,0,"请输入转向链接！","转向链接不能大于255个字符！")
		Case 7
			Fk_Module_Dir=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Dir")))
			Fk_Module_PageCode=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_PageCode")))
			Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
			Fk_Module_LowTemplate=Trim(Request.Form("Fk_Module_LowTemplate"))
			Fk_Module_PageCount=Trim(Request.Form("Fk_Module_PageCount"))
			Call FKFun.ShowString(Fk_Module_Dir,0,50,2,"请输入生成目录！","生成目录不能大于50个字符！")
			Call FKFun.ShowString(Fk_Module_PageCode,0,255,2,"请输入栏目页码字符！","栏目页码字符不能大于255个字符！")
			Call FKFun.ShowNum(Fk_Module_Template,"请选择栏目模板！")
			Call FKFun.ShowNum(Fk_Module_LowTemplate,"请选择栏目子内容模板！")
			Call FKFun.ShowNum(Fk_Module_PageCount,"请选择每页条数！")
			If UBound(Split(Fk_Module_PageCode,"|--|"))<11 Then
				Response.Write("页码字符请按格式编写！")
				Call FKDB.DB_Close()
				Response.End()
			End If
	End Select
	If Id=Fk_Module_Level Then
		Response.Write("自己不能成为自己的分类哦！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Fk_Module_Level>0 Then
		Fk_Module_LevelList=GetModuleLevelList(Fk_Module_Level)
	Else
		Fk_Module_LevelList=""
	End If
	If Instr(",0,3,",Fk_Module_Type)>0 And Fk_Module_FileName<>"" Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=3 And Fk_Module_Id<>"&Id&" And Fk_Module_FileName='"&Fk_Module_FileName&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Response.Write("生成页面已经被占用，请重新选择一个！")
			Rs.Close
			Call FKDB.DB_Close()
			Response.End()
		End If
		Rs.Close
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%' And Fk_Module_Id=" & Fk_Module_Level
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Response.Write("不能转移到子分类旗下！如需要请先中转一个分类！")
		Rs.Close
		Call FKDB.DB_Close()
		Response.End()
	End If
	Rs.Close
	If Instr(",1,2,",Fk_Module_Type)>0 And Fk_Module_Dir<>"" Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Dir='"&Fk_Module_Dir&"' And Fk_Module_Id<>"&Id&""
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			Response.Write("目录名已经被占用！")
			Rs.Close
			Call FKDB.DB_Close()
			Response.End()
		End If
		Rs.Close
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Dim OldLevelList
		OldLevelList=Rs("Fk_Module_LevelList")
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Seotitle")=Fk_Module_Seotitle
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		Rs("Fk_Module_Type")=Fk_Module_Type
		Rs("Fk_Module_Level")=Fk_Module_Level
		Rs("Fk_Module_Show")=Fk_Module_Show
		Rs("Fk_Module_LevelList")=Fk_Module_LevelList
		Rs("Fk_Module_Menu")=Fk_Module_Menu
		Select Case Fk_Module_Type
			Case 0
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_FileName")=Fk_Module_FileName
			Case 1
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 2
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 3
				Rs("Fk_Module_FileName")=Fk_Module_FileName
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_Url")=Fk_Module_Url
			Case 4
				Rs("Fk_Module_FileName")=Fk_Module_FileName
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
			Case 5
				Rs("Fk_Module_Url")=Fk_Module_Url
			Case 7
				Rs("Fk_Module_Dir")=Fk_Module_Dir
				Rs("Fk_Module_PageCode")=Fk_Module_PageCode
				Rs("Fk_Module_Template")=Fk_Module_Template
				Rs("Fk_Module_LowTemplate")=Fk_Module_LowTemplate
				Rs("Fk_Module_PageCount")=Fk_Module_PageCount
		End Select
		Rs.Update()
		Response.Write("栏目修改成功！")
		If Fk_Module_LevelList<>OldLevelList Then
			Rs.Close
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%'"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
				Rs("Fk_Module_LevelList")=Replace(Rs("Fk_Module_LevelList"),OldLevelList,Fk_Module_LevelList)
				Rs.Update
				Rs.MoveNext
			Wend
		End If
		Application.UnLock()
	Else
		Response.Write("栏目不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ModuleDelDo
'作    用：执行删除栏目
'参    数：
'==============================
Sub ModuleDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_LevelList Like '%%,"&Id&",%%'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write("此栏目有子栏目，暂无法删除！")
		Response.End()
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("栏目删除成功！")
	Else
		Response.Write("栏目不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ModuleOrderForm
'作    用：栏目排序表单
'参    数：
'==============================
Sub ModuleOrderForm()
	MenuId=Clng(Request.QueryString("MenuId"))
	Sqlstr="Select * From [Fk_Menu] Where Fk_Menu_Id=" & MenuId
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Menu_Name=Rs("Fk_Menu_Name")
	Else
		PageErr=1
	End If
	Rs.Close
%>
<!--#include file="head.asp"-->
<div class="page">
	<!--#include file="nav.asp"-->
	<div class="pagemian">
     	<div class="pagemian2">
			<!--#include file="leftlist.asp"-->
            <div class="pageright gjcbl" style="border-top:0;">
            	
				<form id="ModuleOrderSet" name="ModuleOrderSet" method="post" action="ModuleSEO.asp?Type=8" onsubmit="return false;">
                <table width="100%" cellspacing="0" cellpadding="0" border="0">
                    <tbody>
                    	<tr>
                        <th align="center" width="90">编号</th>
                        <th align="left">栏目名称</th>
						<th align="left">seo标题<span>(<50个汉字[100个字符])</span></th>
                        <th align="left" width="240">关键词<span>(<50个汉字[ 100个字符 ])</span></th>
                        <th align="left" width="240"> 描述<span>(<100汉字 [ 200个字符 ]，3句为宜 )</span> </th>
                    	</tr>
<%
	Call OrderModuleListseo(MenuId)
%>
						<tr class="last">
                        	<td colspan="4"> <input type="hidden" name="MenuId" value="<%=MenuId%>" /><input class="sz" name="" type="submit" value="设置" onclick="Sends('ModuleOrderSet','ModuleSEO.asp?Type=8',1,'Moduleseo.asp?type=7&MenuId=<%=MenuId%>',0,0,'','');" style="margin-right:20px"/><input name="" type="button" value="重置"/></td>
                        </tr>
                        
                        
                	</tbody>
                </table>
				</form>
            </div>
        </div>
     </div>

</div>

</body>
</html>
<%
End Sub

'==============================
'函 数 名：ModuleOrderDo
'作    用：执行栏目SEO保存
'参    数：
'==============================
Sub ModuleOrderDo
	MenuId=Trim(Request.Form("MenuId"))
	Call FKFun.ShowNum(MenuId,"MenuId系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuId&" Order By Fk_Module_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	While Not Rs.Eof
		Fk_Module_Keyword=Trim(Request.Form("Fk_Module_Keyword"&Rs("Fk_Module_Id")))
		Call FKFun.ShowString(Fk_Module_Keyword,1,100,2,"请输入栏目关键词！",Rs("Fk_Module_Name")&"关键词不能大于100个字符！")
		Fk_Module_Seotitle=Trim(Request.Form("Fk_Module_Seotitle"&Rs("Fk_Module_Id")))
		Call FKFun.ShowString(Fk_Module_Seotitle,1,255,2,"请输入栏目SEO标题！",Rs("Fk_Module_Name")&"SEO标题不能大于255个字符！")
		Fk_Module_Description=Trim(Request.Form("Fk_Module_Description"&Rs("Fk_Module_Id")))
		Call FKFun.ShowString(Fk_Module_Description,1,200,2,"请输入栏目描述！",Rs("Fk_Module_Name")&"描述不能大于200个字符！")
		'Call FKFun.ShowNum(Fk_Module_Order,Rs("Fk_Module_Name")&"栏目的序号不是数字，排序序号必须是有效数字！")
		if(trim(Fk_Module_Keyword&" ")<>"" and trim(Fk_Module_Seotitle&" ")<>"" and trim(Fk_Module_Description&" ")<>"") then
			Rs("Fk_Module_Keyword")=trim(Fk_Module_Keyword&" ")
			Rs("Fk_Module_Seotitle")=trim(Fk_Module_Seotitle&" ")
			Rs("Fk_Module_Description")=trim(Fk_Module_Description&" ")
			Rs.Update()
		end if
		Rs.MoveNext
	Wend
	Application.UnLock()
	Rs.Close
	Response.Write("栏目SEO设置完成")
End Sub
	
'==============================
'函 数 名：ShowModuleListSEO
'作    用：输出Module列表
'参    数：要输出的菜单MenuIds
'==============================
Public Function ShowModuleListSEO(MenuIds)
	Call ShowModuleListM(MenuIds,0,"")
End Function
Public Function ShowModuleListM(MenuIds,LevelId,TitleBack)
	Dim Rs2,Rs3,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Set Rs3=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack="<span class='lm2'>"
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
		If Rs2("Fk_Module_Template")>0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Rs2("Fk_Module_Template")
			Rs3.Open Sqlstr,Conn,1,3
			If Not Rs3.Eof Then
				Temp=Rs3("Fk_Template_Name")
			Else
				Temp="未知模板"
			End If
			Rs3.Close
		Else
			Temp="默认模板"
		End If
	%>
	<tr>
		<td height="22" align="center"><%=Rs2("Fk_Module_Id")%></td>
		<td align="left">&nbsp;&nbsp;<%=TitleBack%><%=Rs2("Fk_Module_Name")%></td>
		<td align="center">&nbsp;&nbsp;<img src="images/caidan<%response.write Rs2("Fk_Module_show")%>.png"></td>
		<td align="center"><%=Rs2("Fk_Module_Seotitle")%></td>
		<td align="center"><%=Rs2("Fk_Module_Keyword")%></td>
		<td align="left"><%=Rs2("Fk_Module_Description")%></td>
		<td align="center" style="display:none"><%=Temp%></td>
		<td align="left"><a title="修改栏目设置 " href="javascript:void(0);" onclick="ShowBox('ModuleSEO.asp?Type=4&MenuId=<%=MenuIds%>&Id=<%=Rs2("Fk_Module_Id")%>');"><img src="images/edit.png"></a></td>
		
	</tr>
	<%
		If LevelId=0 Then
			TitleBacks="&nbsp;&nbsp;<span class='lm1'>└&nbsp;"
		Else
			TitleBacks="&nbsp;&nbsp;&nbsp;&nbsp;<span class='lm4'>"&TitleBack&"</span>"
		End If
		Call ShowModuleListM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks)
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
	Set Rs3=Nothing
End Function
		
'==============================
'函 数 名：OrderModuleListseo
'作    用：输出Module排序操作列表
'参    数：要输出的菜单MenuIds
'==============================
Public Function OrderModuleListseo(MenuIds)
	Call OrderModuleListM(MenuIds,0,"")
End Function
Public Function OrderModuleListM(MenuIds,LevelId,TitleBack)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	''response.write TitleBack&"-----" &LevelId &"<br>"
	' If LevelId=0 Then
		' TitleBack="<span class='yiji'>"
	' End If
	If LevelId=0 Then
		TitleBack="<span class='lm2'>"
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
	%>
	<tr>
		<td align="center"><%=Rs2("Fk_Module_Id")%></td>
		<td class="no2"><%=TitleBack%><%=Rs2("Fk_Module_Name")%></td>
		<td align="center"><textarea name="Fk_Module_Seotitle<%=Rs2("Fk_Module_Id")%>" id="Fk_Module_Seotitle<%=Rs2("Fk_Module_Id")%>" cols="25" rows="5" class="TextArea2"><%=Rs2("Fk_Module_Seotitle")%></textarea></td>
		<td><textarea name="Fk_Module_Keyword<%=Rs2("Fk_Module_Id")%>" id="Fk_Module_Keyword<%=Rs2("Fk_Module_Id")%>" cols="40" rows="5" class="TextArea2"><%=Rs2("Fk_Module_keyword")%></textarea></td>
		<td><textarea name="Fk_Module_Description<%=Rs2("Fk_Module_Id")%>" id="Fk_Module_Description<%=Rs2("Fk_Module_Id")%>" cols="40" rows="5" class="TextArea2"><%=Rs2("Fk_Module_Description")%></textarea></td>
	</tr>
<%
		' If LevelId=0 Then
			' TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
		' Else
			' TitleBacks="<span class='erji'>└&nbsp;"
		' End If
		If LevelId=0 Then
			TitleBacks="&nbsp;&nbsp;<span class='lm1'>└&nbsp;"
		Else
			TitleBacks="&nbsp;&nbsp;&nbsp;&nbsp;<span class='lm4'>"&TitleBack&"</span>"
		End If
		Call OrderModuleListM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks)
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：ShowModuleSelect
'作    用：输出ModuleSelect列表
'参    数：要输出的菜单MenuIds
'==============================
Public Function ShowModuleSelect(MenuIds,AutoId)
	Call ShowModuleSelectM(MenuIds,0,"",AutoId)
End Function
Public Function ShowModuleSelectM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
	%>
					<option value="<%=Rs2("Fk_Module_Id")%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
	<%
		If LevelId=0 Then
			TitleBacks="&nbsp;&nbsp;&nbsp;├"
		Else
			TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
		End If
		Call ShowModuleSelectM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：GetModuleLevelList
'作    用：输出分类级数参数
'参    数：要输出的栏目ModuleLevelId
'==============================
Public Function GetModuleLevelList(ModuleLevelId)
	GetModuleLevelList=","&GetModuleLevelListM(ModuleLevelId)&ModuleLevelId&","
End Function
Public Function GetModuleLevelListM(ModuleLevelId)
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleLevelId
	Rs2.Open Sqlstr,Conn,1,3
	If Not Rs2.Eof Then
		If Rs2("Fk_Module_Level")>0 Then
			GetModuleLevelListM=GetModuleLevelListM(Rs2("Fk_Module_Level"))&Rs2("Fk_Module_Level")&","
		Else
			GetModuleLevelListM=""
		End If
	End If
	Rs2.Close
	Set Rs2=Nothing
End Function
%>
<!--#Include File="../Code.asp"-->