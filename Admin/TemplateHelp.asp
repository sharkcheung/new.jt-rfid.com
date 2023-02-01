<!--#Include File="AdminCheck.asp"--><head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
</head>

<%
'==========================================
'文 件 名：TemplateHelp.asp
'文件用途：模版标签生成器拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System3") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call TemplateHelpBox() '读取标签生成器
	Case 2
		Call GetTemplate() '读取标签
	Case 3
		Call GetTemplate2() '读取标签2
End Select

'==========================================
'函 数 名：LoginBox()
'作    用：读取登录信息
'参    数：
'==========================================
Sub TemplateHelpBox()
%>


<div id="ListTop">
    模板标签生成器
</div>
<div id="ListContent">
<table width="98%" border="0" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="22" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=1');">全站常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=30');">统计和客服代码</a><!--<a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=2');">静态页常规标签</a>--></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=3');">单页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=4');">新闻列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=6');">新闻页常规标签</a></td>
        </tr>
    <tr>
        <td height="22" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=5');">图文列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=7');">图文页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=12');">下载列表页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=13');">下载页常规标签</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=8');">留言类页专用标签</a></td>
    </tr>
    <tr>
        <td height="22" align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=10');">IF标签使用方法</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=11');">搜索页标签</a></td>
        <td align="center"><a style="color: #CCCCCC;" href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=14');">招聘页标签(暂)</a></td>
        <td align="center"><a style="color: #CCCCCC;" href="javascript:void(0);" onclick="SetRContent('Template','TemplateHelp.asp?Type=2&Id=9');">专题页标签(暂)</a></td>
        <td align="center"></td>
    </tr>
<form id="Get1" name="Get1" method="post" action="TemplateHelp.asp?Type=3&Id=1" onsubmit="return false;">
    <tr>
        <td height="22" align="right">菜单/分类标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId1" class="Input" id="MenuId1" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId1');">
            <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>
            <select name="ModuleId1" class="Input" id="ModuleId1">
                <option value="">请先选择菜单</option>
                </select>
            <select name="GetCount1" class="Input" id="GetCount1">
                <option value="1">读取1级</option>
                <option value="2">读取2级</option>
                <option value="3">读取3级</option>
                <option value="4">读取4级</option>
                <option value="5">读取5级</option>
                </select>
            <select name="GetCount2" class="Input" id="GetCount2">
                <option value="0">无需回溯</option>
                <option value="-1">回溯1级</option>
                <option value="-2">回溯2级</option>
                <option value="-3">回溯3级</option>
                <option value="-4">回溯4级</option>
                <option value="-5">回溯5级</option>
                </select>
            <input type="submit" onclick="Sends_Div('Get1','TemplateHelp.asp?Type=3&Id=1','Template');" class="Button" name="button2" id="button2" value="生 成" />
            </td>
        </tr>
</form>
<form id="Get2" name="Get2" method="post" action="TemplateHelp.asp?Type=3&Id=2" onsubmit="return false;">
    <tr>
        <td height="22" align="right">新闻列表标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId2" class="Input" id="MenuId2" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId2');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId2" class="Input" id="ModuleId2">
                <option value="">请先选择菜单</option>
        </select>
        <select name="GetType1" class="Input" id="GetType1">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="GetCount3" class="Input" id="GetCount3">
                <option value="0">无需设置条数（分页模式）</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select><br />
         <select name="GetType2" class="Input" id="GetType2">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="GetType3" class="Input" id="GetType3">
                <option value="0">非推荐新闻</option>
<%
	Sqlstr="Select * From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Recommend_Id")%>"><%=Rs("Fk_Recommend_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
         <select name="GetType4" class="Input" id="GetType4">
                <option value="0">非专题新闻</option>
<%
	Sqlstr="Select * From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Subject_Id")%>"><%=Rs("Fk_Subject_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="GetCount7" class="Input" id="GetCount7">
                <option value="0">读取全部标题</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>字</option>
<%
	Next
%>
        </select>
       <input type="submit" onclick="Sends_Div('Get2','TemplateHelp.asp?Type=3&Id=2','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
        </tr>
</form>
<form id="Get3" name="Get3" method="post" action="TemplateHelp.asp?Type=3&Id=3" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;图文列表标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId3" class="Input" id="MenuId3" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId3');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId3" class="Input" id="ModuleId3">
                <option value="">请先选择菜单</option>
        </select>
        <select name="GetType5" class="Input" id="GetType5">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="GetCount4" class="Input" id="GetCount4">
                <option value="0">无需设置条数（分页模式）</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select><br />
         <select name="GetType6" class="Input" id="GetType6">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="GetType7" class="Input" id="GetType7">
                <option value="0">非推荐图文</option>
<%
	Sqlstr="Select * From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Recommend_Id")%>"><%=Rs("Fk_Recommend_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
         <select name="GetType8" class="Input" id="GetType8">
                <option value="0">非专题图文</option>
<%
	Sqlstr="Select * From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Subject_Id")%>"><%=Rs("Fk_Subject_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="GetCount8" class="Input" id="GetCount8">
                <option value="0">读取全部标题</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>字</option>
<%
	Next
%>
        </select>
       <input type="submit" onclick="Sends_Div('Get3','TemplateHelp.asp?Type=3&Id=3','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get8" name="Get8" method="post" action="TemplateHelp.asp?Type=3&Id=8" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;下载列表标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId5" class="Input" id="MenuId5" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId5');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId5" class="Input" id="ModuleId5">
                <option value="">请先选择菜单</option>
        </select>
        <select name="GetType15" class="Input" id="GetType15">
                <option value="0">按ID倒序</option>
                <option value="1">按时间倒序</option>
                <option value="2">按点击倒序</option>
                <option value="3">按ID正序</option>
                <option value="4">按时间正序</option>
                <option value="5">按点击正序</option>
        </select>
        <select name="GetCount9" class="Input" id="GetCount9">
                <option value="0">无需设置条数（分页模式）</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select><br />
         <select name="GetType16" class="Input" id="GetType16">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
         <select name="GetType17" class="Input" id="GetType17">
                <option value="0">非推荐下载</option>
<%
	Sqlstr="Select * From [Fk_Recommend]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Recommend_Id")%>"><%=Rs("Fk_Recommend_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
         <select name="GetType18" class="Input" id="GetType18">
                <option value="0">非专题下载</option>
<%
	Sqlstr="Select * From [Fk_Subject]"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Subject_Id")%>"><%=Rs("Fk_Subject_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="GetCount10" class="Input" id="GetCount10">
                <option value="0">读取全部标题</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>字</option>
<%
	Next
%>
        </select>
       <input type="submit" onclick="Sends_Div('Get8','TemplateHelp.asp?Type=3&Id=8','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get11" name="Get11" method="post" action="TemplateHelp.asp?Type=3&Id=11" onsubmit="return false;">
   <tr>
       <td height="22" align="right">留言类列表标签生成：</td>
       <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId7" class="Input" id="MenuId7" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId7');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId7" class="Input" id="ModuleId7">
                <option value="">请先选择菜单</option>
        </select>
        <select name="GetCount11" class="Input" id="GetCount11">
                <option value="0">无需设置条数（分页模式）</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select>
        <select name="GetType21" class="Input" id="GetType21">
                <option value="0">所有记录</option>
                <option value="1">已回复记录</option>
        </select>
         <select name="GetType22" class="Input" id="GetType22">
                <option value="0">不分页</option>
                <option value="1">分页</option>
        </select>
       <input type="submit" onclick="Sends_Div('Get11','TemplateHelp.asp?Type=3&Id=11','Template');" class="Button" name="button" id="button" value="生 成" />            
       </td>
   </tr>
</form>
<form id="Get4" name="Get4" method="post" action="TemplateHelp.asp?Type=3&Id=4" onsubmit="return false;">
   <tr>
        <td height="22" align="right">&nbsp;&nbsp;友情链接标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
         <select name="GetType9" class="Input" id="GetType9">
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
        </select>
        <select name="GetType10" class="Input" id="GetType10">
                <option value="1">LOGO模式</option>
                <option value="2">文字模式</option>
        </select>
        <select name="GetCount5" class="Input" id="GetCount5">
                <option value="0">所有</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select>
       <input type="submit" onclick="Sends_Div('Get4','TemplateHelp.asp?Type=3&Id=4','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
        </tr>
</form>
<form id="Get5" name="Get5" method="post" action="TemplateHelp.asp?Type=3&Id=5" onsubmit="return false;">
    <tr style="display:none">
        <td height="22" align="right">&nbsp;&nbsp;招聘列表标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="GetCount6" class="Input" id="GetCount6">
                <option value="0">读取所有</option>
<%
	For i=1 To 20
%>
                <option value="<%=i%>">读取<%=i%>条</option>
<%
	Next
%>
        </select>
        <select name="GetType11" class="Input" id="GetType11">
                <option value="0">所有</option>
                <option value="1">有效</option>
                <option value="2">无效</option>
        </select>
       <input type="submit" onclick="Sends_Div('Get5','TemplateHelp.asp?Type=3&Id=5','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get6" name="Get6" method="post" action="TemplateHelp.asp?Type=3&Id=6" onsubmit="return false;">
    <tr style="display:none">
        <td height="22" align="right">&nbsp;&nbsp;专题列表标签生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
       <input type="submit" onclick="Sends_Div('Get6','TemplateHelp.asp?Type=3&Id=6','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get7" name="Get7" method="post" action="TemplateHelp.asp?Type=3&Id=7" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;轮换代码生成：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId4" class="Input" id="MenuId4" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId4');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId4" class="Input" id="ModuleId4">
                <option value="">请先选择菜单</option>
        </select>
        <select name="GetType12" class="Input" id="GetType12">
                <option value="1">默认FLASH轮换方案</option>
        </select>
        宽度：<input type="text" name="GetType13" id="GetType13" size="1" class="Input" />
        高度：<input type="text" name="GetType14" id="GetType14" size="1" class="Input" />
       <input type="submit" onclick="Sends_Div('Get7','TemplateHelp.asp?Type=3&Id=7','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get9" name="Get9" method="post" action="TemplateHelp.asp?Type=3&Id=9" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;弹出留言框代码：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId6" class="Input" id="MenuId6" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId6');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId6" class="Input" id="ModuleId6">
                <option value="">请先选择菜单</option>
        </select>
       <input type="submit" onclick="Sends_Div('Get9','TemplateHelp.asp?Type=3&Id=9','Template');" class="Button" name="button" id="button" value="生 成" />
       *必须选择留言模块，否则代码无效     
        </td>
    </tr>
</form>
<form id="Get10" name="Get10" method="post" action="TemplateHelp.asp?Type=3&Id=10" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;首页模块（原信息块）：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="GetType19" class="Input" id="GetType19">
                <option value="">请选择模块信息</option>
<%
	Sqlstr="Select * From [Fk_Info] Order By Fk_Info_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Info_Id")%>"><%=Rs("Fk_Info_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
       <input type="submit" onclick="Sends_Div('Get10','TemplateHelp.asp?Type=3&Id=10','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
<form id="Get15" name="Get15" method="post" action="TemplateHelp.asp?Type=3&Id=15" onsubmit="return false;">
    <tr>
        <td height="22" align="right">&nbsp;&nbsp;More链接：</td>
        <td height="22" colspan="4" style="padding:5px;">
        <select name="MenuId52" class="Input" id="MenuId52" onchange="ChangeSelect('Ajax.asp?Type=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId52');">
                <option value="">请选择菜单</option>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
                <option value="<%=Rs("Fk_Menu_Id")%>"><%=Rs("Fk_Menu_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </select>
        <select name="ModuleId52" class="Input" id="ModuleId52">
                <option value="">请先选择菜单</option>
        </select>
       <input type="submit" onclick="Sends_Div('Get15','TemplateHelp.asp?Type=3&Id=15','Template');" class="Button" name="button" id="button" value="生 成" />     
        </td>
    </tr>
</form>
    <tr>
        <td height="22" colspan="5" align="center">&nbsp;&nbsp;标签生成结果</td>
    </tr>
    <tr>
        <td height="22" colspan="5" id="Template" style="padding:10px; line-height:22px; font-size:14px;"></td>
        </tr>
</table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：GetTemplate()
'作    用：读取标签
'参    数：
'==========================================
Sub GetTemplate()
	Id=Clng(Request.QueryString("Id"))
	Select Case Id
		Case 1 '全站常规标签
%>
<p><input type="text" class="Input" size="50" value="{$SiteName$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点名称</p>
<p><input type="text" class="Input" size="50" value="{$SiteSeoTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点SEO标题</p>
<p><input type="text" class="Input" size="50" value="{$SiteUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点链接</p>
<p><input type="text" class="Input" size="50" value="{$SiteKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点关键字</p>
<p><input type="text" class="Input" size="50" value="{$SiteDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点描述</p>
<p><input type="text" class="Input" size="50" value="{$SiteSkin$}" />&nbsp;&nbsp;&nbsp;&nbsp;模板路径</p>
<p><input type="text" class="Input" size="50" value="{$SiteDir$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点路径</p>
<p><input type="text" class="Input" size="50" value="{$ImgCdnUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;图片资源CDN</p>
<p><input type="text" class="Input" size="50" value="{$CssCdnUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;JS资源CDN</p>
<p><input type="text" class="Input" size="50" value="{$JsCdnUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;CSS资源CDN</p>
<p><input type="text" class="Input" size="50" value="{$FileCdnUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;文件资源CDN</p>

<p><input type="text" class="Input" size="50" value="{$Tel$}" />&nbsp;&nbsp;&nbsp;&nbsp;普通电话</p>
<p><input type="text" class="Input" size="50" value="{$Tel400$}" />&nbsp;&nbsp;&nbsp;&nbsp;400电话</p>
<p><input type="text" class="Input" size="50" value="{$Fax$}" />&nbsp;&nbsp;&nbsp;&nbsp;传真号码</p>
<p><input type="text" class="Input" size="50" value="{$Add$}" />&nbsp;&nbsp;&nbsp;&nbsp;地址</p>
<p><input type="text" class="Input" size="50" value="{$Lianxiren$}" />&nbsp;&nbsp;&nbsp;&nbsp;联系人</p>
<p><input type="text" class="Input" size="50" value="{$Beian$}" />&nbsp;&nbsp;&nbsp;&nbsp;ICP备案</p>
<p><input type="text" class="Input" size="50" value="{$Email$}" />&nbsp;&nbsp;&nbsp;&nbsp;Email</p>
<p><input type="text" class="Input" size="50" value="{$SiteLogo$}" />&nbsp;&nbsp;&nbsp;&nbsp;LOGO标志</p>
<p><input type="text" class="Input" size="50" value="{$Sitepic1$}" />&nbsp;&nbsp;&nbsp;&nbsp;第1张幻灯图地址(1,2,3...)</p>
<p><input type="text" class="Input" size="50" value="{$Sitepicurl1$}" />&nbsp;&nbsp;&nbsp;&nbsp;第1张幻灯图链接(1,2,3...)</p>
<p><input type="text" class="Input" size="50" value="{$Sitepictext1$}" />&nbsp;&nbsp;&nbsp;&nbsp;第1张幻灯图标题(1,2,3...)</p>


<p><input type="text" class="Input" size="50" value="{$SitePageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前模块</p>
<p><input type="text" class="Input" size="50" value="{$PageNows$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前位置</p>
<p><input type="text" class="Input" size="50" value="{$File(区块名称)$}" />&nbsp;&nbsp;&nbsp;&nbsp;调用区块模板文件</p>
<%
		Case 30 '静态页常规标签
%>
<p><input type="text" class="Input" size="120" value="<div style='display:none;float:right;'><script type='text/javascript' src='{$TjUrl$}/counter.asp?id={$Tjid$}&icon=2'></script></div>" />&nbsp;&nbsp;&nbsp;&nbsp;统计代码</p>
<p><input type="text" class="Input" size="120" value="<script src='{$KfUrl$}/kf/?u={$Kfid$}'  charset='gb2312'></script>" />&nbsp;&nbsp;&nbsp;&nbsp;在线客服代码</p>


<%
		Case 2 '静态页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$PageTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;静态页标题</p>
<p><input type="text" class="Input" size="50" value="{$PageKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;静态页关键字</p>
<p><input type="text" class="Input" size="50" value="{$PageDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;静态页描述</p>
<%
		Case 3 '信息页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$InfoTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;信息标题</p>
<p><input type="text" class="Input" size="50" value="{$InfoKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;信息关键字</p>
<p><input type="text" class="Input" size="50" value="{$InfoDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;信息描述</p>
<p><input type="text" class="Input" size="50" value="{$InfoContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;信息内容</p>
<%
		Case 4 '新闻列表页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$ModuleContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块介绍</p>
<p><input type="text" class="Input" size="50" value="{$ArticleCategoryName$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ArticleCategoryKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻模块关键字</p>
<p><input type="text" class="Input" size="50" value="{$ArticleCategoryDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻模块描述</p>
<p><input type="text" class="Input" size="50" value="{$ArticleCategoryPage$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻模块页码</p>
<p><input type="text" class="Input" size="50" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="50" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="50" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="50" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 5 '图文列表页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$ModuleContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块介绍</p>
<p><input type="text" class="Input" size="50" value="{$ProductCategoryName$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ProductCategoryKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文模块关键字</p>
<p><input type="text" class="Input" size="50" value="{$ProductCategoryDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文模块描述</p>
<p><input type="text" class="Input" size="50" value="{$ProductCategoryPage$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文模块页码</p>
<p><input type="text" class="Input" size="50" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="50" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="50" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="50" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 6 '新闻页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$ArticleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻编号</p>
<p><input type="text" class="Input" size="50" value="{$ArticleTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻标题</p>
<p><input type="text" class="Input" size="50" value="{$ArticlePic$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻题图小图</p>
<p><input type="text" class="Input" size="50" value="{$ArticlePicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻题图大图</p>
<p><input type="text" class="Input" size="50" value="{$ArticleKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻关键字</p>
<p><input type="text" class="Input" size="50" value="{$ArticleDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻描述</p>
<p><input type="text" class="Input" size="50" value="{$ArticleFrom$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻来源</p>
<p><input type="text" class="Input" size="50" value="{$ArticleContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻内容</p>
<p><input type="text" class="Input" size="50" value="{$ArticleClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻点击量</p>
<p><input type="text" class="Input" size="50" value="{$ArticleTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加时间</p>
<p><input type="text" class="Input" size="50" value="{$ArticlePrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$ArticlePrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="50" value="{$ArticleNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$ArticleNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="100" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=1&Id={$ArticleId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;新闻HTML点击JS，放置页面底部</p>
<%
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
%>
<p><input type="text" class="Input" size="50" value="{$Article_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
				Rs.MoveNext
			Wend
			Rs.Close
		Case 7 '图文页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$ProductId$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文编号</p>
<p><input type="text" class="Input" size="50" value="{$ProductTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文名称</p>
<p><input type="text" class="Input" size="50" value="{$ProductKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文关键字</p>
<p><input type="text" class="Input" size="50" value="{$ProductDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文描述</p>

<p><input type="text" class="Input" size="50" value="{$ProductPicSummary$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文内容摘要 <b style="color:red">(New)</b></p>
<p><input type="text" class="Input" size="50" value="{$ProductPicSlidesFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示第一张图片URL <b style="color:red">(New)</b></p>
<p><input type="text" class="Input" size="50" value="{$ProductPicSlidesImgs$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示图片切换遍历代码 <b style="color:red">(New)</b></p>
<p><input type="text" class="Input" size="50" value="{$ProductPicSlidesImgList$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示图片集 <b style="color:red">(New)</b></p>
<p><input type="text" class="Input" size="50" value="{$Product_ContentEx1$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文扩展详细内容1 <b style="color:red">(New)</b></p>
<p><input type="text" class="Input" size="50" value="{$Product_ContentEx2$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文扩展详细内容2 <b style="color:red">(New)</b></p>

<p><input type="text" class="Input" size="50" value="{$ProductContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文简介</p>
<p><input type="text" class="Input" size="50" value="{$ProductClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文点击量</p>
<p><input type="text" class="Input" size="50" value="{$ProductTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加时间</p>
<p><input type="text" class="Input" size="50" value="{$ProductDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加日期</p>
<p><input type="text" class="Input" size="50" value="{$ProductMonth$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加月份</p>
<p><input type="text" class="Input" size="50" value="{$ProductDay$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加日</p>
<p><input type="text" class="Input" size="50" value="{$ProductPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文图片小图</p>
<p><input type="text" class="Input" size="50" value="{$ProductPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文图片大图</p>
<p><input type="text" class="Input" size="50" value="{$ProductPrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$ProductPrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="50" value="{$ProductNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$ProductNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="100" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=2&Id={$ProductId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;图文HTML点击JS，放置页面底部</p>
<%
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
%>
<p><input type="text" class="Input" size="50" value="{$Product_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
				Rs.MoveNext
			Wend
			Rs.Close
		Case 8 '留言页专用标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$GBookTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言模块标题</p>
<p><input type="text" class="Input" size="50" value="{$GBookKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言模块关键字</p>
<p><input type="text" class="Input" size="50" value="{$GBookDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言模块描述</p>
<p><input type="text" class="Input" size="50" value="{$GBookPage$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言模块页码</p>
<p><input type="text" class="Input" size="50" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="50" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="50" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="50" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 9 '专题页标签
%>
<p><input type="text" class="Input" size="50" value="{$SubjectId$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题编号</p>
<p><input type="text" class="Input" size="50" value="{$SubjectName$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题名称</p>
<%
		Case 10 'IF标签使用方法
%>
<p><input type="text" class="Input" size="50" value="{$If(参数1,参数2,比较方式)$}" />
&nbsp;&nbsp;&nbsp;&nbsp;IF标签开始，比较方式支持&lt;/&gt;/=/&gt;=/&lt;=/&lt;&gt;</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="如果成立输出的HTML" /></p>
<p><input type="text" class="Input" size="50" value="{$Else$}" /></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="如果不成立输出的HTML" /></p>
<p><input type="text" class="Input" size="50" value="{$End If$}" />&nbsp;&nbsp;&nbsp;&nbsp;IF标签结束</p>
<%
		Case 11 '搜索页标签
%>
<p><input type="text" class="Input" size="50" value="{$SearchStr$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索关键字</p>
<p><input type="text" class="Input" size="50" value="{$SearchType$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索类型</p>
<p><input type="text" class="Input" size="50" value="{$SearchPage$}" />&nbsp;&nbsp;&nbsp;&nbsp;搜索页码</p>
<%
		Case 12 '下载列表页常规标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$DownCategoryName$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载模块名称</p>
<p><input type="text" class="Input" size="50" value="{$DownCategoryKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载模块关键字</p>
<p><input type="text" class="Input" size="50" value="{$DownCategoryDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载模块描述</p>
<p><input type="text" class="Input" size="50" value="{$DownCategoryPage$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载模块页码</p>
<p><input type="text" class="Input" size="50" value="{$PageFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;第一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PagePrev$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNext$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageLast$}" />&nbsp;&nbsp;&nbsp;&nbsp;最后一页URL</p>
<p><input type="text" class="Input" size="50" value="{$PageNow$}" />&nbsp;&nbsp;&nbsp;&nbsp;当前页码</p>
<p><input type="text" class="Input" size="50" value="{$PageCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;页数</p>
<p><input type="text" class="Input" size="50" value="{$PageRecordCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;总记录数</p>
<p><input type="text" class="Input" size="50" value="{$PageSize$}" />&nbsp;&nbsp;&nbsp;&nbsp;每页数量</p>
<%
		Case 13 '下载页标签
%>
<p><input type="text" class="Input" size="50" value="{$MenuId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块菜单</p>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$DownId$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载编号</p>
<p><input type="text" class="Input" size="50" value="{$DownTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载名称</p>
<p><input type="text" class="Input" size="50" value="{$DownKeyword$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载关键字</p>
<p><input type="text" class="Input" size="50" value="{$DownDescription$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载描述</p>
<p><input type="text" class="Input" size="50" value="{$DownContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载简介</p>
<p><input type="text" class="Input" size="50" value="{$DownSystem$}" />&nbsp;&nbsp;&nbsp;&nbsp;适用系统</p>
<p><input type="text" class="Input" size="50" value="{$DownLanguage$}" />&nbsp;&nbsp;&nbsp;&nbsp;语言</p>
<p><input type="text" class="Input" size="50" value="{$DownFile$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载链接</p>
<p><input type="text" class="Input" size="50" value="{$DownClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载点击量</p>
<p><input type="text" class="Input" size="50" value="{$DownCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载量</p>
<p><input type="text" class="Input" size="50" value="{$DownTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加时间</p>
<p><input type="text" class="Input" size="50" value="{$DownDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加日期</p>
<p><input type="text" class="Input" size="50" value="{$DownMonth$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加月份</p>
<p><input type="text" class="Input" size="50" value="{$DownDay$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加日</p>
<p><input type="text" class="Input" size="50" value="{$DownPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载图片小图</p>
<p><input type="text" class="Input" size="50" value="{$DownPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载图片大图</p>
<p><input type="text" class="Input" size="50" value="{$DownPrevTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$DownPrevUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;上一篇链接</p>
<p><input type="text" class="Input" size="50" value="{$DownNextTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇标题</p>
<p><input type="text" class="Input" size="50" value="{$DownNextUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下一篇链接</p>
<p><input type="text" class="Input" size="50" value="<script type=&quot;text/javascript&quot; src=&quot;{$SiteDir$}Click.asp?Type=3&Id={$DownId$}&quot;></script>" />&nbsp;&nbsp;&nbsp;&nbsp;下载HTML点击JS，放置页面底部</p>
<%
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
			Rs.Open Sqlstr,Conn,1,3
			While Not Rs.Eof
%>
<p><input type="text" class="Input" size="50" value="{$Down_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
				Rs.MoveNext
			Wend
			Rs.Close
		Case 14 '招聘页标签
%>
<p><input type="text" class="Input" size="50" value="{$ModuleId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleFId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块父栏目编号</p>
<p><input type="text" class="Input" size="50" value="{$ModuleName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p><input type="text" class="Input" size="50" value="{$ModuleUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p><input type="text" class="Input" size="50" value="{$JobTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘页名称</p>
<%
		Case Else
			Response.Write("没有该标签哦！")
	End Select
End Sub

'==========================================
'函 数 名：GetTemplate2()
'作    用：读取标签2
'参    数：
'==========================================
Sub GetTemplate2() '
	Id=Clng(Request.QueryString("Id"))
	If Id=1 Then '菜单标签生成
		Dim MenuId1,ModuleId1,GetCount1,GetCount2
		MenuId1=Trim(Request.Form("MenuId1"))
		If MenuId1="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId1=Trim(Request.Form("ModuleId1"))
		GetCount1=Trim(Request.Form("GetCount1"))
		GetCount2=Trim(Request.Form("GetCount2"))
%>
<p><input type="text" class="Input" size="50" value="{$For(Nav,<%=MenuId1%>/<%=ModuleId1%>/<%=GetCount1%>/<%=GetCount2%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavId$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavName$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavI$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavType$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块类型</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$Nav_Content$}" />&nbsp;&nbsp;&nbsp;&nbsp;模块介绍 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$NavSub$}" />&nbsp;&nbsp;&nbsp;&nbsp;二级菜单标签</p>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=2 Then '新闻列表标签生成
		Dim MenuId2,ModuleId2,GetType1,GetCount3,GetType2,GetType3,GetType4,GetCount7
		MenuId2=Trim(Request.Form("MenuId2"))
		If MenuId2="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId2=Trim(Request.Form("ModuleId2"))
		GetType1=Trim(Request.Form("GetType1"))
		GetCount3=Trim(Request.Form("GetCount3"))
		GetType2=Trim(Request.Form("GetType2"))
		GetType3=Trim(Request.Form("GetType3"))
		GetType4=Trim(Request.Form("GetType4"))
		GetCount7=Trim(Request.Form("GetCount7"))
%>
<p><input type="text" class="Input" size="50" value="{$For(ArticleList,<%=MenuId2%>/<%=ModuleId2%>/<%=GetType1%>/<%=GetCount3%>/<%=GetType2%>/<%=GetType3%>/<%=GetType4%>/<%=GetCount7%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo2$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻内容缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加时间</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加日期</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListYear$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加年份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListMonth$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加月份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListDay$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻添加日</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻发布时间差</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;新闻点击量</p>
<%
		Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
		Rs.Open Sqlstr,Conn,1,3
		While Not Rs.Eof
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ArticleList_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
			Rs.MoveNext
		Wend
		Rs.Close
%>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=3 Then '图文列表标签生成
		Dim MenuId3,ModuleId3,GetType5,GetCount4,GetType6,GetType7,GetType8,GetCount8
		MenuId3=Trim(Request.Form("MenuId3"))
		If MenuId3="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId3=Trim(Request.Form("ModuleId3"))
		GetType5=Trim(Request.Form("GetType5"))
		GetCount4=Trim(Request.Form("GetCount4"))
		GetType6=Trim(Request.Form("GetType6"))
		GetType7=Trim(Request.Form("GetType7"))
		GetType8=Trim(Request.Form("GetType8"))
		GetCount8=Trim(Request.Form("GetCount8"))
%>
<p><input type="text" class="Input" size="50" value="{$For(ProductList,<%=MenuId3%>/<%=ModuleId3%>/<%=GetType5%>/<%=GetCount4%>/<%=GetType6%>/<%=GetType7%>/<%=GetType8%>/<%=GetCount8%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo2$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文内容缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPicSummary$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文内容摘要 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPicSlidesFirst$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示第一张图片URL <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPicSlidesImgs$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示图片切换遍历代码 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPicSlidesImgList$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文多图展示图片集 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductList_ContentEx1$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文扩展详细内容1 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductList_ContentEx2$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文扩展详细内容2 <b style="color:red">(New)</b></p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文点击量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加时间</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加日期</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListYear$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加年份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListMonth$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加月份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文发布时间差</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductListDay$}" />&nbsp;&nbsp;&nbsp;&nbsp;图文添加日</p>
<%
		Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
		Rs.Open Sqlstr,Conn,1,3
		While Not Rs.Eof
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ProductList_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
			Rs.MoveNext
		Wend
		Rs.Close
%>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=4 Then '友情链接列表标签生成
		Dim GetType9,GetType10,GetCount5
		GetType9=Trim(Request.Form("GetType9"))
		If GetType9="" Then
			Response.Write("请选择友情链接类型！")
			Response.End()
		End If
		GetType10=Trim(Request.Form("GetType10"))
		GetCount5=Trim(Request.Form("GetCount5"))
%>
<p><input type="text" class="Input" size="50" value="{$For(FriendsList,<%=GetType9%>/<%=GetType10%>/<%=GetCount5%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$FriendsName$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$FriendsAbout$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点简介</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$FriendsUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$FriendsLogo$}" />&nbsp;&nbsp;&nbsp;&nbsp;站点LOGO</p>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=5 Then '招聘列表标签生成
		Dim GetCount6,GetType11
		GetCount6=Trim(Request.Form("GetCount6"))
		GetType11=Trim(Request.Form("GetType11"))
%>
<p><input type="text" class="Input" size="50" value="{$For(JobList,<%=GetCount6%>/<%=GetType11%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobName$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘数量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobAbout$}" />&nbsp;&nbsp;&nbsp;&nbsp;招聘简介</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobArea$}" />&nbsp;&nbsp;&nbsp;&nbsp;工作地点</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;有效期限</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$JobTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;发布时间</p>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=6 Then '专题列表标签生成
%>
<p><input type="text" class="Input" size="50" value="{$For(SubjectList,1)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$SubjectListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$SubjectListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题图片</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$SubjectListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;专题链接</p>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=7 Then '轮换代码生成
		Dim MenuId4,ModuleId4,GetType12,GetType13,GetType14
		MenuId4=Trim(Request.Form("MenuId4"))
		If MenuId4="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId4=Trim(Request.Form("ModuleId4"))
		GetType12=Trim(Request.Form("GetType12"))
		GetType13=Trim(Request.Form("GetType13"))
		GetType14=Trim(Request.Form("GetType14"))
%>
<p><input type="text" class="Input" size="80" value="<script type='text/javascript' src='{$SiteDir$}Flash/V.asp?Type=<%=GetType12%>&Menu=<%=MenuId4%>&Module=<%=ModuleId4%>&Width=<%=GetType13%>&Height=<%=GetType14%>'></script>" />&nbsp;&nbsp;&nbsp;&nbsp;Flash轮换代码</p>
<%
	ElseIf Id=8 Then '下载列表标签生成
		Dim MenuId5,ModuleId5,GetType15,GetCount9,GetType16,GetType17,GetType18,GetCount10
		MenuId5=Trim(Request.Form("MenuId5"))
		If MenuId5="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId5=Trim(Request.Form("ModuleId5"))
		GetType15=Trim(Request.Form("GetType15"))
		GetCount9=Trim(Request.Form("GetCount9"))
		GetType16=Trim(Request.Form("GetType16"))
		GetType17=Trim(Request.Form("GetType17"))
		GetType18=Trim(Request.Form("GetType18"))
		GetCount10=Trim(Request.Form("GetCount10"))
%>
<p><input type="text" class="Input" size="50" value="{$For(DownList,<%=MenuId5%>/<%=ModuleId5%>/<%=GetType15%>/<%=GetCount9%>/<%=GetType16%>/<%=GetType17%>/<%=GetType18%>/<%=GetCount10%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo$}" />&nbsp;&nbsp;&nbsp;&nbsp;序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ListNo2$}" />&nbsp;&nbsp;&nbsp;&nbsp;带分页的序号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块名称</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$ModuleListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;所属模块链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListId$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载编号</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListTitleAll$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载标题（全部标题）</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载简介缩略</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListUrl$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载链接</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListPic$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图小图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListPicBig$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载题图大图</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListSystem$}" />&nbsp;&nbsp;&nbsp;&nbsp;适用系统</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListLanguage$}" />&nbsp;&nbsp;&nbsp;&nbsp;语言</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListFile$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载地址</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListClick$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载点击量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListCount$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载量</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加时间</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListDate$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加日期</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListYear$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加年份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListMonth$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加月份</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListNew$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载发布时间差</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownListDay$}" />&nbsp;&nbsp;&nbsp;&nbsp;下载添加日</p>
<%
		Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
		Rs.Open Sqlstr,Conn,1,3
		While Not Rs.Eof
%>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$DownList_<%=Rs("Fk_Field_Tag")%>$}" />&nbsp;&nbsp;&nbsp;&nbsp;<%=Rs("Fk_Field_Name")%></p>
<%
			Rs.MoveNext
		Wend
		Rs.Close
%>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%
	ElseIf Id=9 Then '弹出留言框代码生成
		Dim MenuId6,ModuleId6
		MenuId6=Trim(Request.Form("MenuId6"))
		If MenuId6="" Then
			Response.Write("请先选择菜单！")
			Response.End()
		End If
		ModuleId6=Trim(Request.Form("ModuleId6"))
%>
<p><input type="text" class="Input" size="100" value="<script type='text/javascript' src='{$SiteDir$}GBookJs.asp?Id=<%=ModuleId6%>'></script>" />&nbsp;&nbsp;&nbsp;&nbsp;弹出留言板代码</p>
<%
	ElseIf Id=10 Then '独立信息生成
		Dim GetType19
		GetType19=Trim(Request.Form("GetType19"))
		If GetType19="" Then
			Response.Write("请先独立信息！")
			Response.End()
		End If
%>
<p><input type="text" class="Input" size="80" value="{$Info(<%=GetType19%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;独立信息标签</p>
<%
	ElseIf Id=11 Then '留言列表生成
		Dim ModuleId7,GetCount11,GetType21,GetType22
		ModuleId7=Trim(Request.Form("ModuleId7"))
		GetCount11=Trim(Request.Form("GetCount11"))
		GetType21=Trim(Request.Form("GetType21"))
		GetType22=Trim(Request.Form("GetType22"))
		If ModuleId7="" Then
			Response.Write("请先独立信息！")
			Response.End()
		End If
%>
<p><input type="text" class="Input" size="50" value="{$For(GBookList,<%=ModuleId7%>/<%=GetCount11%>/<%=GetType21%>/<%=GetType22%>)$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环开始</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListTitle$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言标题</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListName$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言者</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;留言时间</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListReContent$}" />&nbsp;&nbsp;&nbsp;&nbsp;回复内容</p>
<p style="margin-left:20px;"><input type="text" class="Input" size="50" value="{$GBookListReTime$}" />&nbsp;&nbsp;&nbsp;&nbsp;回复时间</p>
<p><input type="text" class="Input" size="50" value="{$Next$}" />&nbsp;&nbsp;&nbsp;&nbsp;For循环结束</p>
<%

	ElseIf Id=15 Then 'More链接生成
		Dim ModuleId52
		ModuleId52=Trim(Request.Form("ModuleId52"))
		If ModuleId52="" Then
			Response.Write("请选择！")
			Response.End()
		End If
%>
<p><input type="text" class="Input" size="50" value="{$HomeUrlMore(<%=ModuleId52%>)$}" /></p>
<%

	Else
		Response.Write("没有该标签哦！")
	End If
End Sub
%><!--#Include File="../Code.asp"-->