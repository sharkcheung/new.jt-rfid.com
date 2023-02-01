<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"--><%
'==========================================
'文 件 名：HTML.asp
'文件用途：HTML生成拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System9") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义常量
Dim FKHTML,Fk_Subject_Name,CategoryDirName
Set FKHTML=New Cls_HTML
Dim jj

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call HTMLBox() '读取HTML生成器
	Case 2
		Call HTMLDo1() '生成常规HTML
	Case 3
		Call HTMLDo2() '生成选择项HTML
	Case 4
		Call HTMLDo3() '一键生成HTML
End Select

'==========================================
'函 数 名：HTMLBox()
'作    用：读取HTML生成器
'参    数：
'==========================================
Sub HTMLBox()
%>
<div id="BoxTop" style="width:700px;">HTML生成器[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:700px;">
<%
	If SiteHtml=1 Then
%>
<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=1';">生成首页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=2';">生成所有信息页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=3';">生成所有静态页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=4';">生成所有留言页</a></td>
        </tr>
    <tr>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=7';">生成所有文章页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=8';">生成所有产品页</a></td>
        <td height="25" align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=2&Id=9';">生成所有下载页</a></td>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=4&amp;Id=1';">一键今日更新</a></td>
    </tr>
    <tr>
        <td align="center"><a href="javascript:void(0);" onclick="document.getElementById('Gets').src='HTML.asp?Type=4&amp;Id=2';">一键2日内更新</a></td>
        <td align="center"></td>
        <td align="center"></td>
        <td align="center"></td>
    </tr>
    <tr>
        <td height="30" align="center">单独模块生成：</td>
        <td height="30" colspan="3" style="padding:5px;">
        <select name="MenuId" class="Input" id="MenuId" onchange="ChangeSelect('Ajax.asp?Type=1&Temp=1&Id='+this.options[this.options.selectedIndex].value,'ModuleId');">
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
            <select name="ModuleId" class="Input" id="ModuleId">
                <option value="">请先选择菜单</option>
                </select><br />
            <select name="CreatType" class="Input" id="CreatType">
                <option value="0">生成分类页+内容页</option>
                <option value="1">只生成分类页</option>
                <option value="2">只生成内容页</option>
                </select>
            <select name="CreatDay" class="Input" id="CreatDay">
                <option value="0">生成所有</option>
                <option value="1">生成1天内</option>
                <option value="2">生成2天内</option>
                <option value="7">生成7天内</option>
                </select>
            <input type="button" onclick="document.getElementById('Gets').src='HTML.asp?Type=3&MenuId='+document.all.MenuId.options[document.all.MenuId.selectedIndex].value+'&ModuleId='+document.all.ModuleId.options[document.all.ModuleId.selectedIndex].value+'&CreatType='+document.all.CreatType.options[document.all.CreatType.selectedIndex].value+'&CreatDay='+document.all.CreatDay.options[document.all.CreatDay.selectedIndex].value;" class="Button" name="button2" id="button2" value="生 成" />
            </td>
        </tr>
    <tr>
        <td height="25" colspan="4" id="Template" style="padding:10px; line-height:22px; font-size:14px;"><iframe src="" id="Gets" width="600px" height="200px"></iframe></td>
        </tr>
</table>
<%
	Else
%>
<p style="text-align:center;">系统设置为动态模式，无需HTML生成！</p>
<%
	End If
%>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==============================
'函 数 名：HTMLDo1()
'作    用：生成常规HTML
'参    数：
'==============================
Sub HTMLDo1()
	'定义变量
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Id=Clng(Request.QueryString("Id"))
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
	If Id=1 Then
		Call FKHTML.CreatIndex()
	ElseIf Id=2 Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=3 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			Call FKHTML.CreatInfo(Rs2("Fk_Module_Template"),Rs2("Fk_Module_FileName"),Rs2("Fk_Module_Name"),0)
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=3 Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=0 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			Call FKHTML.CreatPage(Rs2("Fk_Module_Template"),Rs2("Fk_Module_FileName"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=4 Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=4 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatGBook(Rs2("Fk_Module_Template"),Rs2("Fk_Module_FileName"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=5 Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=6 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			Call FKHTML.CreatJob(Rs2("Fk_Module_Template"),Rs2("Fk_Module_FileName"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=6 Then
		Sqlstr="Select * From [Fk_Subject] Order By Fk_Subject_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		While Not Rs2.Eof
			Id=Rs2("Fk_Subject_Id")
			Fk_Subject_Name=Rs2("Fk_Subject_Name")
			FKHTML.CreatSubject(Rs2("Fk_Subject_Template"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=7 Then
		Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Show=1 Order By Fk_Article_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				Id=Rs2("Fk_Article_Id")
				Call FKHTML.CreatArticle(Rs2("Fk_Article_Template"),Rs2("Fk_Article_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Article_FileName"),Rs2("Fk_Article_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=2&Id=7&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write("文章部分生成完毕！<br />")
			Response.Flush()
			Response.Clear()
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=1 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatArticleCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=8 Then
		Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Show=1 Order By Fk_Product_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				Id=Rs2("Fk_Product_Id")
				Call FKHTML.CreatProduct(Rs2("Fk_Product_Template"),Rs2("Fk_Product_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Product_FileName"),Rs2("Fk_Product_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=2&Id=8&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write("产品生成完毕！<br />")
			Response.Flush()
			Response.Clear()
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=2 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatProductCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	ElseIf Id=9 Then
		Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Show=1 Order By Fk_Down_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				Id=Rs2("Fk_Down_Id")
				Call FKHTML.CreatDown(Rs2("Fk_Down_Template"),Rs2("Fk_Down_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Down_FileName"),Rs2("Fk_Down_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=2&Id=9&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write("下载生成完毕！<br />")
			Response.Flush()
			Response.Clear()
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=7 Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatDownCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
	End If
	Set Rs2=Nothing
	Response.Write("生成完毕！")
End Sub

'==============================
'函 数 名：HTMLDo2()
'作    用：生成选择项HTML
'参    数：
'==============================
Sub HTMLDo2()
	'定义变量
	Dim PageCode,PageCodes,TemplateId,TemplateId2,MenuId,ModuleId,ModuleType,ModuleDir,ModuleFileName,ModuleName
	Dim CreatType,CreatDay,Pages
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	MenuId=Request.QueryString("MenuId")
	ModuleId=Request.QueryString("ModuleId")
	CreatType=Request.QueryString("CreatType")
	CreatDay=Request.QueryString("CreatDay")
	If MenuId="" Then
		Response.Write("请先选择主菜单！")
		Response.End()
	Else
		MenuId=Clng(MenuId)
	End If
	If ModuleId="" Then
		Response.Write("请选择需要生成的模块！")
		Response.End()
	Else
		ModuleId=Clng(ModuleId)
	End If
	If CreatType="" Then
		CreatType=0
	Else
		CreatType=Clng(CreatType)
	End If
	If CreatDay="" Then
		CreatDay=0
	Else
		CreatDay=Clng(CreatDay)
	End If
	PageNow=Request.QueryString("Page")
	If PageNow<>"" Then
		PageNow=Clng(PageNow)
	Else
		PageNow=1
	End If
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuId&" And Fk_Module_Id=" & ModuleId
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		TemplateId=Rs("Fk_Module_Template")
		TemplateId2=TemplateId
		ModuleType=Rs("Fk_Module_Type")
		ModuleDir=Rs("Fk_Module_Dir")
		ModuleFileName=Rs("Fk_Module_FileName")
		ModuleName=Rs("Fk_Module_Name")
		Id=ModuleId
		If Rs("Fk_Module_PageCode")<>"" Then
			PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
		End If
		If Rs("Fk_Module_PageCount")>0 Then
			PageSizes=Rs("Fk_Module_PageCount")
		End If
		Rs.Close
		If ModuleType=0 Then
			Call FKHTML.CreatPage(TemplateId,ModuleFileName,ModuleName)
		ElseIf ModuleType=1 Then
			If CreatType=0 Or CreatType=2 Then
				Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Menu="&MenuId&" And Fk_Module_Id=" & ModuleId
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And DateDiff('d',Fk_Article_Time,'"&Now()&"')<="&CreatDay&""
				End If
				Sqlstr=Sqlstr&" Order By Fk_Article_Id Desc"
				'获取模板
				Rs2.Open Sqlstr,Conn,1,3
				If Not Rs2.Eof Then
					Rs2.PageSize=PageSizes
					If PageNow>Rs2.PageCount Or PageNow<=0 Then
						PageNow=1
					End If
					PageCounts=Rs2.PageCount
					Rs2.AbsolutePage=PageNow
					PageAll=Rs2.RecordCount
					i=1
					While (Not Rs2.Eof) And i<PageSizes+1
						Id=Rs2("Fk_Article_Id")
						Call FKHTML.CreatArticle(Rs2("Fk_Article_Template"),ModuleId,Rs2("Fk_Module_Dir"),Rs2("Fk_Article_FileName"),Rs2("Fk_Article_Title"),0)
						Rs2.MoveNext
						i=i+1
					Wend
				End If
				Rs2.Close
				If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=3&MenuId=<%=MenuId%>&ModuleId=<%=ModuleId%>&CreatType=<%=CreatType%>&CreatDay=<%=CreatDay%>&Page=<%=PageNow+1%>">
<%
					Call FKDB.DB_Close()
					Set Rs2=Nothing
					Response.End()
				Else
					Response.Write(ModuleName&"模块文章生成完毕！")
					Response.Flush()
					Response.Clear()
				End If
			End If
			If CreatType=0 Or CreatType=1 Then
				Id=ModuleId
				Call FKHTML.CreatArticleCategory(TemplateId2,ModuleDir,ModuleName)
			End If
		ElseIf ModuleType=2 Then
			If CreatType=0 Or CreatType=2 Then
				Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Menu="&MenuId&" And Fk_Module_Id=" & ModuleId
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And DateDiff('d',Fk_Product_Time,'"&Now()&"')<="&CreatDay&""
				End If
				Sqlstr=Sqlstr&" Order By Fk_Product_Id Desc"
				'获取模板
				Rs2.Open Sqlstr,Conn,1,3
				If Not Rs2.Eof Then
					Rs2.PageSize=PageSizes
					If PageNow>Rs2.PageCount Or PageNow<=0 Then
						PageNow=1
					End If
					PageCounts=Rs2.PageCount
					Rs2.AbsolutePage=PageNow
					PageAll=Rs2.RecordCount
					i=1
					While (Not Rs2.Eof) And i<PageSizes+1
						Id=Rs2("Fk_Product_Id")
						Call FKHTML.CreatProduct(Rs2("Fk_Product_Template"),ModuleId,Rs2("Fk_Module_Dir"),Rs2("Fk_Product_FileName"),Rs2("Fk_Product_Title"),0)
						Rs2.MoveNext
						i=i+1
					Wend
				End If
				Rs2.Close
				If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=3&MenuId=<%=MenuId%>&ModuleId=<%=ModuleId%>&CreatType=<%=CreatType%>&CreatDay=<%=CreatDay%>&Page=<%=PageNow+1%>">
<%
					Call FKDB.DB_Close()
					Set Rs2=Nothing
					Response.End()
				Else
					Response.Write(ModuleName&"产品模块生成完毕！")
					Response.Flush()
					Response.Clear()
				End If
			End If
			If CreatType=0 Or CreatType=1 Then
				Id=ModuleId
				Call FKHTML.CreatProductCategory(TemplateId2,ModuleDir,ModuleName)
			End If
		ElseIf ModuleType=3 Then
			Call FKHTML.CreatInfo(TemplateId,ModuleFileName,ModuleName,0)
		ElseIf ModuleType=4 Then
			Call FKHTML.CreatGBook(TemplateId,ModuleFileName,ModuleName)
		ElseIf ModuleType=5 Then
			Response.Write("转向链接无需生成！")
		ElseIf ModuleType=6 Then
			Call FKHTML.CreatJob(TemplateId,ModuleFileName,ModuleName)
		ElseIf ModuleType=7 Then
			If CreatType=0 Or CreatType=2 Then
				Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Menu="&MenuId&" And Fk_Module_Id=" & ModuleId
				If CreatDay>0 Then
					Sqlstr=Sqlstr&" And DateDiff('d',Fk_Down_Time,'"&Now()&"')<="&CreatDay&""
				End If
				Sqlstr=Sqlstr&" Order By Fk_Down_Id Desc"
				'获取模板
				Rs2.Open Sqlstr,Conn,1,3
				If Not Rs2.Eof Then
					Rs2.PageSize=PageSizes
					If PageNow>Rs2.PageCount Or PageNow<=0 Then
						PageNow=1
					End If
					PageCounts=Rs2.PageCount
					Rs2.AbsolutePage=PageNow
					PageAll=Rs2.RecordCount
					i=1
					While (Not Rs2.Eof) And i<PageSizes+1
						Id=Rs2("Fk_Down_Id")
						Call FKHTML.CreatDown(Rs2("Fk_Down_Template"),ModuleId,Rs2("Fk_Module_Dir"),Rs2("Fk_Down_FileName"),Rs2("Fk_Down_Title"),0)
						Rs2.MoveNext
						i=i+1
					Wend
				End If
				Rs2.Close
				If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=3&MenuId=<%=MenuId%>&ModuleId=<%=ModuleId%>&CreatType=<%=CreatType%>&CreatDay=<%=CreatDay%>&Page=<%=PageNow+1%>">
<%
					Call FKDB.DB_Close()
					Set Rs2=Nothing
					Response.End()
				Else
					Response.Write(ModuleName&"下载模块生成完毕！")
					Response.Flush()
					Response.Clear()
				End If
			End If
			If CreatType=0 Or CreatType=1 Then
				Id=ModuleId
				Call FKHTML.CreatDownCategory(TemplateId2,ModuleDir,ModuleName)
			End If
		End If
	Else
		Rs.Close
		Response.Write("要生成的模块不存在！")
	End If
	Set Rs2=Nothing
	Response.Write("生成完毕！")
End Sub

'==============================
'函 数 名：HTMLDo3()
'作    用：一键生成HTML
'参    数：
'==============================
Sub HTMLDo3()
	'定义变量
	Dim PageCode,PageCodes,TemplateId,MenuId,ModuleId,ModuleType,ModuleDir,ModuleFileName,ModuleName
	Dim Pages,Id2,Dos
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Id2=Clng(Request.QueryString("Id"))
	Dos=Request.QueryString("Dos")
	If Dos="" Then
		Dos=1
	Else
		Dos=Clng(Dos)
	End If
%>
<STYLE> 
* {
	margin:0;
	padding:0;
}
body {
	font-size:12px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 12px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
</STYLE>
<%
	PageNow=Request.QueryString("Page")
	If PageNow<>"" Then
		PageNow=Clng(PageNow)
	Else
		PageNow=1
	End If
	If Dos=1 Then
		If Session("ArticleCategory")="" Then
			Session("ArticleCategory")=","
		End If
		Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Show=1"
		If Id2>0 Then
			Sqlstr=Sqlstr&" And DateDiff('d',Fk_Article_Time,'"&Now()&"')<="&Id2&""
		End If
		Sqlstr=Sqlstr&" Order By Fk_Article_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				If Instr(Session("ArticleCategory"),","&Rs2("Fk_Article_Module")&",")=0 Then
					Session("ArticleCategory")=Session("ArticleCategory")&Rs2("Fk_Article_Module")&","
				End If
				Id=Rs2("Fk_Article_Id")
				Call FKHTML.CreatArticle(Rs2("Fk_Article_Template"),Rs2("Fk_Article_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Article_FileName"),Rs2("Fk_Article_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=<%=Dos%>&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write("文章部分生成完毕！")
			Response.Flush()
			Response.Clear()
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=2">
<%
		End If
	End If
	If Dos=2 Then
		If Session("ProductCategory")="" Then
			Session("ProductCategory")=","
		End If
		Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Show=1"
		If Id2>0 Then
			Sqlstr=Sqlstr&" And DateDiff('d',Fk_Product_Time,'"&Now()&"')<="&Id2&""
		End If
		Sqlstr=Sqlstr&" Order By Fk_Product_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				If Instr(Session("ProductCategory"),","&Rs2("Fk_Product_Module")&",")=0 Then
					Session("ProductCategory")=Session("ProductCategory")&Rs2("Fk_Product_Module")&","
				End If
				Id=Rs2("Fk_Product_Id")
				Call FKHTML.CreatProduct(Rs2("Fk_Product_Template"),Rs2("Fk_Product_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Product_FileName"),Rs2("Fk_Product_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=<%=Dos%>&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write(ModuleName&"产品模块生成完毕！")
			Response.Flush()
			Response.Clear()
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=3">
<%
		End If
	End If
	If Dos=3 Then
		If Session("DownCategory")="" Then
			Session("DownCategory")=","
		End If
		Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Show=1"
		If Id2>0 Then
			Sqlstr=Sqlstr&" And DateDiff('d',Fk_Down_Time,'"&Now()&"')<="&Id2&""
		End If
		Sqlstr=Sqlstr&" Order By Fk_Down_Id Desc"
		'获取模板
		Rs2.Open Sqlstr,Conn,1,3
		If Not Rs2.Eof Then
			Rs2.PageSize=PageSizes
			If PageNow>Rs2.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs2.PageCount
			Rs2.AbsolutePage=PageNow
			PageAll=Rs2.RecordCount
			i=1
			While (Not Rs2.Eof) And i<PageSizes+1
				If Instr(Session("DownCategory"),","&Rs2("Fk_Down_Module")&",")=0 Then
					Session("DownCategory")=Session("DownCategory")&Rs2("Fk_Down_Module")&","
				End If
				Id=Rs2("Fk_Down_Id")
				Call FKHTML.CreatDown(Rs2("Fk_Down_Template"),Rs2("Fk_Down_Module"),Rs2("Fk_Module_Dir"),Rs2("Fk_Down_FileName"),Rs2("Fk_Down_Title"),0)
				Rs2.MoveNext
				i=i+1
			Wend
		End If
		Rs2.Close
		If PageNow<PageCounts Then
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=<%=Dos%>&Page=<%=PageNow+1%>">
<%
			Call FKDB.DB_Close()
			Set Rs2=Nothing
			Response.End()
		Else
			Response.Write(ModuleName&"下载模块生成完毕！")
			Response.Flush()
			Response.Clear()
%>
<meta http-equiv="refresh" content="1;URL=HTML.asp?Type=4&Id=<%=Id2%>&Dos=4">
<%
		End If
	End If
	If Dos=4 Then
		If Len(Session("ArticleCategory"))>2 Then
			Session("ArticleCategory")=Left(Session("ArticleCategory"),Len(Session("ArticleCategory"))-1)
			Session("ArticleCategory")=Right(Session("ArticleCategory"),Len(Session("ArticleCategory"))-1)
		Else
			Session("ArticleCategory")=0
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=1 And Fk_Module_Id In ("&Session("ArticleCategory")&") Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatArticleCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
		If Len(Session("ProductCategory"))>2 Then
			Session("ProductCategory")=Left(Session("ProductCategory"),Len(Session("ProductCategory"))-1)
			Session("ProductCategory")=Right(Session("ProductCategory"),Len(Session("ProductCategory"))-1)
		Else
			Session("ProductCategory")=0
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=2 And Fk_Module_Id In ("&Session("ProductCategory")&") Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatProductCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
		If Len(Session("DownCategory"))>2 Then
			Session("DownCategory")=Left(Session("DownCategory"),Len(Session("DownCategory"))-1)
			Session("DownCategory")=Right(Session("DownCategory"),Len(Session("DownCategory"))-1)
		Else
			Session("DownCategory")=0
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=7 And Fk_Module_Id In ("&Session("DownCategory")&") Order By Fk_Module_Id Desc"
		Rs2.Open Sqlstr,Conn,1,3
		jj=PageSizes
		While Not Rs2.Eof
			Id=Rs2("Fk_Module_Id")
			PageArr=Split(Rs2("Fk_Module_PageCode"),"|--|")
			PageSizes=jj
			If Rs2("Fk_Module_PageCount")>0 Then
				PageSizes=Rs2("Fk_Module_PageCount")
			End If
			Call FKHTML.CreatDownCategory(Rs2("Fk_Module_Template"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_Name"))
			Rs2.MoveNext
		Wend
		Rs2.Close
		Session("ArticleCategory")=""
		Session("ProductCategory")=""
		Session("DownCategory")=""
		Call FKHTML.CreatIndex()
		Response.Write("一键生成更新完成！")
	End If
	Set Rs2=Nothing
End Sub
%>
<!--#Include File="../Code.asp"-->