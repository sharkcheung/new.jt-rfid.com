<!--#Include File="../AdminCheck.asp"-->
<!--#Include File="../../Class/Cls_HTML.asp"-->
<%
'==========================================
'文 件 名：Weixin_GetArticle.asp
'文件用途：内容管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Article_Title,Fk_Article_Content,Fk_Article_Click,Fk_Article_Show,Fk_Article_Time,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_Template,Fk_Article_FileName,Fk_Article_Subject,Fk_Article_Recommend,Fk_Article_Keyword,Fk_Article_Description,Fk_Article_From,Fk_Article_Color,Fk_Article_Url,Fk_Article_Field,Fk_Article_onTop,Fk_Article_px
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Article_Module
Dim Temp2,KeyWordlist,kwdrs,ki

'===================================== 
'过滤字符 
'===================================== 
Function FilterText(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
FilterText=t0 
End Function 

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ArticleList() '内容列表
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：ArticleList()
'作    用：内容列表
'参    数：
'==========================================
Sub ArticleList()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr,Fk_ModuleType,Fk_Module_FileName,Fk_Module_Show,Fk_Module_Click,Fk_Module_Order,Fk_Module_Time,Fk_Module_Dir
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	if Trim(Request.QueryString("ModuleId"))="" then
		Sqlstr="Select top 1 * From [Fk_Module]"
	else
		Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	end if
	if Trim(Request.QueryString("ModuleType"))="" then
		Fk_ModuleType=1
	else
		Fk_ModuleType=Trim(Request.QueryString("ModuleType"))
	end if
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Id=Rs("Fk_Module_Id")
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_FileName=Rs("Fk_Module_FileName")
		Fk_Module_Show=Rs("Fk_Module_Show")
		Fk_Module_Click=Rs("Fk_Module_Click")
		Fk_Module_Order=Rs("Fk_Module_Order")
		Fk_Module_Time=Rs("Fk_Module_Time")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
	Else
		PageErr=1
	End If
	Rs.Close
	'response.write Fk_Module_Id&"-"&Fk_ModuleType&"-"&Sqlstr
	'response.end
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<link href="../Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../../Js/function.js"></script>
</head>
<body>
<div id="ListTop" style="text-indent:0;width:100%"><%=Fk_Module_Name%>栏目&nbsp;<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;" size=10/>&nbsp;<input type="button" class="Button" onclick="window.location.href='Weixin_GetArticle.asp?Type=1&ModuleId=<%=Fk_Module_Id%>&ModuleType=<%=Fk_ModuleType%>&SearchStr='+escape(document.all.SearchStr.value);" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>&nbsp;请选择栏目：
<select name="D1" id="D1" style="vertical-align:middle;" onChange="eval(this.options[this.selectedIndex].value);">
      <option value="alert('请选择栏目');">请选择栏目</option>
<%
Call ModuleSelectUrl(Fk_Module_Menu,0,Fk_Module_Id)
%>
</select>
</div>
<div id="ListContent" style="width:100%">
    <form name="DelList" id="DelList" method="post" action="Weixin_GetArticle.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="left" class="ListTdTop">标题</td>
            <td align="left" class="ListTdTop">显示</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">添加时间</td>
        </tr>
<%dim toUrl
					if instr(SiteUrl,"http://")>0 then
						toUrl=SiteUrl&"/wap"
					else
						toUrl="http://"&SiteUrl&"/wap"
					end if
	if Fk_ModuleType=3 or Fk_ModuleType=4 then%>
		<tr>
            <td height="20" align="center"><input type="radio" name="ListId" class="Checks" /><input type="hidden" class="hid" value="<%=toUrl&FKTemplate.GetGoUrl(Fk_ModuleType,Fk_Module_Id,Fk_Module_Dir,Fk_Module_FileName)%>"></td>
            <td align="left" class="td1">&nbsp;&nbsp;<%=Fk_Module_Name%></td>
            <td align="left"><%If Fk_Module_Show=1 Then%><img src="../images/caidan1.png" style="vertical-align:middle;"/><%Else%><img src="../images/caidan0.png" style="vertical-align:middle;"/><%End If%></td>
            
            <td height="20" align="center"><%=Fk_Module_Order%></td>
            <td align="center"><%=Fk_Module_Time%></td>
        </tr>
		
	<%else
		Dim Rs2,urlto
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		if Fk_ModuleType=1 then
			Sqlstr="Select Fk_Article_Title,Fk_Article_Color,Fk_Article_Url,Fk_Article_Show,Fk_Article_Pic,Fk_Article_Click,px,Fk_Article_Time,Fk_Article_FileName,Fk_Article_Id  From [Fk_Article] Where Fk_Article_Module="&Fk_Module_Id&""
			If SearchStr<>"" Then
				Sqlstr=Sqlstr&" And Fk_Article_Title Like '%%"&SearchStr&"%%'"
			End If
			Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc,Px desc,Fk_Article_Time Desc"
		elseif Fk_ModuleType=2 then
			Sqlstr="Select Fk_Product_Title,Fk_Product_Color,Fk_Product_Url,Fk_Product_Show,Fk_Product_Pic,Fk_Product_Click,px,Fk_Product_Time,Fk_Product_FileName,Fk_Product_Id  From [Fk_Product] Where Fk_Product_Module="&Fk_Module_Id&""
			If SearchStr<>"" Then
				Sqlstr=Sqlstr&" And Fk_Product_Title Like '%%"&SearchStr&"%%'"
			End If
			Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc,Px desc,Fk_Product_Time Desc"
		else
			Sqlstr="Select Fk_Down_Title,Fk_Down_Color,Fk_Down_Url,Fk_Down_Show,Fk_Down_Pic,Fk_Down_Click,px,Fk_Down_Time,Fk_Down_FileName ,Fk_Down_Id From [Fk_Down] Where Fk_Down_Module="&Fk_Module_Id&""
			If SearchStr<>"" Then
				Sqlstr=Sqlstr&" And Fk_Down_Title Like '%%"&SearchStr&"%%'"
			End If
			Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc,Px desc,Fk_Down_Time Desc"
		end if
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			Dim ArticleTemplate
			Rs.PageSize=PageSizes
			If PageNow>Rs.PageCount Or PageNow<=0 Then
				PageNow=1
			End If
			PageCounts=Rs.PageCount
			Rs.AbsolutePage=PageNow
			PageAll=Rs.RecordCount
			i=1
			While (Not Rs.Eof) And i<PageSizes+1
			
			
					If Rs(2)<>"" Then
						urlto=Rs(2)
					Else
						If Fk_Module_Dir<>"" Then
							urlto=Fk_Module_Dir&"/"
						Else
							if Fk_ModuleType=1 then
								urlto="Article"&Fk_Module_Id&"/"
							elseif Fk_ModuleType=2 then
								urlto="Product"&Fk_Module_Id&"/"
							else
								urlto="Down"&Fk_Module_Id&"/"
							end if
						End If
						If Rs(8)<>"" Then
							urlto=urlto&Rs(8)&".html"
						Else
							urlto=urlto&Rs(9)&".html"
						End If
						If SiteHtml=1 Then
							urlto="/html"&SiteDir&urlto&""
						Else
							urlto=SiteDir&sTemp&"?"&urlto
						End If
					End If
					urlto=toUrl&urlto
%>
        <tr>
            <td height="20" align="center"><input type="radio" name="ListId" class="Checks" /><input type="hidden" class="hid" value="<%=urlto%>"></td>
            <td align="left" class="td1">&nbsp;&nbsp;<%=Rs(0)%><%If Rs(1)<>"" Then%><span style="color:<%=Rs(1)%>">■</span><%End If%><%If Rs(2)<>"" Then%>[转向链接]<%End If%></td>
            <td align="left"><%If Rs(3)=1 Then%><img src="../images/caidan1.png" style="vertical-align:middle;"/><%Else%><img src="../images/caidan0.png" style="vertical-align:middle;"/><%End If%><%If Rs(4)<>"" Then%><span style="color:#F00">[图]</span><%End If%><a style="display:none;" href="javascript:void(0);" title="<%=Fk_Article_Template%> ">[模]</a></td>
            <td height="20" align="center"><%=Rs(6)%></td>
            <td align="center"><%=Rs(7)%></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="5" align="center"><%Call ShowPageCode("Weixin_GetArticle.asp?Type=1&ModuleId="&Fk_Module_Id&"&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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
end if
%>
    </table>
    </form>
</div>
<div id="ListBottom">

</div>
</body>
</html>
<%
End Sub
	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCode(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href="""&PageUrl&"1"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href="""&PageUrl&(Nows-1)&""">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href="""&PageUrl&(Nows+1)&""">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href="""&PageUrl&AllPage&""">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""window.location.href='"&PageUrl&"'+this.options[this.selectedIndex].value;"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function
'==============================
'函 数 名：ModuleSelectUrl
'作    用：输出ModuleSelectURL列表
'参    数：要输出的菜单MenuIds，要输出级数LevelId，默认选择AutoId
'==============================
Public Function ModuleSelectUrl(MenuIds,LevelId,AutoId)
	Call ModuleSelectUrlM(MenuIds,LevelId,"",AutoId)
End Function
Public Function ModuleSelectUrlM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,1
	While Not Rs2.Eof
		If FkFun.CheckLimit("Module"&Rs2("Fk_Module_Id")) Then
	%>
					<option value="<%=GetNavGo(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"))%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
	<%
			If LevelId=0 Then
				TitleBacks="&nbsp;&nbsp;&nbsp;├"
			Else
				TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
			End If
			Call ModuleSelectUrlM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		End If
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：ModuleSelectId
'作    用：输出ModuleSelectId列表
'参    数：要输出的菜单MenuIds，要输出级数LevelId，默认选择AutoId
'==============================
Public Function ModuleSelectId(MenuIds,LevelId,AutoId)
	Call ModuleSelectIdM(MenuIds,LevelId,"",AutoId)
End Function
Public Function ModuleSelectIdM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,1
	While Not Rs2.Eof
		'If FkFun.CheckLimit("Module"&Rs2("Fk_Module_Id")) Then
	%>
					<option value="<%=Rs2("Fk_Module_Id")%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
	<%
			If LevelId=0 Then
				TitleBacks="&nbsp;&nbsp;&nbsp;├"
			Else
				TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
			End If
			Call ModuleSelectIdM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		'End If
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：GetNavGo
'作    用：输出分类级数参数
'参    数：要输出的模块ModuleLevelId
'==============================
Function GetNavGo(GetModuleType,GetModuleId)
	Select Case GetModuleType
		Case 0
			GetNavGo="alert('静态模块无需内容修改，如有修改直接改模板！');"
		Case 5
			GetNavGo="alert('转向链接无需内容修改！');"
		Case else
			GetNavGo="window.location.href='Weixin_GetArticle.asp?Type=1&ModuleId="&GetModuleId&"&ModuleType="&GetModuleType&"'"
	End Select
End Function
%><!--#Include File="../../Code.asp"-->