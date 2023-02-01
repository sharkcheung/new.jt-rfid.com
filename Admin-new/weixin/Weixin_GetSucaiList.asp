<!--#Include File="../AdminCheck.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<link href="/admin/Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Js/function.js"></script>
</head>
<body>
<%
'判断权限
If Not FkFun.CheckLimit("System2") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_menuName,Fk_menuType,Fk_menuEvent,Fk_menuStatus,Fk_menuPx,Fk_menuParent

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinSucaiList() '微信素材列表
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WeixinSucaiList()
'作    用：微信素材列表
'参    数：
'==========================================
Sub WeixinSucaiList()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
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
%>

<style type="text/css">
body{padding:4px;}
#ListTop,#ListContent{width:100%;margin:0 auto}
#ListContent table{border-right:0}
#ListContent table td{line-height:34px;}
</style>
<div id="ListTop">
    模块&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','Down.asp?Type=1&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>&nbsp;&nbsp;请选择模块：
<select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);" style="vertical-align:middle;">
      <option value="alert('请选择模块');">请选择模块</option>
</select>
</div>
<div id="ListContent">
<form name="form">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">标题</td>
            <td align="center" class="ListTdTop">类型</td>
            <td align="center" class="ListTdTop">素材</td>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_Sucai] Where Sucai_type=0 and Sucai_status=0"
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Sucai_title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By Sucai_px Desc"
	Rs.Open Sqlstr,Conn,1,1
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
            <td height="20" align="center"><input type="radio" name="ListId" class="Checks" /><input type="hidden" id="picurl" value="<%=Rs("Sucai_file")%>"></td>
            <td><%=Rs("Sucai_title")%></td>
            <td align="center">图片</td>
            <td align="center"><img src="<%=rs("Sucai_file")%>" width="64" height="64"/></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="4" style="text-indent:24px;">
            <%Call FKFun.ShowPageCode("/admin/weixin_GetSucaiList.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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
	</form>
</div>
<%
End Sub
%><!--#Include File="../../Code.asp"-->