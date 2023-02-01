<!--#Include File="../AdminCheck.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：Weixin_subscribeList.asp
'文件用途：微信关注管理拉取页面
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
Dim Fk_menuName,Fk_menuType,Fk_menuEvent,Fk_menuStatus,Fk_menuPx,Fk_menuParent

'获取参数
Types=Clng(Request.QueryString("Type"))
id=Clng(Request.QueryString("id"))

Select Case Types
	Case 1
		Call WeixinsubscribeList() '微信关注列表
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WeixinsubscribeList()
'作    用：微信关注列表
'参    数：
'==========================================
Sub WeixinsubscribeList()
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
<script type="text/javascript">
 function closeWin()
   {
    window.parent.ymPrompt.doHandler("error",true);
   }
</script>
<style type="text/css">
body{padding:4px;}
#ListTop,#ListContent{width:100%;margin:0 auto}
#ListContent table{border-right:0}
#ListContent table td{line-height:34px;}
</style>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false;">关注用户管理</a></li>
    </ul>
</div>
<div id="ListContent">
<form name="form">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">OpenID</td>
            <td align="center" class="ListTdTop">关注时间</td>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_subscribeList] "
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And openID Like '%%"&SearchStr&"%%'"
	End If
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
            <td height="20" align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("id")%>" id="List<%=Rs("id")%>" /></td>
            <td><%=Rs("openID")%></td>
            <td align="center"><%=Rs("subscribe_time")%></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="3" style="text-indent:24px;"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="SelectAll('ListId')"> 全选
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("/admin/weixin/weixin_subscribeList.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="3" align="center">暂无记录</td>
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