<!--#Include File="../AdminCheck.asp"-->
<link href="/admin/Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Js/function.js"></script>
<%
'==========================================
'文 件 名：Weixin_Menu.asp
'文件用途：微信图文管理拉取页面
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

Call WeixinImgTextList() '微信图文列表

'==========================================
'函 数 名：WeixinImgTextList()
'作    用：微信图文列表
'参    数：
'==========================================
Sub WeixinImgTextList()
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
<div id="ListContent">
<form name="form">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">标题</td>
            <td align="center" class="ListTdTop">状态</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_imageText] Where id<>"&id
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And imgText_Title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By imgText_px Desc"
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
            <td height="20" align="center"><input type="<%if Types=1 then%>checkbox<%else%>radio<%end if%>" name="ListId" class="Checks" value="<%=Rs("id")%>" id="List<%=Rs("id")%>" /></td>
            <td><%=Rs("imgText_Title")%><input type="hidden" id="news_<%=Rs("id")%>" value="<%=Rs("imgText_Title")%>"></td>
            <td align="center"><%if Rs("imgText_status")=0 then:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_1.gif' title='启用'>":else:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif' title='禁用'>":end if%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin_GetNewsList.asp?Type=4&Id=<%=Rs("id")%>');">预览</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="4" style="text-indent:24px;"><%if Types=1 then%><input name="chkall" type="checkbox" id="chkall" value="select" onClick="SelectAll('ListId')"> 全选<%end if%>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("/admin/weixin_GetNewsList.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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