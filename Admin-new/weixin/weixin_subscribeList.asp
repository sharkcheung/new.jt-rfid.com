<!--#Include File="CheckToken.asp"-->
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

private function CheckFields(FieldsName,TableName)
	dim blnFlag,chkStrSql,chkStrRs
	blnFlag=False
	chkStrSql="select * from "&TableName
	Set chkStrRs=Conn.Execute(chkStrSql)
	for i = 0 to chkStrRs.Fields.Count - 1
		if lcase(chkStrRs.Fields(i).Name)=lcase(FieldsName) then
			blnFlag=True
			Exit For
		else
			blnFlag=False
		end if
	Next
	CheckFields=blnFlag
End Function
'==========================================
'函 数 名：WeixinsubscribeList()
'作    用：微信关注列表
'参    数：
'==========================================
Sub WeixinsubscribeList()
	if not CheckFields("wxnickname","weixin_subscribeList") then
		conn.execute("alter table weixin_subscribeList add column wxnickname varchar(100) null")
		conn.execute("alter table weixin_subscribeList add column wxsex int default 0")
		conn.execute("alter table weixin_subscribeList add column wxlanguage varchar(100) null")
		conn.execute("alter table weixin_subscribeList add column wxcity varchar(50) null")
		conn.execute("alter table weixin_subscribeList add column wxprovince varchar(50) null")
		conn.execute("alter table weixin_subscribeList add column wxcountry varchar(50) null")
		conn.execute("alter table weixin_subscribeList add column wxheadimgurl varchar(200) null")
		conn.execute("alter table weixin_subscribeList add column wxremark varchar(255) null")
	end if
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

<div class="pageright xgzs ssyqtj" style="border-top:0;">
	<div class="xgzsbtm" style="padding-top:0;">
		<form name="form">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="khdtj">
			<tr>
				<th class="ListTdTop" width="200">头像</th>
				<th class="ListTdTop">昵称</th>
				<th class="ListTdTop">性别</th>
				<th class="ListTdTop">国家</th>
				<th class="ListTdTop">省份</th>
				<th class="ListTdTop">城市</th>
				<th class="ListTdTop" width="220">关注时间</th>
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
				<td align="center"><img style="padding:5px 0;" src="<%=Rs("wxheadimgurl")%>" width="64" height="64"/></td>
				<td align="center"><%=Rs("wxnickname")%></td>
				<td align="center"><%if Rs("wxsex")=1 then 
				response.write "男"
				else
				response.write "女"
				end if%></td>
				<td align="center"><%=Rs("wxcountry")%></td>
				<td align="center"><%=Rs("wxprovince")%></td>
				<td align="center"><%=Rs("wxcity")%></td>
				<td align="center"><%=Rs("subscribe_time")%></td>
			</tr>
			<%
						Rs.MoveNext
						i=i+1
					Wend
					
			%>
					<tr>
						<td style="padding:5px 0;" colspan="7" align="center"><%Call FKFun.ShowPageCode("/admin/weixin/weixin_subscribeList.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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
</div>
<%
End Sub
%><!--#Include File="../../Code.asp"-->






















