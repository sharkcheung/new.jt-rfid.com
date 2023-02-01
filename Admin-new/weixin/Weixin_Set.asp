<!--#Include File="CheckToken.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：Weixin_Set.asp
'文件用途：微信接口设置拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================



'定义页面变量
dim wx_token,wx_raw_id,wx_AppId,wx_AppSecret,wx_url,wx_Subscribe,wx_Repeat,wx_Random,wx_NoneReply

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinSet() '微信接口配置
	Case 2
		Call WeixinSetDo() '微信接口配置执行修改
	Case 3
		Call AddText() '微信接口配置回复文本
	Case 4
		Call AddTextDo() '微信接口配置执行修改回复文本
	Case 5
		Call AddNews() '微信接口配置回复图文
	Case 6
		Call AddNewsDo() '微信接口配置执行修改回复图文
	Case 7
		Call KeyAddText() '微信接口配置关键词未匹配回复文本
	Case 8
		Call KeyAddTextDo() '微信接口配置执行修改关键词未匹配回复文本
	Case 9
		Call KeyAddNews() '微信接口配置关键词未匹配回复图文
	Case 10
		Call KeyAddNewsDo() '微信接口配置执行修改关键词未匹配回复图文
	Case Else
		Response.Write("没有找到此功能项！")
End Select

Function randKey(obj) 
	 Dim char_array(80) 
	 Dim temp 
	 For i = 0 To 9  
	  char_array(i) = Cstr(i) 
	 Next 
	 For i = 10 To 35 
	  char_array(i) = Chr(i + 55) 
	 Next 
	 For i = 36 To 61 
	  char_array(i) = Chr(i + 61) 
	 Next 
	 Randomize 
	 For i = 1 To obj 
	  'rnd函数返回的随机数在0~1之间，可等于0，但不等于1 
	  '公式：int((上限-下限+1)*Rnd+下限)可取得从下限到上限之间的数，可等于下限但不可等于上限 
	  temp = temp&char_array(int(62 - 0 + 1)*Rnd + 0) 
	 Next 
	 randKey = temp 
End Function

'==============================
'函 数 名：KeyAddNewsDo
'作    用：微信接口配置执行修改关键词未匹配回复图文
'参    数：
'==============================
Sub KeyAddNewsDo()
	id = Trim(Request.Form("ListId"))
	wx_NoneReply= "[wx_news="""& id &"""]"
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_NoneReply")	= wx_NoneReply
	Rs.Update()
	Application.UnLock()
	Response.Write("修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：KeyAddTextDo
'作    用：微信接口配置关键词未匹配回复图文
'参    数：
'==============================
Sub KeyAddTextDo()
	wx_NoneReply= Trim(Request.Form("wx_NoneReply"))
	
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_NoneReply")	= wx_NoneReply
	Rs.Update()
	Application.UnLock()
	Response.Write("修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：AddNewsDo
'作    用：微信接口配置执行修改回复图文
'参    数：
'==============================
Sub AddNewsDo()
	id = Trim(Request.Form("ListId"))
	wx_Subscribe= "[wx_news="""& id &"""]"
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_Subscribe")	= wx_Subscribe
	Rs.Update()
	Application.UnLock()
	Response.Write("修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：AddTextDo
'作    用：微信接口配置执行修改回复文本
'参    数：
'==============================
Sub AddTextDo()
	wx_Subscribe= Trim(Request.Form("wx_Subscribe"))
	
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_Subscribe")	= wx_Subscribe
	Rs.Update()
	Application.UnLock()
	Response.Write("修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：WeixinSetDo
'作    用：执行微信接口配置修改
'参    数：
'==============================
Sub WeixinSetDo()
	wx_token	= FKFun.HTMLEncode(Trim(Request.Form("wx_token")))
	wx_raw_id	= FKFun.HTMLEncode(Trim(Request.Form("wx_raw_id")))
	wx_AppId	= FKFun.HTMLEncode(Trim(Request.Form("wx_AppId")))
	wx_AppSecret= FKFun.HTMLEncode(Trim(Request.Form("wx_AppSecret")))
	wx_url		= Trim(Request.Form("wx_url"))
	wx_Subscribe= Trim(Request.Form("wx_Subscribe"))
	wx_NoneReply= Trim(Request.Form("wx_NoneReply"))
	wx_Repeat	= Trim(Request.Form("wx_Repeat"))
	wx_Random	= Trim(Request.Form("wx_Random"))
	
	Call FKFun.ShowString(wx_raw_id,1,50,0,"微信原始账号为必填项","微信原始账号不能大于50个字符！")
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_token")		= wx_token
	Rs("wx_raw_id")		= wx_raw_id
	Rs("wx_AppId")		= wx_AppId
	Rs("wx_AppSecret")	= wx_AppSecret
	Rs("wx_url")		= wx_url
	Rs("wx_Subscribe")	= wx_Subscribe
	Rs("wx_NoneReply")	= wx_NoneReply
	Rs("wx_Repeat")		= wx_Repeat
	Rs("wx_Random")		= wx_Random
	Rs.Update()
	Application.UnLock()
	Response.Write("设置成功！")
	Rs.Close
End Sub

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
'函 数 名 WeixinSet()
'作    用 读取微信接口设置
'参    数
'==========================================
Sub WeixinSet()
if not CheckFields("wx_NoneReply","weixin_config") then
	conn.execute("alter table weixin_config add column wx_NoneReply varchar(200) null")
end if
set rs=conn.execute("select * from weixin_config")
if not rs.eof then
	wx_token	= trim(rs("wx_token")&" ")
	wx_raw_id	= trim(rs("wx_raw_id")&" ")
	wx_AppId	= trim(rs("wx_AppId")&" ")
	wx_AppSecret= trim(rs("wx_AppSecret")&" ")
	wx_Repeat	= trim(rs("wx_Repeat")&" ")
	wx_Random	= trim(rs("wx_Random")&" ")
	wx_Subscribe	= trim(rs("wx_Subscribe")&" ")
	wx_NoneReply	= trim(rs("wx_NoneReply")&" ")
end if
rs.close
set rs=nothing
if wx_Repeat="" then
	wx_Repeat= "0"
end if
if wx_Random="" then
	wx_Random= "0"
end if
if wx_token="" then
	wx_token=ucase(randKey(32))
end if
%>

<div class="pageright xgzs ssyqtj" style="border-top:0;">
	<div class="xgzstop">
		<div class="xgzstopsm">
			<strong>温馨提示：</strong>
			系统提供微信公众平台上自动回复的接口，请按以下步骤操作<br /> 
1. 请先登录 <a href='https://mp.weixin.qq.com/' target="_blank"><span style='color: #f06100;'>微信公众平台</span></a> 如果没有，请先<a href="https://mp.weixin.qq.com/cgi-bin/readtemplate?t=register/step1_tmpl&lang=zh_CN" target="_blank"><span style='color: #0071b7;'>开通公众账号</span></a><br /> 
2. 请设置好以下必填项(带"*"号为必填项)，并点击保存设置<br />
3. 进入[高级功能]，在关闭[编辑模式]后，开启并进入[开发模式]配置页面<br /> 
4. 把下列的接口配置信息 <span style='color: #f06100;'>URL</span> 和 <span style='color: #f06100;'>Token</span> 写入微信公众平台的配置页面后提交即可<br /> 
5. 如果您的公众号拥有高级接口权限，请将公众平台中的 <span style='color: #f06100;'>开发者凭据(AppId和AppSecret)</span> 设置在此页面 
		</div>
	</div>
	<div class="xgzsbtm" style="padding-top:20px;">
		<form id="SystemSet" name="SystemSet" method="post" action="Weixin_Set.asp?Type=2" onsubmit="return false;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="khdtj">
			<tr>
				<th colspan="3">微信接口设置</th>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop" width="120">URL</td>
				<td width="320"><input name="wx_url" type="text" class="Input" id="wx_url" value="http://<%=Request.ServerVariables("Http_Host")%>/admin/weixin/" size="50"  style="vertical-align:middle" readonly/></td>
				<td><span style='color: #f06100;'>*</span> 此项由系统自动生成</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">Token</td>
				<td><input name="wx_token" type="text" class="Input" id="wx_token" value="<%=wx_token%>" size="50"  style="vertical-align:middle" readonly/></td>
				<td><span style='color: #f06100;'>*</span> 此项由系统自动生成</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">微信原始账号</td>
				<td><input name="wx_raw_id" type="text" class="Input" id="wx_raw_id" value="<%=wx_raw_id%>" size="50"  style="vertical-align:middle"/></td>
				<td><span style='color: #f06100;'>*</span> 即微信后台账户信息中显示的原始ID</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">AppId</td>
				<td><input name="wx_AppId" type="text" class="Input" id="wx_AppId" value="<%=wx_AppId%>" size="50"  style="vertical-align:middle"/></td>
				<td>选填，请到微信后台->高级功能->开发模式的开发者凭据中获取</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">AppSecret</td>
				<td><input name="wx_AppSecret" type="text" class="Input" id="wx_AppSecret" size="50" value="<%=wx_AppSecret%>"/></td>
				<td>同上，配置AppId和AppSecret可以配置自定义菜单和调用相应的高级接口</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">微信开场白</td>
				<td><textarea style="margin-bottom:5px;" name="wx_Subscribe" id="wx_Subscribe" rows="4" cols="35" readonly><%=wx_Subscribe%></textarea> <br><input style="margin-left:10px;" type="button" value="回复文本" class="Button addText" onclick="ShowBox('/admin-new/weixin/weixin_Set.asp?type=3&id=0','设置回复文本','500px')"> <input type="button" value="回复图文" onclick="ShowBox('/admin-new/weixin/weixin_Set.asp?type=5&id=0','设置回复图文','700px')" class="Button"></td>
				<td>注：用户关注您的微信账号后将收到此内容，留空则不做处理</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">随机回复</td>
				<td style="padding-left:10px;"><input name="wx_Random" type="radio" id="wx_Random" value="1" style="vertical-align:middle;" <%if wx_Random="1" then response.write "checked"%>/><label for="wx_Random" style="vertical-align:middle;">开启</label> &nbsp;<input name="wx_Random" type="radio" style="vertical-align:middle;" id="wx_Random1" value="0" <%if wx_Random="0" then response.write "checked"%>/><label for="wx_Random1" style="vertical-align:middle;">关闭</label></td>
				<td>注：关闭则默认回复第一条内容，开启则随机挑选回复内容</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">关键词匹配回复</td>
				<td><textarea style="margin-bottom:5px;" name="wx_NoneReply" id="wx_NoneReply" rows="4" cols="35" readonly><%=wx_NoneReply%></textarea> <br><input style="margin-left:10px;" type="button" value="回复文本" class="Button" onclick="ShowBox('/admin-new/weixin/weixin_Set.asp?type=7&id=0','设置回复文本','500px')"> <input type="button" value="回复图文" class="Button" onclick="ShowBox('/admin-new/weixin/weixin_Set.asp?type=9&id=0','设置回复图文','700px')"></td>
				<td>注：用户关注您的微信账号后将收到此内容，留空则不做处理</td>
			</tr>
			<tr>
				<td height="25" align="right" class="MainTableTop">重复机制</td>
				<td><input name="wx_Repeat" type="text" class="Input" id="wx_Repeat" size="50" value="<%=wx_Repeat%>"/></td>
				<td> 注：0=允许重复提问，1=重复一定次数后不再处理</td>
			</tr>
			<tr class="last">
				<td style="border-bottom:none;background:none;">&nbsp;</td>
				<td colspan="2" style="padding-left:10px;border-bottom:none;background:none;"><input style="margin-right:10px;" type="submit" onclick="Sends('SystemSet','/admin/weixin/Weixin_Set.asp?Type=2',0,'',1,0,'','');" class="Button" name="button" id="button" value="保存设置"/></td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%
End Sub

'==========================================
'函 数 名：AddText()
'作    用：回复文本
'参    数：
'==========================================
Sub AddText()

set rs=conn.execute("select * from weixin_config")
if not rs.eof then
	wx_Subscribe	= trim(rs("wx_Subscribe")&" ")
	if instr(wx_Subscribe,"[wx_news")>0 then
		wx_Subscribe = ""
	end if
end if
rs.close
set rs=nothing
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/Weixin_Set.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td>文本内容<br><textarea name="wx_Subscribe" id="val" cols="60" style="width:400px;" rows="6"><%=wx_Subscribe%></textarea></td>
	    </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/Weixin_Set.asp?Type=4',0,'',0,1,'MainRight','/admin-new/weixin/Weixin_Set.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：AddNews()
'作    用：微信接口配置回复图文
'参    数：
'==========================================
Sub AddNews()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr,id
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	
	id=Clng(Request.QueryString("id"))

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

<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/Weixin_Set.asp?Type=6" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
            <td><%=Rs("imgText_Title")%><input type="hidden" name="wx_NoneReply" id="news_<%=Rs("id")%>" value="<%=Rs("imgText_Title")%>"></td>
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
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/Weixin_Set.asp?Type=6',0,'',0,1,'MainRight','/admin-new/weixin/Weixin_Set.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：KeyAddText()
'作    用：关键词未匹配回复文本
'参    数：
'==========================================
Sub KeyAddText()

set rs=conn.execute("select wx_NoneReply from weixin_config")
if not rs.eof then
	wx_NoneReply	= trim(rs("wx_NoneReply")&" ")
	if instr(wx_NoneReply,"[wx_news")>0 then
		wx_NoneReply = ""
	end if
end if
rs.close
set rs=nothing
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/Weixin_Set.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td>文本内容<br><textarea name="wx_NoneReply" id="val" style="width:400px;" cols="60" rows="6"><%=wx_NoneReply%></textarea></td>
	    </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/Weixin_Set.asp?Type=8',0,'',0,1,'MainRight','/admin-new/weixin/Weixin_Set.asp?Type=1');" class="Button" name="button" id="button" value="确 定" />
        <input type="button" onclick="layer.closeAll();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：KeyAddNews()
'作    用：微信接口配置关键词未匹配回复图文
'参    数：
'==========================================
Sub KeyAddNews()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr,id
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	
	id=Clng(Request.QueryString("id"))

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

<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/Weixin_Set.asp?Type=10" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
            <td><%=Rs("imgText_Title")%></td>
            <td><%=Rs("imgText_Title")%></td>
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
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/Weixin_Set.asp?Type=10',0,'',0,1,'MainRight','/admin-new/weixin/Weixin_Set.asp?Type=1');" class="Button" name="button" id="button" value="确 定" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub
%><!--#Include File="../../Code.asp"-->






















