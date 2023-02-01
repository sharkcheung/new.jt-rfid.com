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

Select Case Types
	Case 1
		Call WeixinImgTextList() '微信图文列表
	Case 2
		Call WeixinImgTextAdd() '添加微信图文
	Case 3
		Call WeixinImgTextAddDo() '添加微信图文
	Case 4
		Call WeixinImgTextEditForm() '修改微信菜单
	Case 5
		Call WeixinImgTextEditDo() '执行修改微信菜单
	Case 6
		Call WeixinImgTextDelDo() '执行删除微信菜单
	Case 7
		Call WeixinImgTextMake() '生成微信菜单
	Case 8
		Call WeixinImgTextMakeDo() '执行生成微信菜单
	Case Else
		Response.Write("没有找到此功能项！")
End Select

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
            <td align="center" class="ListTdTop">大小</td>
            <td align="center" class="ListTdTop">素材</td>
            <td align="center" class="ListTdTop">状态</td>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_Sucai] where Sucai_type=1 "
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
            <td height="20" align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("id")%>" id="List<%=Rs("id")%>" /></td>
            <td><%=Rs("Sucai_title")%><input type="hidden" id="news_<%=Rs("id")%>" value="<%=Rs("Sucai_title")%>"></td>
            <td align="center"><%if trim(Rs("Sucai_fileSize")&" ")="" then 
			response.write "-"
			else
			response.write Rs("Sucai_fileSize")
			end if%></td>
            <td align="center"><embed flashvars="mp3=<%=Rs("Sucai_file")%>&autoplay=0" height="20" src="http://image001.dgcloud01.qebang.cn/website/weixin/music_player.swf" type="application/x-shockwave-flash" width="160" wmode="transparent"></td>
            <td align="center"><%if Rs("Sucai_px")=0 then:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_1.gif' title='启用'>":else:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif' title='禁用'>":end if%></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="4" style="text-indent:24px;"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="SelectAll('ListId')"> 全选
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("/admin/weixin_GetMusicList.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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

'==========================================
'函 数 名：WeixinImgTextAdd()
'作    用：添加微信图文
'参    数：
'==========================================
Sub WeixinImgTextAdd()
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加菜单</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">菜单名称：</td>
	        <td>&nbsp;<input name="Fk_menuName" type="text" class="Input" id="Fk_menuName" size="40" /></td>
	        <td>顶级菜单不超过8个字节(4个汉字);子菜单不超过14个字节(个7汉字)</td>
	    </tr>
	    <tr>
	        <td height="25" align="right">触发问题：</td>
	        <td>&nbsp;<input name="Fk_menuEvent" type="text" class="Input" id="Fk_menuEvent" size="40" /></td>
			<td>问题：深圳天气，URL: http://www.qebang.cn</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">上级菜单：</td>
	        <td>&nbsp;<select name="Fk_menuParent" class="Input" id="Fk_menuParent" style="vertical-align:middle;">
            <option value="0">-顶级分类-</option>
<%
	Sqlstr="Select * From [weixin_menu] Where menuParent=0 order by menuPx desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("id")%>"><%=Rs("menuName")%></option>
    <%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select></td>
			<td>顶级菜单下可以创建2~5个子菜单</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">事件类型：</td>
	        <td>&nbsp;<select name="Fk_menuType" class="Input" id="Fk_menuType" style="vertical-align:middle;">
            <option value="click">点击事件</option>
            <option value="view">访问网页</option>
            </select></td>
			<td>点击事件：触发问题，访问网页：跳转到指定url</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">排 序 值：</td>
	        <td>&nbsp;<input name="Fk_menuPx" class="Input" type="text" id="Fk_menuPx" value="0"></td>
			<td>值越大排列越靠前</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">菜单状态：</td>
	        <td>&nbsp;<input name="Fk_menuStatus" class="Input" type="radio" id="Fk_menuStatus" value="0" checked="checked" />启用
            <input type="radio" name="Fk_menuStatus" class="Input" id="Fk_menuStatus1" value="1" />禁用</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_menu.asp?Type=3',0,'',0,1,'MainRight','/admin/weixin/Weixin_menu.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinImgTextAddDo
'作    用：执行添加微信菜单
'参    数：
'==============================
Sub WeixinImgTextAddDo()
	Fk_menuName		= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuName")))
	Fk_menuEvent	= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuEvent")))
	Fk_menuParent	= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuParent")))
	Fk_menuType		= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuType")))
	Fk_menuPx		= Trim(Request.Form("Fk_menuPx"))
	Fk_menuStatus	= Trim(Request.Form("Fk_menuStatus"))
	if Fk_menuParent="0" then
		Call FKFun.ShowString(Fk_menuName,1,8,0,"请输入菜单名称！","顶级菜单名称不能大于8个字节(4个汉字)！")
	else
		Call FKFun.ShowString(Fk_menuName,1,14,0,"请输入菜单名称！","子菜单名称不能大于14个字节(7个汉字)！")
	end if
	Call FKFun.ShowString(Fk_menuEvent,1,255,0,"请输入触发问题！","触发问题不能大于255个字符！")
	Sqlstr="Select * From [weixin_menu]"
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs.AddNew()
		Rs("menuName")=Fk_menuName
		Rs("menuOnEvent")=Fk_menuEvent
		Rs("menuParent")=Fk_menuParent
		Rs("menuType")=Fk_menuType
		Rs("menuPx")=Fk_menuPx
		Rs("menuStatus")=Fk_menuStatus
		Rs.Update()
		Application.UnLock()
		Response.Write("菜单添加成功！")
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinImgTextEditForm
'作    用：修改微信菜单
'参    数：
'==========================================
Sub WeixinImgTextEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_menu] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_menuName		= Rs("menuName")
		Fk_menuEvent	= Rs("menuOnEvent")
		Fk_menuParent	= Rs("menuParent")
		Fk_menuType		= Rs("menuType")
		Fk_menuPx		= Rs("menuPx")
		Fk_menuStatus	= Rs("menuStatus")
	End If
	Rs.Close
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin_menu.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改菜单</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">菜单名称：</td>
	        <td>&nbsp;<input name="Fk_menuName" type="text" class="Input" id="Fk_menuName" size="40" value="<%=Fk_menuName%>"/><input type="hidden" value="<%=id%>" id="id" name="id"></td>
	        <td>顶级菜单不超过8个字节(4个汉字);子菜单不超过14个字节(个7汉字)</td>
	    </tr>
	    <tr>
	        <td height="25" align="right">触发问题：</td>
	        <td>&nbsp;<input name="Fk_menuEvent" type="text" class="Input" id="Fk_menuEvent" size="40" value="<%=Fk_menuEvent%>"/></td>
			<td>问题：深圳天气，URL: http://www.qebang.cn</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">上级菜单：</td>
	        <td>&nbsp;<select name="Fk_menuParent" class="Input" id="Fk_menuParent" style="vertical-align:middle;">
            <option value="0">-顶级分类-</option>
<%
	Sqlstr="Select * From [weixin_menu] Where menuParent=0 order by menuPx desc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("id")%>" <%if trim(Fk_menuParent)=trim(Rs("id")) then response.write "selected"%>><%=Rs("menuName")%></option>
    <%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select></td>
			<td>顶级菜单下可以创建2~5个子菜单</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">事件类型：</td>
	        <td>&nbsp;<select name="Fk_menuType" class="Input" id="Fk_menuType" style="vertical-align:middle;">
            <option value="click" <%if Fk_menuType="click" then response.write "selected"%>>点击事件</option>
            <option value="view" <%if Fk_menuType="view" then response.write "selected"%>>访问网页</option>
            </select></td>
			<td>点击事件：触发问题，访问网页：跳转到指定url</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">排 序 值：</td>
	        <td>&nbsp;<input name="Fk_menuPx" class="Input" type="text" id="Fk_menuPx" value="<%=Fk_menuPx%>" ></td>
			<td>值越大排列越靠前</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">菜单状态：</td>
	        <td>&nbsp;<input name="Fk_menuStatus" class="Input" type="radio" id="Fk_menuStatus" value="0" <%if Fk_menuStatus=0 then response.write "checked"%>/>启用
            <input type="radio" name="Fk_menuStatus" class="Input" id="Fk_menuStatus1" value="1"  <%if Fk_menuStatus=1 then response.write "checked"%>/>禁用</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_menu.asp?Type=5',0,'',0,1,'MainRight','/admin/weixin/Weixin_menu.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：WeixinImgTextMake()
'作    用：生成微信图文
'参    数：
'==========================================
Sub WeixinImgTextMake()
dim wx_AppId,wx_AppSecret
set rs=conn.execute("select wx_AppId,wx_AppSecret from weixin_config")
if not rs.eof then
	wx_AppId	 = rs("wx_AppId")
	wx_AppSecret = rs("wx_AppSecret")
end if
rs.close
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin_menu.asp?Type=8" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>生成图文</span></div>
<div id="BoxContents" style="width:98%;">
	<div style="margin:20px;padding:10px;word-wrap:break-word;word-break:break-all;border: 1px solid #ffbe7a;background: #fffced;">
	   请首先确保已创建图文<br>
	   请到公众号官方后台->高级功能->开发模式 中获取以下信息
	   </div>
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">AppId</td>
	        <td>&nbsp;<input name="Fk_wx_AppId" type="text" class="Input" id="Fk_wx_AppId" size="40" value="<%=wx_AppId%>"/></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">AppSecret</td>
	        <td>&nbsp;<input name="Fk_wx_AppSecret" class="Input" type="text" id="Fk_wx_AppSecret" value="<%=wx_AppSecret%>" size="40"></td>
	    </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_menu.asp?Type=8',0,'',0,1,'MainRight','/admin/weixin/Weixin_menu.asp?Type=1');" class="Button" name="button" id="button" value="生 成" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub


'==============================
'函 数 名：WeixinImgTextMakeDo
'作    用：执行生成菜单
'参    数：
'==============================
Sub WeixinImgTextMakeDo()
	dim wx_AppId,wx_AppSecret
	wx_AppId	= FKFun.HTMLEncode(Trim(Request.Form("Fk_wx_AppId")))
	wx_AppSecret= FKFun.HTMLEncode(Trim(Request.Form("Fk_wx_AppSecret")))
	Call FKFun.ShowString(wx_AppId,1,50,0,"请输入AppId！","AppId不能大于8个字节(4个汉字)！")
	Call FKFun.ShowString(wx_AppSecret,1,50,0,"请输入AppSecret！","AppSecret不能大于50个字符！")
	' Sqlstr="Select * From [weixin_config]"
	' Rs.Open Sqlstr,Conn,1,3
	' Application.Lock()
	' if rs.eof then
		' rs.addnew()
	' end if
	' Rs("wx_AppId")=wx_AppId
	' Rs("wx_AppSecret")=wx_AppSecret
	' Rs.Update()
	' Application.UnLock()
	' Rs.Close
	dim jsonHtml,subrs,i,j
	jsonHtml="{"&vbcrlf
	set rs=conn.execute("select * from weixin_menu where menuParent=0")
	if not rs.eof then
		jsonHtml=jsonHtml&" ""button"":["&vbcrlf
		i=0
		do while not rs.eof
			if i<>0 then jsonHtml=jsonHtml&","&vbcrlf
			set subrs=conn.execute("select * from weixin_menu where menuParent="&rs("id"))
			if not subrs.eof then	'存在子菜单
				jsonHtml=jsonHtml&"{"&vbcrlf
				jsonHtml=jsonHtml&"""name"":"""&rs("menuName")&""","&vbcrlf
				jsonHtml=jsonHtml&"""sub_button"":["&vbcrlf
				j=0
				do while not subrs.eof
					if j<>0 then jsonHtml=jsonHtml&","&vbcrlf
					jsonHtml=jsonHtml&"{"&vbcrlf
					jsonHtml=jsonHtml&"""type"":"""&subrs("menuType")&""","&vbcrlf
					jsonHtml=jsonHtml&"""name"":"""&subrs("menuName")&""","&vbcrlf
					if subrs("menuType")="view" then
						jsonHtml=jsonHtml&"""url"":"""&subrs("menuOnEvent")&""""&vbcrlf
					else
						jsonHtml=jsonHtml&"""key"":"""&subrs("menuOnEvent")&""""&vbcrlf
					end if
					jsonHtml=jsonHtml&"}"&vbcrlf
					
					j=j+1
				subrs.movenext
				loop
				jsonHtml=jsonHtml&"]"&vbcrlf
				jsonHtml=jsonHtml&"}"&vbcrlf
			else
			
				jsonHtml=jsonHtml&"{"&vbcrlf
				jsonHtml=jsonHtml&"""type"":"""&rs("menuType")&""","&vbcrlf
				jsonHtml=jsonHtml&"""name"":"""&rs("menuName")&""","&vbcrlf
				if rs("menuType")="view" then
					jsonHtml=jsonHtml&"""url"":"""&rs("menuOnEvent")&""""&vbcrlf
				else
					jsonHtml=jsonHtml&"""key"":"""&rs("menuOnEvent")&""""&vbcrlf
				end if
				jsonHtml=jsonHtml&"}"&vbcrlf
				
			end if
			subrs.close
			i=i+1
		rs.movenext
		loop
		jsonHtml=jsonHtml&"]"&vbcrlf
		jsonHtml=jsonHtml&"}"&vbcrlf
		dim access_token,returnMsg
		access_token=DoGet("https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_AppId&"&secret="&wx_AppSecret)
		access_token=strCut(access_token,"access_token"":""","""",2)
		returnMsg=DoPost("https://api.weixin.qq.com/cgi-bin/menu/create?access_token="&access_token,jsonHtml)
		if returnMsg="{""errcode"":0,""errmsg"":""ok""}" then
			Response.Write("菜单生成成功！请重启微信查看菜单效果")
		else
			Response.Write("菜单生成失败！请重试")
		end if
	else
		response.write "还未创建图文，请先创建好菜单再生成"
	end if
	rs.close

End Sub

Function ByteToStr(vIn)
	Dim strReturn,i,ThisCharCode,innerCode,Hight8,Low8,NextCharCode
	strReturn = "" 
	For i = 1 To LenB(vIn)
	ThisCharCode = AscB(MidB(vIn,i,1))
	If ThisCharCode < &H80 Then
	strReturn = strReturn & Chr(ThisCharCode)
	Else
	NextCharCode = AscB(MidB(vIn,i+1,1))
	strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
	i = i + 1
	End If
	Next
	ByteToStr = strReturn 
End Function

Function DoGet(url)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	'.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send()
	DoGet = .ResponseText
	End With
	Set Http = Nothing
	'DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoGet=""
	end if
End Function

Function DoPost(url,PostStr)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send(PostStr)
	DoPost = .ResponseBody
	End With
	Set Http = Nothing
	DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoPost=""
	end if
End Function

	'写入文件法调试
	public Function WriteFile(content)
		dim filepath,fso,fopen
		filepath=server.mappath(".")&"\wx.txt"
		Set fso = Server.CreateObject("scripting.FileSystemObject")
		set fopen=fso.OpenTextFile(filepath, 8 ,true)
		content = content&vbcrlf&"************line seperate("&now()&")*****************"
		fopen.writeline(content)
		set fso=nothing
		set fopen=Nothing
	End Function

Function strCut(strContent,StartStr,EndStr,CutType)
	Dim strHtml,S1,S2
	strHtml = strContent
	On Error Resume Next
	Select Case CutType
	Case 1
		S1 = InStr(strHtml,StartStr)
		S2 = InStr(S1,strHtml,EndStr)+Len(EndStr)
	Case 2
		S1 = InStr(strHtml,StartStr)+Len(StartStr)
		S2 = InStr(S1,strHtml,EndStr)
	End Select
	If Err Then
		strCute = ""
		Err.Clear
		Exit Function
	Else
		strCut = Mid(strHtml,S1,S2-S1)
	End If
End Function

'==============================
'函 数 名：WeixinImgTextEditDo
'作    用：执行修改菜单
'参    数：
'==============================
Sub WeixinImgTextEditDo()
	Id=Trim(Request.Form("Id"))
	Fk_menuName		= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuName")))
	Fk_menuEvent	= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuEvent")))
	Fk_menuParent	= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuParent")))
	Fk_menuType		= FKFun.HTMLEncode(Trim(Request.Form("Fk_menuType")))
	Fk_menuPx		= Trim(Request.Form("Fk_menuPx"))
	Fk_menuStatus	= Trim(Request.Form("Fk_menuStatus"))
	if Fk_menuParent="0" then
		Call FKFun.ShowString(Fk_menuName,1,8,0,"请输入菜单名称！","顶级菜单名称不能大于8个字节(4个汉字)！")
	else
		Call FKFun.ShowString(Fk_menuName,1,14,0,"请输入菜单名称！","子菜单名称不能大于14个字节(7个汉字)！")
	end if
	Call FKFun.ShowString(Fk_menuEvent,1,255,0,"请输入触发问题！","触发问题不能大于255个字符！")
	Sqlstr="Select * From [weixin_menu] where id="&id
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs("menuName")=Fk_menuName
		Rs("menuOnEvent")=Fk_menuEvent
		Rs("menuParent")=Fk_menuParent
		Rs("menuType")=Fk_menuType
		Rs("menuPx")=Fk_menuPx
		Rs("menuStatus")=Fk_menuStatus
		Rs.Update()
		Application.UnLock()
		Response.Write("菜单修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：WeixinImgTextDelDo
'作    用：执行删除微信菜单
'参    数：
'==============================
Sub WeixinImgTextDelDo()
	Id=Trim(Request("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [weixin_menu] Where id in("& Id &")"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("微信菜单删除成功！")
	Else
		Response.Write("微信菜单不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../../Code.asp"-->