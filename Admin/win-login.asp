<!--#Include File="Include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<!--#Include File="nocache.asp"-->
<%
'==========================================
'用途：WIN系统端登录验证
'==========================================

'定义页面变量
Dim Fk_Admin_LoginName,Fk_Admin_LoginPass

Call LoginDo() '登录操作


'==========================================
'函 数 名：LoginDo()
'作    用：登录操作
'参    数：
'==========================================
Sub LoginDo()
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request("code1")))
	Fk_Admin_LoginPass=Trim(Request("code2"))
if instr(Request.ServerVariables("HTTP_REFERER"),"win.qebang.net")<1  then
	Response.Write("警告：非法登录！") '非法登录
	Response.end
else
	if Md5(Md5(Fk_Admin_LoginName,32),32)=Fk_Admin_LoginPass then
		Response.Cookies("FkAdminName")="admin"
		Response.Cookies("FkAdminPass")="574fc0017b58dbd2"
		'Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass,32),16)
		Response.Cookies("FkAdminIp")=Request.ServerVariables("REMOTE_ADDR")
		Response.Cookies("FkAdminTime")=Now()
		Response.Cookies("FkAdminName").Expires=#May 10,2020#
		Response.Cookies("FkAdminPass").Expires=#May 10,2020#
		Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		Application.Lock()
		Conn.Execute(Sqlstr)
		Application.UnLock()
		Server.Transfer "/admin/Index-shangwin.asp"
		'Response.Write("1") '登录成功
	else
		Response.Write("非法登录！") '登录失败
		Response.end
	end if
end if
End Sub
%><!--#Include File="../Code.asp"-->