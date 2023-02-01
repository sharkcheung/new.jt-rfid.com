<!--#Include File="Include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<link href="Css/1.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../Js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="../Js/jquery.form.min.js"></script>
<script type="text/javascript" src="../Js/function.js"></script>
<script type="text/javascript" src="../Js/xheditor-zh-cn.min.js"></script>
</head>
<!--#Include File="nocache.asp"-->
<%
'==========================================
'文 件 名：Login.asp
'文件用途：用户登录拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Admin_LoginName,Fk_Admin_LoginPass

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call LoginBox() '读取登录信息
	Case 2
		Call LoginDo() '登录操作
End Select

'==========================================
'函 数 名：LoginBox()
'作    用：读取登录信息
'参    数：
'==========================================
Sub LoginBox()
%>
<div id="shangwin-login">
<form id="AdminLogin" name="AdminLogin" method="post" action="Login-shangwin.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:300px;">用户登录[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:300px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">用户名：</td>
	        <td>&nbsp;<input type="text" name="AdminName" id="AdminName" class="Input Input150" value='<%=request("name")%>' /></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">密码：</td>
	        <td>&nbsp;<input type="password" name="AdminPass" id="AdminPass" class="Input Input150" value='<%=request("pass")%>' /></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:280px;">
        <input type="submit" onclick="Sends('AdminLogin','Login-shangwin.asp?Type=2',1,'/admin/',0,0,'','');" class="Button" name="button" id="button" value="登 录" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
</div>
<%
End Sub

'==========================================
'函 数 名：LoginDo()
'作    用：登录操作
'参    数：
'==========================================
Sub LoginDo()
	'Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request.Form("AdminName")))
	'Fk_Admin_LoginPass=FKFun.HTMLEncode(Trim(Request.Form("AdminPass")))
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request("name")))
	Fk_Admin_LoginPass=FKFun.HTMLEncode(Trim(Request("pass")))
	Call FKFun.ShowString(Fk_Admin_LoginName,1,50,0,"请输入登录名！","登录名名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_LoginPass,1,50,0,"请输入登录密码！","登录密码不能大于50个字符！")
	Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&Fk_Admin_LoginName&"' And Fk_Admin_LoginPass='"&Md5(Md5(Fk_Admin_LoginPass,32),16)&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Response.Cookies("FkAdminName")=Fk_Admin_LoginName
		Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass,32),16)
		Response.Cookies("FkAdminIp")=Request.ServerVariables("REMOTE_ADDR")
		Response.Cookies("FkAdminTime")=Now()
		
		Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		Application.Lock()
		Conn.Execute(Sqlstr)
		Application.UnLock()
		'Response.Write("登录成功！")
		Server.Transfer("index-shangwin.asp") 
	Else
		Response.Write("用户或密码错误！")
	End If
	Rs.Close
End Sub
%>
<!--#Include File="../Code.asp"-->