<!--#Include File="Include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<!--#Include File="nocache.asp"-->
<%
'==========================================
'用途：商赢快车软件端登录验证
'==========================================

'定义页面变量
Dim Fk_Admin_LoginName,Fk_Admin_LoginToken,Fk_Admin_UserType,Fk_Admin_Name,Fk_Admin_Limit, token
dim op
op=FKFun.HTMLEncode(Trim(Request("op")))
token="3PVcDkYEbL8dXuaTM5JUzNjbPCWRuQq5"
if op = "sync_login" then
Call LoginDo() '登录操作
elseif op= "sync_word" then
Call LoginWord() '跳转到关键内链
end if

'转换时间 时间格式化 
Function formatDate(Byval t,Byval ftype) 
	dim y, m, d, h, mi, s 
	formatDate=""
	If IsDate(t)=False Then Exit Function
	y=cstr(year(t)) 
	m=cstr(month(t)) 
	If len(m)=1 Then m="0" & m 
	d=cstr(day(t)) 
	If len(d)=1 Then d="0" & d 
	h = cstr(hour(t)) 
	If len(h)=1 Then h="0" & h 
	mi = cstr(minute(t)) 
	If len(mi)=1 Then mi="0" & mi 
	s = cstr(second(t)) 
	If len(s)=1 Then s="0" & s 
	select case cint(ftype) 
	case 1 
	' yyyy-mm-dd 
	formatDate=y & "-" & m & "-" & d 
	case 2 
	' yy-mm-dd 
	formatDate=right(y,2) & "-" & m & "-" & d 
	case 3 
	' mm-dd 
	formatDate=m & "-" & d 
	case 4 
	' yyyy-mm-dd hh:mm:ss 
	formatDate=y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
	case 5 
	' hh:mm:ss 
	formatDate=h & ":" & mi & ":" & s 
	case 6 
	' yyyy年mm月dd日 
	formatDate=y & "年" & m & "月" & d & "日"
	case 7 
	' yyyymmdd 
	formatDate=y & m & d 
	case 8 
	'yyyymmddhhmmss 
	formatDate=y & m & d & h & mi & s 
	case 9
	' yyyy-mm-dd hh:mm:ss 
	formatDate=y & "-" & m & "-" & d 
	end select 
End Function

sub chkToken(strMobile,strUsertype,strToken)
	dim strTime,strWebToken
	Call FKFun.ShowString(strMobile,1,50,0,"非法操作，001","非法操作，001")
	Call FKFun.ShowString(strUsertype,1,50,0,"非法操作，002","非法操作，002")
	Call FKFun.ShowString(strToken,1,50,0,"非法操作，003","非法操作，003")
    strTime = formatDate(Now, 9)
    strWebToken = MD5(strMobile & token & strUsertype & strTime, 32)
	' response.write strToken&"_"&strWebToken
	if strToken<>strWebToken then
		response.write "非法操作，004"
		response.end
	end if
end sub

'==========================================
'函 数 名：LoginDo()
'作    用：登录操作
'参    数：
'==========================================
Sub LoginDo()
	Fk_Admin_LoginName=FKFun.HTMLEncode(Trim(Request("mobile")))
	Fk_Admin_Name=FKFun.HTMLEncode(Trim(Request("truename")))
	Fk_Admin_UserType=FKFun.HTMLEncode(Trim(Request("userType")))
	Fk_Admin_LoginToken=FKFun.HTMLEncode(Trim(Request("token")))
	call chkToken(Fk_Admin_LoginName,Fk_Admin_UserType,Fk_Admin_LoginToken)
	
	'读取权限表，判断数据库是否存在裁切权限
	Sqlstr="Select top 1 Fk_Limit_Id,Fk_Limit_Content,Fk_Limit_Name From [Fk_Limit] Where Fk_Limit_Name='裁切'"
	Rs.Open Sqlstr,Conn,1,3
	if not rs.eof then
		Fk_Admin_Limit=rs("Fk_Limit_Id")
	else
		'无裁切权限，则添加
		rs.addnew()
		rs("Fk_Limit_Content")=",System1,System11,System2,System5,System6,System10,System13,System15,System21,System22,System16,System3,System9,"
		rs("Fk_Limit_Name")="裁切"
		rs.update()
		Fk_Admin_Limit=rs("Fk_Limit_Id")
	end if
	rs.close
	
	'判断后台用户表中是否存在此账号
	Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_LoginName='"&Fk_Admin_LoginName&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		'不存在则判断用户类型，如果uertype=99，则是超级管理员，无需添加账号，如果usertype=20,则是裁切，添加账号且类型为裁切
		if Fk_Admin_UserType="99" then
		else
			if Fk_Admin_UserType<>"20" then
				Fk_Admin_Limit= 1
			end if
			Application.Lock()
			rs.addnew()
			rs("Fk_Admin_LoginName")=Fk_Admin_LoginName
			rs("Fk_Admin_LoginPass")=Md5(Md5(Fk_Admin_LoginToken,32),16)
			rs("Fk_Admin_Name")		=Fk_Admin_Name
			rs("Fk_Admin_Limit")	=Fk_Admin_Limit
			rs("Fk_Admin_User")		=1
			rs.update()
			Application.UnLock()
		end if
	else
		if Fk_Admin_UserType<>"99" then
			if Fk_Admin_UserType<>"20" then
				Fk_Admin_Limit= 1
			end if
			if Fk_Admin_Limit<>rs("Fk_Admin_Limit") then
				Application.Lock()
				rs("Fk_Admin_Limit")	=Fk_Admin_Limit
				rs.update()
				Application.UnLock()
			end if
		end if
	end if
	rs.close
	Response.Cookies("FkAdminName")	=Fk_Admin_LoginName
	Response.Cookies("FkAdminPass")	=Md5(Md5(Fk_Admin_LoginToken,32),16)
	Response.Cookies("FkAdminIp")	=Request.ServerVariables("REMOTE_ADDR")
	Response.Cookies("FkAdminTime")	=Now()
	Response.Cookies("FkAdminName").Expires=#May 10,2030#
	Response.Cookies("FkAdminPass").Expires=#May 10,2030#
	
	Sqlstr="Insert Into [Fk_Log](Fk_Log_Text,Fk_Log_Ip) Values('用户“"&Fk_Admin_LoginName&"”成功登录！','"&Request.ServerVariables("REMOTE_ADDR")&"')"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	response.redirect "index-shangwin.asp"
End Sub

sub LoginWord()
	dim strtjurl, strkfurl, strMobile, strUsertype, strToken, strWebToken
	strtjurl=FKFun.HTMLEncode(Trim(Request("tjurl")))
	strkfurl=FKFun.HTMLEncode(Trim(Request("kfurl")))
	strMobile=FKFun.HTMLEncode(Trim(Request("mobile")))
	strUsertype=FKFun.HTMLEncode(Trim(Request("usertype")))
	strToken=FKFun.HTMLEncode(Trim(Request("token")))
	if strtjurl="" or strkfurl="" or strMobile="" or strUsertype="" or strToken="" then
		strtjurl = Request.cookies("strtjurl")
		strkfurl = Request.cookies("strkfurl")
		strMobile=Request.Cookies("FkAdminName")
		strUsertype=Request.Cookies("Usertype")
		strToken=Request.Cookies("token")
	else
		Response.Cookies("FkAdminName")	=strMobile
		Response.Cookies("Usertype")	=strUsertype
		Response.Cookies("strkfurl")	=strkfurl
		Response.Cookies("strtjurl")	=strtjurl
		Response.Cookies("token")	=strToken
	end if
	Call FKFun.ShowString(strtjurl,1,600,0,"非法操作，004","非法操作，004")
	Call FKFun.ShowString(strkfurl,1,600,0,"非法操作，005","非法操作，005")
	
    strWebToken = MD5(strMobile & token & strUsertype , 32)
	if strToken<>strWebToken then
		response.write "非法操作，001"
		response.end
	end if
	response.redirect "index-word.asp?type=1&mobile="&strMobile&"&userType="&strUsertype&"&token="&strToken&"&tjurl="&Server.URLEncode(strtjurl)&"&kfurl="& Server.URLEncode(strkfurl)
end sub
%><!--#Include File="../Code.asp"-->