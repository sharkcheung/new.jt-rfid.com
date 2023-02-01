<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
</head>

<%
u_name=trim(request("u_name"))
u_sex=int(request("u_sex"))
u_pass=trim(request("u_pass"))
u_pass_re=trim(request("u_pass_re"))
u_ask=request("u_ask")
u_answer=trim(request("u_answer"))
u_mail=trim(request("u_mail"))
CheckCode=trim(request("CheckCode"))

u_name_zs=trim(request("u_name_zs"))
member_tel=request("member_tel")
member_mobile=request("member_mobile")
member_address=trim(request("member_address"))
m_uage=int(request("member_age"))
member_zip=request("member_zip")
member_web=trim(request("member_web"))
u_qq=trim(request("u_qq"))

u_ip=request.ServerVariables("HTTP_X_FORWARDED_FOR")
if u_ip="" then u_ip=request.ServerVariables("REMOTE_ADDR")


if u_name="" or u_name_zs="" or u_sex="" or u_pass="" or u_pass_re="" or  u_ask="" or  u_answer="" or  u_mail="" or  CheckCode="" then
   call errmsg("请将信息填写完整!")
end if

if CheckCode<>trim(session("validateCode")) then
   call errmsg("验证码不正确!")
end if

if len(u_name)>10 or len(u_name)<4 then
   call errmsg("账号填写不正确!")
end if

if len(u_pass)>16 or len(u_pass)<6 then
   call errmsg("密码填写不正确!")
end if

if len(u_pass_re)>16 or len(u_pass_re)<6 then
   call errmsg("确认密码填写不正确!")
end if

set urs = Server.CreateObject("ADODB.RecordSet")
sqlu="select * from u_members where m_uid='"&u_name&"'"
urs.open sqlu,connn,1,3
if not urs.eof then
   urs.close
   set urs=nothing
   response.Write "<script language=javascript>alert('该用户名已被注册！请重新注册');history.back();</script>"
   response.end
else
   set mrs=connn.execute("select m_uemail from u_members where m_uemail='"&u_mail&"'")
   if not mrs.eof then
      mrs.close
	  set mrs=nothing
      response.Write "<script language=javascript>alert('该邮箱已被使用！请换个邮箱');history.back();</script>"
	  response.End
   else
      urs.addnew
      urs("m_uid")=u_name
      urs("m_uname")=u_name_zs
      urs("m_uaddress")=member_address
      urs("m_utel")=member_tel
      urs("m_umobile")=member_mobile
      urs("m_uQQ")=u_qq
      urs("m_uzip")=member_zip
      urs("m_upass")=Md5(Md5(u_pass,32),16)
      urs("m_reg_time")=now()
      urs("m_reg_ip")=u_ip
      urs("m_usex")=u_sex
      urs("m_uage")=m_uage
      urs("m_question")=u_ask
      urs("m_answer")=u_answer
      urs("m_uemail")=u_mail
      urs("m_uFobid")=0
      urs("m_login_count")=0
      urs.update
      urs.close
      set urs=nothing
   end if
   mrs.close
   set mrs=nothing
end if
response.Cookies("login")("u_id")=uname
response.Cookies("login")("u_pass")=Md5(Md5(u_pass,32),16)
session("u_id")=request.Cookies("login")("u_id")
session("u_pass")=request.Cookies("login")("u_pass")
session("qb_login")="qb_yes"
response.Write "<script language=javascript>alert('注册成功!');window.top.location.href='/';</script>"
response.end
sub errmsg(msg)
   response.Write "<script language=javascript>alert('"&msg&"');history.back();</script>"
   response.end
end sub
closeConn()
%>
