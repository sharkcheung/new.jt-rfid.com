<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
</head>

<%dim action,m_uid
action=request.QueryString("action")
m_uid=session("u_id")
select case action
'//收货人信息
case "shouhuoxx"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from u_members where m_uid='"&m_uid&"' ",connn,1,3
'rs("shouname")=trim(request("shouname"))
'rs("shengshi")=trim(request("shengshi"))
'rs("shouhuodizhi")=trim(request("shouhuodizhi"))
'rs("youbian")=cstr(request("youbian"))
'rs("usertel")=trim(request("usertel"))
'rs("songhuofangshi")=int(request("songhuofangshi"))
'rs("zhifufangshi")=int(request("zhifufangshi"))
'rs("shousex")=int(request("shousex"))

'rs("reglx")=2	'设用户为修改了资料
rs("dadyname")=trim(request("dadyname"))
rs("szSheng")=trim(request("szSheng"))
rs("szShi")=trim(request("szShi"))
rs("shouhuodizhi")=trim(request("shouhuodizhi"))
rs("youbian")=trim(request("youbian"))
rs("usermobile")=trim(request("usertel"))
rs("MoMname")=trim(request("MoMname"))
rs("MoMNo")=trim(request("MoMNo"))
rs("yuchan")=trim(request("yuchan"))
rs("babysex")=trim(request("babysex"))
rs("songhuofangshi")=trim(request("songhuofangshi"))
rs("zhifufangshi")=trim(request("zhifufangshi"))

rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('您的详细资料信息保存成功！');</script>"
response.redirect "myuser.asp?action=shouhuoxx"
response.End

'//用户资料
case "userziliao"
qq=trim(request("qq"))
useremail=trim(request("useremail"))
userzhenshiname=trim(request("userzhenshiname"))
sfz=trim(request("sfz"))
shousex=trim(request("shousex"))
nianling=trim(request("nianling"))
hukouprovince=trim(request("hukouprovince"))
shouhuodizhi=trim(request("shouhuodizhi"))
usertel=trim(request("usertel"))
usermobile=trim(request("usermobile"))
youbian=trim(request("youbian"))
hukouprovince=trim(request("hukouprovince"))
m_answer=trim(request("m_answer"))
if useremail="" then
   response.Write "<script language=javascript>alert('请填写邮箱地址!');history.back();</script>"
   response.end
end if
if userzhenshiname="" then
   response.Write "<script language=javascript>alert('请填写真实姓名!');history.back();</script>"
   response.end
end if
if m_answer="" then
   response.Write "<script language=javascript>alert('请填写安全答案!');history.back();</script>"
   response.end
end if
'if sfz="" then
 '  response.Write "<script language=javascript>alert('请填写身份证号码!');history.back();/script>"
  ' response.end
'end if
if shousex<>"" and not isnumeric(shousex) then
   response.Write "<script language=javascript>alert('请选择性别!');history.back();</script>"
   response.end
end if
if hukouprovince="" then
   response.Write "<script language=javascript>alert('请选择地址!');history.back();</script>"
   response.end
end if
if shouhuodizhi="" then
   response.Write "<script language=javascript>alert('请填写详细地址!');history.back();</script>"
   response.end
end if
if usertel="" and usermobile="" then
   response.Write "<script language=javascript>alert('联系电话、手机须填写一项!');history.back();</script>"
   response.end
end if
if youbian="" then
   response.Write "<script language=javascript>alert('请填写邮编!');history.back();</script>"
   response.end
end if
if qq<>"" and not isnumeric(qq) then
   response.Write "<script language=javascript>alert('请正确填写QQ号码!');history.back();</script>"
   response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where m_uid='"&m_uid&"'",connn,1,3
rs("m_uemail")=useremail
rs("m_uname")=userzhenshiname
rs("m_answer")=m_answer
rs("m_sfz")=sfz
rs("m_usex")=shousex
if nianling<>"" and not isnull(nianling) and isnumeric(nianling) then
   rs("m_uage")=nianling
end if
rs("szsheng")=request("hukouprovince")
rs("szshi")=request("hukoucapital")
rs("szxian")=int(request("hukoucity"))
rs("m_uaddress")=shouhuodizhi
rs("m_umobile")=usermobile
rs("m_utel")=usertel

rs("m_uzip")=int(youbian)
if qq<>"" then
rs("m_uQQ")=trim(qq)
end if
rs("m_uWeb")=trim(request("homepage"))
rs("content")=trim(request("content"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的个人资料修改成功！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
response.end

case "savepass"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where m_uid='"&m_uid&"'",connn,1,3
if trim(request("userpassword"))<>"" then
if trim(request("userpassword"))<>trim(request("userpassword2")) then
response.Write "<script language=javascript>alert('两次密码输入不一样!');window.location.href='member_center.asp?action=savepass';</script>"
response.end
else
rs("m_upass")=Md5(Md5(request("userpassword"),32),16)
end if
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('密码更改成功！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
response.End
end if


'//取回密码
case "repass"
set rs=server.CreateObject("adodb.recordset")
rs.open "select m_upass from u_members where m_uid='"&trim(request("username2"))&"'",connn,1,3
rs("m_upass")=Md5(Md5(request("userpassword2"),32),16)
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的密码取回成功，请登陆！');history.go(-1);</script>"
end select
%>