<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
return_url=request.ServerVariables("HTTP_REFERER")
uname=trim(request("login_name"))
pass=trim(request("login_pwd"))
code=int(request("login_code"))
ip=getIP()
if code<>int(session("validateCode")) then
   response.Write "<script language=javascript>alert('验证码不正确!');history.back(-1);</script>"
   response.end
else
   if uname="" or isnull(uname) then
      response.Write "<script language=javascript>alert('会员账号不能为空!');history.back(-1);</script>"
      response.end
   else
      if pass="" or isnull(pass) then
         response.Write "<script language=javascript>alert('会员密码不能为空!');history.back(-1);</script>"
         response.end
	  else
	     set rs=connn.execute("select id,m_uid,m_upass,m_uFobid from u_members where m_uid='"&uname&"' and m_upass='"&Md5(Md5(pass,32),16)&"'")
		    if rs.eof then
		       rs.close
		       set rs=nothing
			   response.Write "<script language=javascript>alert('会员账号,密码不正确或无此用户!');history.back(-1);</script>"
               response.end
			else
			   if rs(3)=1 then
		          rs.close
		          set rs=nothing
			      response.Write "<script language=javascript>alert('此用户已被锁定,请与管理员联系!');history.back(-1);</script>"
                  response.end
			   else
			      connn.execute("update u_members set m_last_logintime="&db_date&",m_login_count=m_login_count+1,m_last_loginip='"&ip&"' where m_uid='"&uname&"'")
			      response.Cookies("login")("id")=rs(0)
			      response.Cookies("login")("u_id")=uname
			      response.Cookies("login")("u_pass")=Md5(Md5(pass,32),16)
				  session("id")=request.Cookies("login")("id")
				  session("u_id")=request.Cookies("login")("u_id")
				  session("u_pass")=request.Cookies("login")("u_pass")
				  session("qb_login")="qb_yes"
				  session("qb_fobid")="qb_no"
	'response.Write "<script language=javascript>window.top.location.href='"&return_url&"';</script>"
	response.Write "<script language=javascript>window.top.location.href='/';</script>"
                  response.end
			   end if
			end if
		 rs.close
		 set rs=nothing
      end if
   end if
end if
%>