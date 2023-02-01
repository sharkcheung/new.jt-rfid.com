<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file = "inc2.asp" -->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
act=trim(request("act"))
select case act
   case "names"
      names=trim(request("name"))
	  if names="" then
	     response.Write "yes"
      else
	     set nrs=connn.execute("select id from u_members where m_uid='"&names&"'")
		 if not nrs.eof then
		    response.Write "yes"
	     else
		    response.Write "no"
		 end if
		 nrs.close
		 set nrs=nothing
	  end if
   case "email"
      mail=trim(request("mail"))
	  if mail="" then
	     response.Write "yes"
	  else
	     set nrs=connn.execute("select id from u_members where m_uemail='"&mail&"'")
		 if not nrs.eof then
		    response.Write "yes"
	     else
		    response.Write "no"
		 end if
		 nrs.close
		 set nrs=nothing
	  end if
   case "code"
      code=trim(request("code"))
	  if code="" then
	     response.Write "yes"
	  else
	     if code<>trim(session("validateCode")) then
		 	response.Write "yes"
		 else
		    response.Write "no"
		 end if
	  end if 
end select
%>