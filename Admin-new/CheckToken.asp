<!--#Include File="Include.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<%
'==========================================
'文 件 名：Include.asp
'文件用途：管理员控制
'版权所有：企帮网络www.qebang.cn
'==========================================
'验证token

dim token,strMobile,strWebToken,strUsertype,strToken,tokenpara,cur_domain,strkfurl,strtjurl
dim pathfilename
cur_domain=Request.ServerVariables("SERVER_NAME")
pathfilename=lcase(trim(Request.ServerVariables("SCRIPT_NAME")))

strMobile=FKFun.HTMLEncode(Trim(Request("mobile")))
strUsertype=FKFun.HTMLEncode(Trim(Request("usertype")))
strToken=FKFun.HTMLEncode(Trim(Request("token")))
strtjurl=FKFun.HTMLEncode(Trim(Request("strtjurl")))
strkfurl=FKFun.HTMLEncode(Trim(Request("strkfurl")))
if strMobile="" or strUsertype="" or strToken="" then
	strMobile=Request.Cookies("FkAdminName")
	strUsertype=Request.Cookies("Usertype")
	strToken=Request.Cookies("token")
	strtjurl = Request.cookies("strtjurl")
	strkfurl = Request.cookies("strkfurl")
else
	Response.Cookies("FkAdminName")	=strMobile
	Response.Cookies("Usertype")	=strUsertype
	Response.Cookies("token")		=strToken
	Response.Cookies("strtjurl")	=strtjurl
	Response.Cookies("strkfurl")	=strkfurl
end if

Call FKFun.ShowString(strMobile,1,50,0,"非法操作，001","非法操作，001")
Call FKFun.ShowString(strUsertype,1,50,0,"非法操作，002","非法操作，002")
Call FKFun.ShowString(strToken,1,50,0,"非法操作，003","非法操作，003")
'Call FKFun.ShowString(strtjurl,1,300,0,"非法操作，004","非法操作，004")
'Call FKFun.ShowString(strkfurl,1,300,0,"非法操作，005","非法操作，005")
token="3PVcDkYEbL8dXuaTM5JUzNjbPCWRuQq5"
strWebToken = MD5(strMobile & token &strUsertype, 32)
tokenpara = "mobile="&strMobile&"&token="&MD5(strMobile & token, 32)
'response.write strToken&"_"&strWebToken 
if strToken<>strWebToken then
	errHtml "非法操作，004"
end if

Response.Cookies("FkAdminName")	=strMobile
Response.Cookies("Usertype")	=strUsertype
Response.Cookies("token")		=strToken
Response.Cookies("strkfurl")	=strkfurl
Response.Cookies("strtjurl")	=strtjurl
Response.Cookies("FkAdminPass")	=Md5(Md5(strToken,32),16)
Response.Cookies("FkAdminIp")	=Request.ServerVariables("REMOTE_ADDR")
Response.Cookies("FkAdminTime")	=Now()
Response.Cookies("FkAdminName").Expires=#May 10,2030#
Response.Cookies("FkAdminPass").Expires=#May 10,2030#
Response.Cookies("Usertype").Expires=#May 10,2030#
Response.Cookies("token").Expires=#May 10,2030#
Response.Cookies("strtjurl").Expires=#May 10,2030#
Response.Cookies("strkfurl").Expires=#May 10,2030#

sub errHtml(str)
	response.write str
	response.end
end sub
%>