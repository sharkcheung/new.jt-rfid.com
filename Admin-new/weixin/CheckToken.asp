<!--#Include File="../Include.asp"-->
<!--#Include File="../../Inc/Md5.asp"-->
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


if Trim(Request("mobile"))<>"" and Trim(Request("usertype"))<>"" and Trim(Request("token"))<>"" then
	strMobile=FKFun.HTMLEncode(Trim(Request("mobile")))
	strUsertype=FKFun.HTMLEncode(Trim(Request("usertype")))
	strToken=FKFun.HTMLEncode(Trim(Request("token")))
else
	if Request.Cookies("FkAdminName")<>"" and Request.Cookies("Usertype")<>"" and Request.Cookies("token")<>"" then
		strMobile=Request.Cookies("FkAdminName")
		strUsertype=Request.Cookies("Usertype")
		strToken=Request.Cookies("token")
	else
		errHtml "非法操作，000"
	end if
end if

Call FKFun.ShowString(strMobile,1,50,0,"非法操作，001","非法操作，001")
Call FKFun.ShowString(strUsertype,1,50,0,"非法操作，002","非法操作，002")
Call FKFun.ShowString(strToken,1,50,0,"非法操作，003","非法操作，003")

token="3PVcDkYEbL8dXuaTM5JUzNjbPCWRuQq5"
strWebToken = MD5(strMobile & token &strUsertype, 32)
tokenpara = "mobile="&strMobile&"&token="&MD5(strMobile & token, 32)
''response.write strToken&"_"&strWebToken & "_" & strUsertype & "_" & strMobile
if strToken<>strWebToken then
	errHtml "非法操作，004"
end if

Response.Cookies("FkAdminName")	=strMobile
Response.Cookies("Usertype")	=strUsertype
Response.Cookies("token")		=strToken
Response.Cookies("FkAdminPass")	=Md5(Md5(strToken,32),16)
Response.Cookies("FkAdminIp")	=Request.ServerVariables("REMOTE_ADDR")
Response.Cookies("FkAdminTime")	=Now()
Response.Cookies("FkAdminName").Expires=#May 10,2030#
Response.Cookies("FkAdminPass").Expires=#May 10,2030#
Response.Cookies("Usertype").Expires=#May 10,2030#
Response.Cookies("token").Expires=#May 10,2030#

sub errHtml(str)
	response.write str
	response.end
end sub
%>