<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Include.asp
'文件用途：管理员控制
'版权所有：企帮网络www.qebang.cn
'==========================================
'验证管理员
If Request.Cookies("FkAdminName")="" Or Request.Cookies("FkAdminPass")="" Then
	Response.Redirect("/admin/")
	Response.End()
End If
%>