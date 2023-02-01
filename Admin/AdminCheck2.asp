<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Include.asp
'文件用途：管理员控制
'版权所有：企帮网络www.qebang.cn
'==========================================
'验证管理员
If Request.Cookies("FkAdminId")="" Or Request.Cookies("FkAdminLimitId")="" Then
	Response.Redirect("Index.asp")
	Response.End()
End If
%>