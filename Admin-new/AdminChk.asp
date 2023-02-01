<%
response.charset="utf-8"
session.codepage=65001
If Request.Cookies("FkAdminName")="" Or Request.Cookies("FkAdminPass")="" Then
	Response.Redirect("/admin/")
	Response.End()
End If
%>