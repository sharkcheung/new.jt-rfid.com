<!--#Include File="../Inc/Config.asp"--><%
'==========================================

'==========================================
Response.Cookies("FkAdminName")=""
Response.Cookies("FkAdminPass")=""
Response.Cookies("FkAdminIp")=""
Response.Cookies("FkAdminTime")=""
Response.Redirect("/admin/")
Response.End()
%>