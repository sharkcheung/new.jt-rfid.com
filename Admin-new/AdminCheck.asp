<!--#Include File="Include.asp"--><%
'==========================================
'�� �� ����Include.asp
'�ļ���;������Ա����
'��Ȩ���У��������www.qebang.cn
'==========================================
'��֤����Ա
If Request.Cookies("FkAdminName")="" Or Request.Cookies("FkAdminPass")="" Then
	Response.Redirect("/admin/")
	Response.End()
End If
%>