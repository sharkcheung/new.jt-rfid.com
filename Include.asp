<!--#Include File="Inc/Config.asp"-->
<!--#Include File="Inc/qb_safe3.asp"-->
<!--#Include File="Inc/PageCode.asp"-->
<!--#Include File="Class/Cls_Template.asp"-->
<!--#Include File="Class/Cls_PageCode.asp"--><%
'==========================================
'文 件 名：Include.asp
'文件用途：前台总控文件
'版权所有：企帮网络www.qebang.cn
'==========================================
If SiteOpen=0 Then
	Response.Write("站点维护中，请稍候访问！")
	Response.End()
End If
Call FKDB.DB_Open()
%>
