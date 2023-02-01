<!--#Include File="Include.asp"-->
<%
'==========================================
'文 件 名：Subject.asp
'文件用途：专题页
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义变量
Dim PageCode,PageCodes,TemplateId,Fk_Subject_Name
Dim FKTemplate
Dim Rst
Set Rst=Server.Createobject("Adodb.RecordSet")
Set FKTemplate=New Cls_Template

'获取变量
Id=Clng(Request.QueryString("Id"))

'获取专题
Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Id
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	Fk_Subject_Name=Rs("Fk_Subject_Name")
	TemplateId=Rs("Fk_Subject_Template")
Else
	Rs.Close
	Call FKDB.DB_Close()
	Response.Write("专题不存在！")
	Response.End()
End If
Rs.Close

'抽取模板内容
If TemplateId=0 Then
	Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='subject'"
Else
	Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
End If
Rs.Open Sqlstr,Conn,1,3
If Not Rs.Eof Then
	PageCode=Rs("Fk_Template_Content")
	Templates=Rs("Fk_Template_Name")
Else
	Rs.Close
	Call FKDB.DB_Close()
	Response.Write("专题页模板获取失败！")
	Response.End()
End If
Rs.Close
If SiteTest=1 Then
	PageCode=FKFso.FsoFileRead("Skin/"&SiteTemplate&"/"&Templates&".html")
End If
'处理函数
PageCode=FKTemplate.FileChange(PageCode)
PageCode=FKTemplate.SiteChange(PageCode)
PageCode=FKTemplate.SubjectChange(PageCode)
PageCode=FKTemplate.TemplateDo(PageCode)

Response.Write(PageCode)

%>
<!--#Include File="Code.asp"-->
