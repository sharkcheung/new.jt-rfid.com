<!--#Include File="../Inc/Config.asp"--><%
'==========================================
'文 件 名：Include.asp
'文件用途：管理员控制
'版权所有：企帮网络www.qebang.cn
'==========================================

'验证管理员

Call FKDB.DB_Open()

If Request.Cookies("FkAdminName")<>"" And Request.Cookies("FkAdminPass")<>"" Then
	Response.Cookies("FkAdminName")=FKFun.HTMLEncode(Request.Cookies("FkAdminName"))
	Response.Cookies("FkAdminPass")=FKFun.HTMLEncode(Request.Cookies("FkAdminPass"))
	'Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&Request.Cookies("FkAdminName")&"' And Fk_Admin_LoginPass='"&Request.Cookies("FkAdminPass")&"'"
	Sqlstr="Select * From [Fk_Admin] Where Fk_Admin_User=1 And Fk_Admin_LoginName='"&Request.Cookies("FkAdminName")&"' "
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Response.Cookies("FkAdminId")=Rs("Fk_Admin_Id")
		Response.Cookies("FkAdminLimitId")=Rs("Fk_Admin_Limit")
		If Rs("Fk_Admin_Limit")>0 Then
			Rs.Close
			Sqlstr="Select * From [Fk_Limit] Where Fk_Limit_Id=" & Request.Cookies("FkAdminLimitId")
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				Response.Cookies("FkAdminLimit")=Rs("Fk_Limit_Content")
			Else
				Response.Cookies("FkAdminLimit")="No"
			End If
		End If
		Login=True
	Else
		Response.Cookies("FkAdminId")=""
		Response.Cookies("FkAdminLimitId")=""
		Login=False
	End If
	Rs.Close
Else
	Response.Cookies("FkAdminLimitId")="100000"
	Response.Cookies("FkAdminId")=""
	Response.Cookies("FkAdminLimitId")=""
	Login=False
End If


'==============================
'函 数 名：ModuleSelectUrl
'作    用：输出ModuleSelectURL列表
'参    数：要输出的菜单MenuIds，要输出级数LevelId，默认选择AutoId
'==============================
Public Function ModuleSelectUrl(MenuIds,LevelId,AutoId)
	Call ModuleSelectUrlM(MenuIds,LevelId,"",AutoId)
End Function
Public Function ModuleSelectUrlM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
		If FkFun.CheckLimit("Module"&Rs2("Fk_Module_Id")) Then
	%>
					<option value="<%=GetNavGo(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"))%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
	<%
			If LevelId=0 Then
				TitleBacks="&nbsp;&nbsp;&nbsp;├"
			Else
				TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
			End If
			Call ModuleSelectUrlM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		End If
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：ModuleSelectId
'作    用：输出ModuleSelectId列表
'参    数：要输出的菜单MenuIds，要输出级数LevelId，默认选择AutoId
'==============================
Public Function ModuleSelectId(MenuIds,LevelId,AutoId)
	Call ModuleSelectIdM(MenuIds,LevelId,"",AutoId)
End Function
Public Function ModuleSelectIdM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack=""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
		'If FkFun.CheckLimit("Module"&Rs2("Fk_Module_Id")) Then
	%>
					<option value="<%=Rs2("Fk_Module_Id")%>"<%=FKFun.BeSelect(AutoId,Rs2("Fk_Module_Id"))%>><%=TitleBack%><%=Rs2("Fk_Module_Name")%></option>
	<%
			If LevelId=0 Then
				TitleBacks="&nbsp;&nbsp;&nbsp;├"
			Else
				TitleBacks="&nbsp;&nbsp;&nbsp;"&TitleBack
			End If
			Call ModuleSelectIdM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		'End If
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

'==============================
'函 数 名：GetNavGo
'作    用：输出分类级数参数
'参    数：要输出的模块ModuleLevelId
'==============================
Function GetNavGo(GetModuleType,GetModuleId)
	Select Case GetModuleType
		Case 0
			GetNavGo="layer.msg('静态模块无需内容修改，如有修改直接改模板！');"
		Case 1
			GetNavGo="SetRContent('MainRight','Article.asp?Type=1&ModuleId="&GetModuleId&"')"
		Case 2
			GetNavGo="SetRContent('MainRight','Product.asp?Type=1&ModuleId="&GetModuleId&"')"
		Case 3
			GetNavGo="ShowBox('Info.asp?Type=1&ModuleId="&GetModuleId&"','信息','940px');"
		Case 4
			GetNavGo="SetRContent('MainRight','GBook.asp?Type=1&ModuleId="&GetModuleId&"')"
		Case 5
			GetNavGo="layer.msg('转向链接无需内容修改！');"
		Case 6
			GetNavGo="layer.msg('招聘请直接通过内容设置菜单管理！');"
		Case 7
			GetNavGo="SetRContent('MainRight','Down.asp?Type=1&ModuleId="&GetModuleId&"')"
	End Select
End Function

Function AdminY()
if Request.Cookies("FkAdminName")<>"admin" then
 response.write " style='display:none;' "
end if
End Function


%>