<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：GBookDo.asp
'文件用途：咨询提交
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_GBook_Title,Fk_GBook_Content,Fk_GBook_Name,Fk_GBook_Contact,Fk_GBook_Module
Dim S

'获取功能选项参数
Types=Clng(Request.QueryString("Type"))
S=Request.QueryString("S")
If S<>"1" Then
	S=0
Else
	S=1
End If
Select Case Types
	Case 1
		Call GBookAddDo() '添加咨询
	Case Else
		Response.Write arrTips(10)
End Select

'==============================
'函 数 名：GBookAddDo
'作    用：添加咨询
'参    数：
'==============================
Sub GBookAddDo()
	Fk_GBook_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_GBook_Title")))
	Fk_GBook_Content=FKFun.HTMLEncode(Trim(Request.Form("Fk_GBook_Content")))
	Fk_GBook_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_GBook_Name")))
	Fk_GBook_Contact=FKFun.HTMLEncode(Trim(Request.Form("Fk_GBook_Contact")))
	Fk_GBook_Module=Trim(Request.Form("Fk_GBook_Module"))
	Call FKFun.AlertString(Fk_GBook_Title,1,50,0,arrTips(11),arrTips(12))
	Call FKFun.AlertString(Fk_GBook_Content,1,500,0,arrTips(13),arrTips(14))
	Call FKFun.AlertString(Fk_GBook_Name,1,50,0,arrTips(15),arrTips(16))
	Call FKFun.AlertString(Fk_GBook_Contact,1,50,0,arrTips(17),arrTips(18))
	Call FKFun.AlertNum(Fk_GBook_Module,arrTips(19))
	If SiteNoTrash=1 Then
		Call FKFun.NoTrash(Fk_GBook_Content)
	End If
	Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Title='"&Fk_GBook_Title&"' And Fk_GBook_Name='"&Fk_GBook_Name&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_GBook_Title")=Fk_GBook_Title
		Rs("Fk_GBook_Content")=Fk_GBook_Content
		Rs("Fk_GBook_Name")=Fk_GBook_Name
		Rs("Fk_GBook_Contact")=Fk_GBook_Contact
		Rs("Fk_GBook_Module")=Fk_GBook_Module
		Rs("Fk_GBook_Ip")=Request.ServerVariables("REMOTE_ADDR")
		Rs.Update()
		Application.UnLock()
		If FetionNum<>"" And FetionPass<>"" Then
			Call FKFun.SmsGo("有新咨询--"&Fk_GBook_Title&"！")
		End If
		If S=0 Then
			Call FKFun.AlertInfo(arrTips(20),SiteDir)
		Else
			Call FKFun.AlertInfo(arrTips(21),SiteDir&"GBookFrame.asp?Id="&Fk_GBook_Module)
		End If
	Else
		If S=0 Then
			Call FKFun.AlertInfo(arrTips(22),SiteDir)
		Else
			Call FKFun.AlertInfo(arrTips(22),SiteDir&"GBookFrame.asp?Id="&Fk_GBook_Module)
		End If
	End If
	Rs.Close
End Sub
%>
<!--#Include File="Code.asp"-->
