<%
'==========================================
'文 件 名：Cls_DB.asp
'文件用途：数据库函数类
'==========================================

Class Cls_DB
	Private ConnStr
	Private DBpath,DBLink,DBi
	'==============================
	'函 数 名：DB_Conn
	'作    用：数据库连接函数
	'参    数：
	'==============================
	Private Sub DB_Conn()
		On Error Resume Next
		Set Conn = Server.CreateObject("Adodb.Connection")
		ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteData)
		Conn.Open ConnStr
		If Err Then
			Call AspErr()
		End If
	End Sub

	'==============================
	'函 数 名：DB_Open
	'作    用：创建读取对象
	'参    数：
	'==============================
	Public Sub DB_Open()
		Call DB_Conn()
		Set Rs=Server.Createobject("Adodb.RecordSet")
	End Sub
	
	'==============================
	'函 数 名：DB_Close
	'作    用：关闭读取对象
	'参    数：
	'==============================	
	Public Sub DB_Close()
		Set Rs=Nothing
		If IsObject(Conn) Then Conn.Close
		Set Conn = Nothing
	End Sub
	
	'==============================
	'函 数 名：AspErr
	'作    用：连接报错
	'参    数：
	'==============================	
	Private Sub AspErr()
		DBLink = Request.ServerVariables("url")
		DBLink = Split(DBLink,"/")
		For DBi = 0 To Ubound(DBLink)-1
			DBpath = DBpath&DBLink(DBi)&"/"
		Next
		If Application("DataErr")<>"" Then
			Application("DataErr")=Application("DataErr")+1
			If Application("DataErr")>=3 Then
				Response.Write "<body style='font-size:12px'>"
				Response.Write "错 误 号：" & Err.Number & "<br />"
				Response.Write "错误描述：" & Err.Description & "<br />"
				Response.Write "错误来源：" & Err.Source & "<br />"
				Response.Write "</body>"
			Else
				If Instr(LCase(Request.ServerVariables("Script_Name")),"/admin/")>0 Then
					Response.Redirect("../Install.asp")
				Else
					Response.Redirect("Install.asp")
				End If
			End If
		Else
			Application("DataErr")=1
			If Instr(LCase(Request.ServerVariables("Script_Name")),"/admin/")>0 Then
				Response.Redirect("../Install.asp")
			Else
				Response.Redirect("Install.asp")
			End If
		End If
		Err.Clear
		Response.End
	End Sub
End Class
%>
