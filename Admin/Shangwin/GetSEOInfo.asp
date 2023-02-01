<!--#Include File="../Include.asp"-->
<%
response.Charset="utf-8"
session.CodePage=65001
Dim c ,t,dn
t=request("t")
If t=0 then
Call FKDB.DB_Open()
	on error resume Next
	c=0
	rs.open "select SVci from keywordSV where SVci<> null",conn,1,1
	If Not rs.eof Then
		Do While Not rs.eof
			c=c+rs("SVci")
		rs.movenext
		If rs.eof Then Exit do
		loop
	End if
	rs.close
	response.write c
Call FKDB.DB_Close()
ElseIf t=1 Then
Dim objXML
dn=request("dn")
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
with objXML
	.open "GET","http://win.qebang.net/web/service/getservertime.ashx?dn="&dn&"",false
	.send(null)
	response.write .responseText
end with
Set objXML=nothing
End if
%>