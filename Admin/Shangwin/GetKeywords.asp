<!--#Include File="../Include.asp"-->
<%
response.Charset="utf-8"
session.CodePage=65001
Dim iii ,t,htmls
t=request("t")
If t="gkw" then
Call FKDB.DB_Open()
	on error resume Next
	rs.open "select * from keywordSV",conn,1,1
	If Not rs.eof Then
		iii=0
		Do While Not rs.eof
			If iii=0 Then
				htmls=rs("SVkeywords")&"{22}"&rs("SVci")&"{22}"&rs("SVpaiming")
			else
				htmls=htmls&"{11}"&rs("SVkeywords")&"{22}"&rs("SVci")&"{22}"&rs("SVpaiming")
			End if
			iii=iii+1
		rs.movenext
		If rs.eof Then Exit do
		loop
	End if
	rs.close
	response.write htmls
Call FKDB.DB_Close()
End if
%>