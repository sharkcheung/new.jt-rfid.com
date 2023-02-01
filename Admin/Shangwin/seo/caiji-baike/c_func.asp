<!--#Include File="../../../../Inc/Conn.asp"-->
<%
On Error Resume Next
Server.ScriptTimeOut=9999999
response.Charset="utf-8"
session.CodePage=65001
Function getHTTPPage(Path,Cset)
t = GetBody(Path)
getHTTPPage=BytesToBstr(t,Cset)
End function

dim conn,connstr,dbpath
dbpath=server.MapPath(SiteData)
sub openConn()
   on error resume next
   connstr="provider=microsoft.jet.oledb.4.0;data source="&dbpath
   set conn=server.CreateObject("adodb.connection")
   conn.open connstr
   if err then
      err.clear
	  response.end
   end if
end sub   

sub closeConn()
   on error resume next
   if isobject(conn) then
      conn.close
	  set conn=nothing
   end if
end sub

'
function ltrimVBcrlf(str)
dim pos,isBlankChar
pos=1
isBlankChar=true
while isBlankChar
if mid(str,pos,1)=" " then
pos=pos+1
elseif mid(str,pos,2)=VBcrlf then
pos=pos+2
else
isBlankChar=false
end if
wend
ltrimVBcrlf=right(str,len(str)-pos+1)
end function

Function GetBody(url) 
on error resume next
Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
With Retrieval 
.Open "Get", url, False, "", "" 
.Send 
GetBody = .ResponseBody
End With 
Set Retrieval = Nothing 
End Function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function

function selectlist()
call openConn()
dim iii,Rs,Sql,Fk_Module_Name,Fk_Module_id,newsclass,Fk_Module_Name2,Fk_Module_id2,Rs2,Sql2
selectlist="<select name=""mdltype"" id=""mdltype"">"
iii=1
Set Rs=conn.execute("Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0")
if not Rs.eof then
	do until Rs.EOF
	   Fk_Module_Name=Rs("Fk_Module_Name")
	   Fk_Module_id=Rs("Fk_Module_id")
	   newsclass=newsclass&"<option value='"&Fk_Module_id&"'>"&Fk_Module_Name&"</option>"
		
		Set Rs2=conn.execute("Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_id)
		if not Rs2.eof then
		do until Rs2.EOF
		   Fk_Module_Name2=Rs2("Fk_Module_Name")
		   Fk_Module_id2=Rs2("Fk_Module_id")
		   newsclass=newsclass&"<option value='"&Fk_Module_id2&"'>&nbsp; &nbsp; &nbsp; "&Fk_Module_Name2&"</option>"
		Rs2.MoveNext
		iii=iii+1
		loop
		end if
		Rs2.close
		Set Rs2=Nothing
		
	Rs.MoveNext
	loop
end if	
selectlist=selectlist&newsclass&"</select>"
Rs.close
Set Rs=Nothing
call closeConn()
end function

Function strCut(strContent,StartStr,EndStr,CutType)
	Dim strHtml,S1,S2
	strHtml = strContent
	On Error Resume Next
	Select Case CutType
	Case 1
		S1 = InStr(strHtml,StartStr)
		S2 = InStr(S1,strHtml,EndStr)+Len(EndStr)
	Case 2
		S1 = InStr(strHtml,StartStr)+Len(StartStr)
		S2 = InStr(S1,strHtml,EndStr)
	End Select
	If Err Then
		strCute = ""
		Err.Clear
		Exit Function
	Else
		strCut = Mid(strHtml,S1,S2-S1)
	End If
End Function


Function py_z_replace(str,r,s)
On Error Resume Next
Set re = New RegExp 
re.IgnoreCase = true
re.Global = True
re.Pattern = r
py_z_replace = re.Replace(str,s)
set re = nothing
If Err.Number <> 0 Then
	py_z_replace = err.description
	On Error GoTo 0
Else
	py_z_replace = py_z_replace  
	On Error GoTo 0
End If  
End Function

Function CUrl(t)
Domain_Name = LCase(Request.ServerVariables("Server_Name"))
Page_Name = LCase(Request.ServerVariables("Script_Name"))
'Quary_Name = LCase(Request.ServerVariables("Quary_String"))
If t =0 Then
CUrl = "http://"&Domain_Name
Else
CUrl = "http://"&Domain_Name&Page_Name
End If
End Function

Public Function ServerPath()
Dim Path
Dim Pos
Path="http://" & Request.ServerVariables("server_name") & Request.ServerVariables("script_name")
Pos=InStrRev(Path,"/")
ServerPath=Left(Path,Pos)
End Function


Dim wstr,str,url,start,over,city
%>
