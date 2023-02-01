<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
'Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
'Response.Expires=-999
Session.Timeout=999
%>

<head>
<link href="column.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- ymPrompt组件 -->
<script type="text/javascript" src="/admin/winskin/ymPrompt.js"></script>
<link rel="stylesheet" type="text/css" href="/admin/winskin/qq/ymPrompt.css" /> 
<!-- ymPrompt组件 -->
<script src="jquery-1.4.4.js" type="text/javascript"></script>

</head>
<body oncontextmenu="return false">
<%
 
'-----------------函数区--------------------------------------------
 
 
'输出分页 
sub pagelist(shuzi)
pageview="<a href='?pp=0'>1</a>"
for i=1 to shuzi
ii=i*15
ppp= "<a href='?pp="&ii&"'>"&i+1&"</a>"
pageview= pageview&" "&ppp
next
response.write pageview
end sub
 
 
 
  '获取页面的HTML
 
function getHTTPPage(url) 
dim Http
set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",url,false
Http.send()
if Http.readystate<>4 then 
exit function
end if
getHTTPPage=bytesToBSTR(Http.responseBody,"gb2312")
'getHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")

getHTTPPage=replace(getHTTPPage,"DoNews","企帮")  '批量替换donews
set http=nothing
if err.number<>0 then err.Clear 
end function

Function BytesToBstr(body,Cset)  '转换中文字符乱码
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

      


'截取字符串,1.包括前后字符串，2.不包括前后字符串
Function strCut(strContent,StartStr,EndStr,CutType)
Dim S1,S2
On Error Resume Next
Select Case CutType
Case 1
  S1 = InStr(strContent,StartStr)
  S2 = InStr(S1,strContent,EndStr)+Len(EndStr)
Case 2
  S1 = InStr(strContent,StartStr)+Len(StartStr)
  S2 = InStr(S1,strContent,EndStr)
End Select
If Err Then
  strCute = "<p align='center' ><font size=-1>截取字符串出错.</font></p>"
  Err.Clear
  Exit Function
Else
  strCut = Mid(strContent,S1,S2-S1)
End If
End Function

%>


<!--#Include File="../../../../inc/conn.asp"-->
<%
if request("act")="1" and Request.Form("T1")<>"" and Request.Form("S1")<>"" and Request.Form("D1")<>"" then  '保存采集

Set Conn = Server.CreateObject("Adodb.Connection")
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteData)
Conn.Open ConnStr

sql = "select top 1 * from [FK_Article ] where [Fk_Article_Title]='"&Request.Form("T1")&"'"
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,3
if rs.recordcount=0 then
  rs.addnew
  rs("Fk_Article_Title")=Request.Form("T1")
  rs("Fk_Article_Content")=Request.Form("S1")
  rs("Fk_Article_From")="互联网"
  rs("Fk_Article_Module")=Request.Form("D1")
  rs("Fk_Article_Menu")=1
  rs("Fk_Article_Recommend")=",0,"
  rs("Fk_Article_Subject")=",0,"
  rs.update
  response.write("<Script language=JavaScript>alert('采集保存成功！采集内容默认不在前台显示，请采集完成以后进入[网站]模块编辑内容、完善关键词后勾选[显示]。');</Script>")
else
  response.write("<Script language=JavaScript>alert('已有同名内容，请确认是否重复采集？');</Script>")
end if
if err<>0 then
  response.write("<Script language=JavaScript>alert('采集保存出错，请重试或放弃采集该信息。');</Script>")
end if
rs.close
set rs=nothing
conn.close
end if




if Request.Cookies("newsclass")="" then
call newsclasslist()
end if

sub newsclasslist() '获取新闻类列表函数
newsclass="<select size='1' name='D1'>"
Set Conn = Server.CreateObject("Adodb.Connection")
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteData)
Conn.Open ConnStr
iii=1
Set Rs=Server.Createobject("Adodb.RecordSet")
Sql="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0 "
	Rs.Open Sql,Conn,1,1
	do until rs.EOF
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_id=Rs("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id&"'>"&Fk_Module_Name&"↓</option>"
		
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sql2="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_id
		Rs2.Open Sql2,Conn,1,1
		do until rs2.EOF
		Fk_Module_Name2=Rs2("Fk_Module_Name")
		Fk_Module_id2=Rs2("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id2&"'>　"&Fk_Module_Name2&"</option>"
		rs2.MoveNext
		iii=iii+1
		loop
		rs2.close
		Set Rs2=Nothing
		
	rs.MoveNext
	loop
	    
	rs.close
	Set Rs=Nothing
	If IsObject(Conn) Then Conn.Close
	Set Conn = Nothing
	
newsclass=newsclass&"</select>"
		Response.Cookies("newsclass")=newsclass
		Response.Cookies("newsclass").Expires=#May 10,2020#
end sub
%>