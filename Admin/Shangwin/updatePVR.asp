<%@language="vbscript" codepage="65001"%> 
<!--#include file="easp.asp"-->
<%dim XmlHttpData,posturl,u,k,r,bdrank,ggrank
Server.ScriptTimeout=999999
response.Charset="utf-8"
session.CodePage=65001
asp.noCache()
u=asp.r("u",0)
'u="qebang.cn"
k=asp.r("k",0)
DataToSend = "a=update&d="&u&"&k="&Easp_Escape(k)
dim xmlhttp
set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
xmlhttp.Open "POST","http://win.qebang.net/web/pvr/update_pvr.asp",false
'xmlhttp.Open "POST","http://localhost:88/update_pvr.asp",false
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.send DataToSend
'asp.w xmlhttp.ResponseText
Set xmlhttp = Nothing
%>