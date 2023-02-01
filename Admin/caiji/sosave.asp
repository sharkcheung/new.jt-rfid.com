<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="c_func.asp"-->
<%Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
session.CodePage=65001
call openConn()
tit=request("tit")
cont=request("cont")
Module=request("Module")
'response.Write cont
'response.end
set rs=server.CreateObject("adodb.recordset")
sql="select top 1 * from [FK_Article ] where [Fk_Article_Title]='"&tit&"'"
rs.open sql,conn,1,3
if not rs.eof then
   response.Write "此信息已采集过! 请不要重复采集"
else
   'response.Write cont
   rs.addnew
   rs("Fk_Article_Title")=tit
   rs("Fk_Article_Content")=cont
   rs("Fk_Article_From")="互联网"
   rs("Fk_Article_Module")=Module
   rs("Fk_Article_Menu")=1
   rs("Fk_Article_Recommend")=",0,"
   rs("Fk_Article_Subject")=",0,"
   rs.update
   response.Write "采集成功!<font style=""color:#555;font-size:14px;""> (注：采集内容默认不在前台显示)</font><font style=""font-size:14px;line-height:150%;""><br>采集完成后请进入【网站】模块编辑内容、完善关键词后勾选【显示】</font>"
end if   
rs.close
set rs=nothing
call closeConn()
%>
