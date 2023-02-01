<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../caiji2/c_func.asp"-->
<%Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
session.CodePage=65001
Response.Charset = "utf-8"
call openConn()
tit=request("tit")
cont=request("cont")
Module=request("Module")
set rs=server.CreateObject("adodb.recordset")
sql="select top 1 * from [FK_Article ] where [Fk_Article_Title]='"&tit&"'"
rs.open sql,conn,1,3
if not rs.eof then
   response.Write "<b>数据库已存在相同内容！</b><br>请确认是否重复采集？！"
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
   response.Write "<b>采集保存成功！默认该条内容前台不显示。</b><br>需在【网站】对应栏目找到该条内容进行编辑。<br>并完善关键词后勾选【显示】即可！"
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="采集信息：【"&tit&"】"
	log_user=Request.Cookies("FkAdminName")
	
	log_ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If log_ip="" Then log_ip = Request.ServerVariables("REMOTE_ADDR")
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
end if   
rs.close
set rs=nothing
call closeConn()
%>
