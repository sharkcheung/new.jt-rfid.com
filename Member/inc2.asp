<%
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
  ' dbpath=SiteData   '这里修改数据库路径或名称
  ' sql_db_name="qb_kf_sql"
  ' sql_db_id="sa"
   'sql_db_pass="qebangschool"
   'sql_db_ip="(local)"
   'sql_db_ip="202.105.135.57,9098"
comn=""
fgf=""
'QBDll调用
Dim QBDll
Set QBDll = Server.CreateObject("shangwindll.shangwin")
%>
<!--#include file="../inc/conn.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/md5.asp"-->