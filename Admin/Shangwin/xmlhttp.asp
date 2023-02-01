<!--#Include File="../../Inc/Config.asp"-->
<!--#include file="easp.asp"-->
<%Dim k,strReferer,url,objXML,result
response.Charset="utf-8"
session.CodePage=65001
asp.noCache()
tjid=asp.r("tjid",1)
k=asp.r("k",0)
strReferer="http://tongji2010.qebang.cn/"
url="http://tongji2010.qebang.cn/user/k_visit.asp?kw="&asp.escape(k)&"&tjid="&tjid
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
with objXML
	.open "GET",url,false
	.send(null)
	result=.responseText
	asp.w result
end with
Set objXML=Nothing
Call FKDB.DB_Open()
conn.execute("update [keywordSV] set SVci="&result&" where SVkeywords='"&k&"'")
Call FKDB.DB_Close()
%>