<!--#include file="easp.asp"-->
<%dim XmlHttpData,strReferer,tjid,k,result,url
response.Charset="utf-8"
session.CodePage=65001
asp.noCache()
tjid=asp.r("tjid",1)
k=asp.r("k",0)
strReferer="http://tongji2010.qebang.cn/"
url="http://tongji2010.qebang.cn/user/k_visit.asp?tjid="&tjid
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
with objXML
	.open "GET",url,false
	.send(null)
	result=.responseText
	asp.w result
end with
Set objXML=nothing
%>