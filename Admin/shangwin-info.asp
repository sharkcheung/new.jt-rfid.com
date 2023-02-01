<%
Response.CodePage=65001
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
%>
<!--#include file = ../inc/Site.asp -->

<%
infoid=request("infoid")
select case infoid

case "company"
response.write SiteName&""  'infoid=company获取公司名

case "kfid"
response.write kfid&""   'infoid=kfid获取客服账号

case "tjid"
response.write tjid&""   'infoid=tjid获取统计id

end select
%>