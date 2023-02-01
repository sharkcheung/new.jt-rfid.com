<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<%

u_id=M_memberID(session("u_id"))
num=request("buy_num")
pid=Cint(request("cpid"))
'danjia=Cint(request("danjia"))
danjia=request("danjia")
laiurl=Request.ServerVariables("HTTP_REFERER")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Fk_Product where Fk_Product_Id="&pid&"",connn,1,1 
ProductTitle=rs("Fk_Product_Title")
rs.movenext
rs.close
set rs=nothing

'if session("u_id")="" then
'   response.Write "<script language=javascript>alert('请先登录以后再购买!');window.top.location.href='"&laiurl&"';<'/script>"
'   response.End
'end if

if pid="" or pid=0 or not isnumeric(pid) then
   response.Write "<script language=javascript>alert('提交的参数有误！');window.top.location.href='"&laiurl&"';</script>"
   response.end
end if
set payrs=server.CreateObject("adodb.recordset")
if session("u_id")<>"" then

paysql="select * from u_shopcart where u_id="&u_id&" and product_id="&pid&""
payrs.open paysql,connn,3,3
if payrs.eof then
payrs.addnew
payrs("product_id")=pid
payrs("cart_time")=now()
payrs("u_id")=u_id
payrs("u_num")=num
payrs("cart_ip")=getIP()
payrs("danjia")=danjia
payrs("ProductTitle")=ProductTitle
else
payrs("u_num")=payrs("u_num")+num
end if
payrs.update
else
'paysql="select * from u_shopcart"
'payrs.open paysql,connn,3,3
'payrs.addnew
'payrs("product_id")=pid
'payrs("cart_time")=now()
'payrs("u_id")=0
'payrs("u_num")=num
'payrs("cart_ip")=getIP()
'payrs("danjia")=danjia
'payrs("ProductTitle")=ProductTitle
response.redirect "direct_buy.asp?p_id="&pid&"&danjia="&danjia&"&p_name="&ProductTitle&""
response.end
end if
payrs.close
set payrs=nothing
response.redirect "shop_cart.asp"
connn.close
set connn=nothing
%>