<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
   response.Write "<script language=javascript>alert('гКох╣гб╪!');window.location.href='./';</script>"
   response.End
end if
connn.execute("delete from u_shopcart where u_id="&M_memberID(session("u_id"))&"")
response.Redirect "shop_cart.asp"
response.end
%>
