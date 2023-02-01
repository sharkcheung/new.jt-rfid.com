<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
response.write "<script language=javascript>window.location.href='../';</script>"
response.End
end if
actionid=request("actionid")
bookid=request("bookid")
bookcount=request("bookcount")
if actionid="" then
response.write "<script language=javascript>alert('对不起，您没有选择商品！');history.back();</script>"
response.End
end if
act_id=split(actionid,",")
b_id=split(bookid,",")
n_id=split(bookcount,",")
for i=0 to ubound(act_id)
if b_id(i)<=0 then
bookcount=1
else
bookcount=n_id(i)
if not isnumeric(bookcount) then
   response.write "<script language=javascript>alert('请输入正确的数量!');history.back();</script>"
   response.end
else
   if bookcount<=0 then
   response.write "<script language=javascript>alert('请输入正确的数量!');history.back();</script>"
   response.end
   end if
end if
end if

connn.execute("update u_shopcart set u_num="&bookcount&" where id="&act_id(i))
next
response.Redirect "shop_cart.asp"
%>
