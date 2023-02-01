<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
   response.write "<script language=javascript>alert('您还未登录,请先登录再查看购物车!');window.top.location.href='../';</script>"
   response.End
end if%>

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title><%=hometit%>-<%=company%></title>
<meta name="keywords" content="<%=keywords%>" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body oncontextmenu="return false" style="padding:5px;">
<div id="main" style="height:450px;">
<table width=100% align=center border=0 cellspacing=0 cellpadding=0 class=table-zuoyou bordercolor=#CCCCCC><tr><td height=38 class=table-shangxia>　<img src=images/ring02.gif align=absmiddle>
<%if session("u_id")<>"" then
username=M_memberID(session("u_id"))
end if
response.write " <a href=index.asp>"&webname&"</a>"
%>
我的购物车</td></tr></table>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_shopcart where u_id="&username&"",connn,1,1 
%>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC" style="border:solid #CCCCCC 1px;">
          <tr>
            <td bgColor=#ffffff colSpan=5 height=1></TD>
          </tr>
          <form name='form1' method='post' action=xgsl.asp>
            <tr bgcolor=#f1f1f1 align=center>
              <td width=36% height="30">商品名称</td>
              <td width=12%>单价(元)              </td>
              <td width=19%>数量</td>
              <td width=19%>小计</td>
              <td width=14%>删除</td>
            </tr>
            <tr>
              <td bgColor=#cccccc colSpan=5 height=1></TD>
            </tr>
            
            <%shuliang=rs.recordcount
jianshu=0
zongji=0
do while not rs.eof
p_id=int(rs("product_id"))
cart_id=rs("id")
num=rs("u_num")
if num="" then num=1
nu=nu+num
d_price=rs("danjia")

t_p=t_p+(num*d_price-0)
%>
            <tr bgcolor="#FFFFFF">
              <td height="30" width="36%" class="table-xia">　<%=rs("ProductTitle")%></a>
                  <input name=bookid type=hidden value="<%=rs("product_id")%>">
                  <input name=actionid type=hidden value="<%=rs("id")%>">
              </td>
              <td width="12%" height="30" align="center" class="table-xia">
        <%=d_price%> 元</td>
              <td align="center" width="19%" height="30" class="table-xia"><input type='text' id='btn_cha_<%=num%>' name='bookcount' maxlength='4' style='width:30px' onKeyDown='if(event.keyCode == 13) event.returnValue = false' value='<%=num%>' /><input type='hidden' name='hidChange<%=p_id%>' value='1' />
              </td>
              <td align="center" width="19%" height="30" class="table-xia"><font color=red><%=d_price*num%></font> 元</td>
              <td align="center" width="14%" height="30" class="table-xia"><a href=?action=del&actionid=<%=cart_id%>><img src=images/del_cart.gif width="16" height="16" border=0 title="点击删除该商品"></a> </td>
            </tr>
            <%
rs.movenext
loop
rs.close
set rs=nothing%>
            <tr bgcolor="#FFFFFF">
              <td height=30 colspan=5 align=center class=table-xia> 购物车里有商品：<font color=red></font><%if session("bookc")<>"" then%><%=session("bookc")%><%else%><%=nu%><%end if%> 件　总数：<font color=red><%if session("bookc")<>"" then%><%=session("bookc")%><%else%><%=nu%><%end if%></font> 件　共计：<span class='price' id='cartBottom_price'></span><%if session("pri")<>"" then%><%=session("pri")%><%else%><%=t_p%><%end if%> 元</b></span>　
              </td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center" height=50 colspan=5><input name="imageField" type="image" src="images/cart01.gif" width="115" height="36" border="0" style="border:0px;" onFocus="this.blur()" onClick="this.form.action='#';this.form.submit();window.top.location.href='/';">
                  <input name="imageField2" style="border:0px;" type="image" src="images/cart03.gif" width="115" height="36" border="0" onFocus="this.blur()" onClick="this.form.action='xgsl.asp';this.form.submit()">
                  <input name="imageField22" style="border:0px;" type="image" src="images/cart02.gif" width="115" height="36" border="0" onFocus="this.blur()" onClick="this.form.action='qkgwc.asp';this.form.submit()">
                  <input name="imageField222" style="border:0px;" type="image" src="images/cart04.gif" width="115" height="36" border="0" onFocus="this.blur()" onClick="this.form.action='Iheeo_car.asp';this.form.submit()">
              </td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td height="60" colspan="5" STYLE="PADDING-LEFT: 20px">
			  <ul>
		         <li>如果您想继续购物，请点选继续购物</li>
                 <li>如果您想更新已在购物车内的产品，请先修改，然后点选修改数量</li>
                 <li>如果您想全部取消已订购在购物车中的产品，请点选清空购物车</li>
                 <li>如果您满意您所购买的产品，请点选去收银台(会员须先登陆，非会员须先免费注册成为会员)</li>
			  </ul>
              </td>
            </tr>
          </form>
      </table>
</div>
<%
dim bookid,username,action
action=request.QueryString("action")
username=session("u_id")
bookid=request.QueryString("id")

if InStr(action,"'")>0 then
response.write"<script>alert(""非法访问!"");window.close();</script>"
response.end
end if

if bookid<>"" then
if not isnumeric(bookid) then 
response.write"<script>alert(""非法访问!"");window.close();</script>"
response.end
else
if not isinteger(bookid) then
response.write"<script>alert(""非法访问!"");window.close();</script>"
end if
end if
end if
'//删除购物车
select case action
case "del"
connn.execute "delete from u_shopcart where id="&request.QueryString("actionid")
response.redirect "shop_cart.asp"
response.End
case "add"
'//商品，判断是否存在
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from Fk_Product where Fk_Product_Id="&bookid,connn,1,1
if request.Cookies("bjx")("reglx")="2" then 
	danjia=rs_s("vipjia")
else
	danjia=rs_s("huiyuanjia")
end if
kucun=rs_s("kucun")
bookname=rs_s("bookname")
shjiaid=rs_s("shjiaid")
rs_s.close
set rs_s=nothing
if kucun<=0 then
response.write "<script language=javascript>alert('你选购的商品“"&bookname&"”暂时缺货不能放到购物车里，请选购其它商品！');window.close();</script>"
response.end
end if
'set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from u_order where u_id="&M_memberID(username)&" and p_id="&bookid&" and zhuangtai=7",connn,1,3
'if rs.recordcount=1 then
'if kucun<(rs("bookcount")+1) then
'response.write "<script language=javascript>alert('你选购的商品“"&bookname&"”暂时缺货不能放到购物车里，请选购其它商品！');window.close();<'/script>"
'response.end
'end if
'rs("zonger")=(rs("bookcount")+1)*danjia
'rs("bookcount")=rs("bookcount")+1
'rs.update
'rs.close
'set rs=nothing
'response.Redirect "buy.asp?action=show"
'else
'//添加购物
'rs.close
'set rs=server.CreateObject("adodb.recordset")
'rs.open "select bookid,username,shjiaid,zhuangtai,zonger,bookcount,niming from BJX_action",connn,1,3
'rs.addnew
'rs("bookid")=bookid
'rs("username")=username
'rs("zhuangtai")=7
'rs("bookcount")=1
'rs("shjiaid")=shjiaid
'rs("zonger")=danjia
'if request.Cookies("bjx")("username")="" then
'rs("niming")=1
'end if
'rs.update
'rs.close
'set rs=nothing
response.Redirect "buy.asp?action=show"
'end if
end select%>
</body>
</html>